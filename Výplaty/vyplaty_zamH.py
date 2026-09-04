"""
Analýza meziročních a meziměsíčních změn mezd zaměstnanců.

Skript:

- načítá data ze všech relevantních listů Excelu,
- kontroluje přítomnost povinných sloupců,
- odstraňuje záznamy s neplatnými mzdami,
- kontroluje kompletnost sledovaných období,
- počítá relativní změny mzdy,
- identifikuje zaměstnance v horním a dolním percentilu změn,
- vytváří výstupní Excel s jednotlivými přehledy,
- generuje boxplot relativních změn mezd.

Vstupvy
-----
Excelový soubor obsahující měsíční přehledy mezd zaměstnanců.

Výstup
------
Excelový soubor RESULT_<název_souboru>.xlsx obsahující:

- Rel_zmena; - Top5_pct; 
- WARNING_MZDY0; - WARNING-kontrola kompletnosti
- Obd_*
- INC_*; - DEC_*

Dále je vytvořen obrázek:

- boxplot_<název_souboru>.png
"""

from __future__ import annotations
from dataclasses import dataclass
import logging
from pathlib import Path
from typing import Final

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ================================

# Architektura               9,9/10
# Čitelnost                  9,9/10
# Dokumentace                9,9/10
# Logging                    9,8/10
# Python styl                9,7/10
# Robustnost                 9,7/10
# Udržovatelnost             10/10


# ==================================================
# KONSTANTY
# ==================================================

@dataclass(frozen=True)
class Config:
    """
    Konfigurační parametry analýzy.

    Obsahuje:

    - cesty ke vstupním a výstupním souborům,
    - názvy používaných sloupců,
    - kontrolovaná období,
    - období pro tvorbu boxplotu,
    - percentily používané pro identifikaci
      extrémních nárůstů a poklesů mezd.
    """

    soubor: str
    soubor_odejiti : str = "Výplaty/odejiti.xlsx"

    sloupec_id: str = "oscis"
    sloupec_obdobi : str = "obdobi"
    sloupec_hmzda : str = "hmzda"
    sloupec_hodnoty: str = "hmzda_odmen"
    sloupec_vedouci: str = "appvdr"

    obdobi_boxplot: tuple[str, ...] = (
        "05.26", "06.26", "07.26",
    )

    kontrolovana_obdobi: tuple[str, ...] = (
        "04.26", "05.26", "06.26",
    )

    povinne_sloupce: tuple[str, ...] = (
        "oscis", "obdobi", "pracvz",
        "prijm", "hmzda", "hmzda_odmen",
        "appvdr",
    )

    vystupni_sloupce: tuple[str, ...] = (
    "oscis", "obdobi", "pracvz",
    "prijm", "hmzda", "hmzda_odmen",
    "appvdr", "Rel_zmena", "Rel_zmena_%",
    )

    top_percentil: float = 0.95
    bottom_percentil: float = 0.05

    @property
    def vstup(self) -> Path:
        """Vrátí cestu ke vstupnímu souboru."""
        return Path(self.soubor)

    @property

    def soubor_vystup(self) -> Path:
        """Vrátí cestu k výstupnímu souboru."""
        return (
            self.vstup.parent /
            f"RESULT_{self.vstup.name}"
        )

    @property
    def vstup_odejiti(self) -> Path:
        """Vrátí cestu k souboru odešlých zaměstnanců."""
        return Path(self.soubor_odejiti)

    def __post_init__(self) -> None:

        if not 0 < self.bottom_percentil < 1:
            raise ValueError(
                "bottom_percentil musí být mezi 0 a 1."
            )

        if not 0 < self.top_percentil < 1:
            raise ValueError(
                "top_percentil musí být mezi 0 a 1."
            )

        if self.bottom_percentil >= self.top_percentil:
            raise ValueError(
                "bottom_percentil musí být menší než top_percentil."
            )


CONFIG = Config(
    soubor="Výplaty/vyplaty_zam4.xlsx"
)


IGNOROVANE_LISTY: Final[set[str]] = {
    "Rel_zmena","Top5_pct", "WARNING_MZDY0", "WARNING-kontrola kompletnosti", "WARNING-pracvz"
}


# ==================================================
# LOGGING
# ==================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

logger = logging.getLogger(__name__)


# ==================================================
# NAČTENÍ DAT
# ==================================================    

def nacti_data(soubor: str) -> pd.DataFrame:

    """
    Načte všechny relevantní listy ze vstupního Excelu.

    Parametry
    ---------
    soubor : str
        Cesta ke vstupnímu Excel souboru.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Sloučená data ze všech zpracovávaných listů.

    Výjimky
    --------
    ValueError
        Pokud nejsou nalezena žádná vstupní data.
    """

    vstup = Path(soubor)

    if not vstup.exists():
        raise FileNotFoundError(
            f"Vstupní soubor neexistuje: {vstup}"
            )

    xls = pd.ExcelFile(soubor)
    
    data: list[pd.DataFrame] = []

    for sheet in xls.sheet_names:

        if sheet.startswith(("INC_", "DEC_", "Obd_")):
            continue

        if sheet in IGNOROVANE_LISTY:
            continue

        logger.info("Načítám list %s", sheet)

        data.append(
            pd.read_excel(
                xls, sheet_name=sheet,
            )
        )

    if not data:
        raise ValueError(
            "Nebyla nalezena žádná vstupní data."
        )

    return pd.concat(
        data, ignore_index=True,
    )

def nacti_odejite(
    config: Config,
) -> pd.DataFrame:

    if not config.vstup_odejiti.exists():

        logger.warning(
            "Soubor odejiti.xlsx nebyl nalezen."
        )

        return pd.DataFrame()

    return pd.read_excel(
        config.vstup_odejiti
    )

# ==================================================
# VALIDACE
# ==================================================

def validuj_sloupce(
    df: pd.DataFrame, povinne_sloupce: tuple[str, ...],
) -> None:

    """
    Ověří přítomnost všech povinných sloupců.

    Parametry
    ---------
    df : pd.DataFrame
        Kontrolovaný datový rámec.

    povinne_sloupce : list[str]
        Seznam sloupců, které musí být přítomny.

    Výjimky
    --------
    ValueError
        Pokud některý z povinných sloupců chybí.
    """

    chybi = set(povinne_sloupce) - set(df.columns)

    if chybi:
        raise ValueError(
        f"Chybí povinné sloupce: {sorted(chybi)}"
    )


    
def kontrola_pracvz(
    df: pd.DataFrame, config: Config,
) -> pd.DataFrame:
    """
    Najde zaměstnance, kteří mají ve sledovaných
    obdobích více různých pracovních vztahů.
    """

    kontrolovana = df[
        df[config.sloupec_obdobi]
        .isin(config.kontrolovana_obdobi)
    ].copy()

    pracvz_count = (
        kontrolovana
        .groupby(config.sloupec_id)["pracvz"]
        .nunique()
    )

    problematicti = set(
        pracvz_count[pracvz_count > 1].index
    )

    if not problematicti:
        return pd.DataFrame()

    warning_df = (
        kontrolovana[
            kontrolovana[config.sloupec_id]
            .isin(problematicti)
        ]
        .copy()
    )

    warning_df["POCET_PRACVZ"] = (
        warning_df[config.sloupec_id]
        .map(pracvz_count)
    )

    warning_df["SEZNAM_PRACVZ"] = (
        warning_df
        .groupby(config.sloupec_id)["pracvz"]
        .transform(
            lambda x: ", ".join(
                sorted(
                    map(str, x.dropna().unique())
                )
            )
        )
    )

    logger.warning(
        "Nalezeno %s zaměstnanců s více pracvz.",
        len(problematicti),
    )

    return warning_df.sort_values(
        [config.sloupec_id, "Obdobi_sort"]
    )

# ==================================================
# OBDOBÍ
# ==================================================

def priprav_obdobi(
    df: pd.DataFrame, config: Config
) -> pd.DataFrame:

    """
    Převede hodnoty období do jednotného formátu.

    Funkce sjednotí hodnoty sloupce 'obdobi'
    na formát MM.YY a vytvoří pomocný sloupec
    'Obdobi_sort' používaný pro chronologické
    řazení.

    Parametry
    ---------
    df : pd.DataFrame
        Vstupní data.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Data s doplněným sloupcem Obdobi_sort.
    """
    df = df.copy()

    if pd.api.types.is_datetime64_any_dtype(
        df[config.sloupec_obdobi]
    ):
        df[config.sloupec_obdobi] = (
            pd.to_datetime(df[config.sloupec_obdobi])
            .dt.strftime("%m.%y")
        )
    else:
        df[config.sloupec_obdobi] = (
            df[config.sloupec_obdobi]
            .astype(str)
            .str.strip()
        )

    df["Obdobi_sort"] = pd.to_datetime(
        df[config.sloupec_obdobi], format="%m.%y", errors="coerce",
        )

    pred = len(df)

    df = df.dropna(
        subset=["Obdobi_sort"]
        )

    pocet_vyrazenych = pred - len(df)

    if pocet_vyrazenych > 0:
        logger.warning(
            "Vyřazeno %s řádků s neplatným obdobím.",
            pocet_vyrazenych,)

    else:
        logger.info(
            "Všechna období jsou validní."
            )

    return df


# ==================================================
# NULOVÉ MZDY
# ==================================================

def odstran_nulove_mzdy(
    df: pd.DataFrame, config: Config
) -> tuple[pd.DataFrame, pd.DataFrame]:

    """
    Vyřadí záznamy s neplatnými mzdami.

    Za neplatné jsou považovány hodnoty:

    - NULL
    - NaN
    - 0

    ve sloupcích hmzda nebo hmzda_odmen.

    Parametry
    ---------
    df : pd.DataFrame
        Vstupní data.

    Návratová hodnota
    -----------------
    tuple[pd.DataFrame, pd.DataFrame]

    První položka:
        Očištěná data.

    Druhá položka:
        Vyřazené záznamy určené pro export
        do listu WARNING_MZDY0.
    """

    warning_df = df[
        (df[config.sloupec_hmzda].isna())
        | (df[config.sloupec_hmzda] == 0)
        | (df[config.sloupec_hodnoty].isna())
        | (df[config.sloupec_hodnoty] == 0)
    ].copy()

    if not warning_df.empty:
        logger.warning(
            "Nalezeno %s neplatných záznamů.",
            len(warning_df),
        )

    df_clean = df[
        (df[config.sloupec_hmzda].notna())
        & (df[config.sloupec_hmzda].ne(0))
        & (df[config.sloupec_hodnoty].notna())
        & (df[config.sloupec_hodnoty] != 0)
    ].copy()

    return df_clean, warning_df


# ==================================================
# KOMPLETNOST OBDOBÍ
# ==================================================

def kontrola_kompletnosti(
    df: pd.DataFrame, config: Config,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Ověří kompletnost sledovaných období.

    Každý zaměstnanec musí mít přítomna
    všechna období definovaná v konfiguraci.
    """

    pozadovana_obdobi = set(
        config.kontrolovana_obdobi
    )

    skutecna_obdobi = (
        df.groupby(config.sloupec_id)[config.sloupec_obdobi]
        .agg(set)
    )

    nekompletni_set = set(
        skutecna_obdobi[
            skutecna_obdobi != pozadovana_obdobi
        ].index
    )

    if not nekompletni_set:
        return df, pd.DataFrame()

    warning_df = (
        df[
            df[config.sloupec_id]
            .isin(nekompletni_set)
        ]
        .copy()
    )

    chybejici_map: dict[str, str] = {}
    navic_map: dict[str, str] = {}

    for zam in nekompletni_set:

        obdobi = skutecna_obdobi[zam]

        chybejici_map[zam] = ", ".join(
            sorted(
                pozadovana_obdobi - obdobi
            )
        )

        navic_map[zam] = ", ".join(
            sorted(
                obdobi - pozadovana_obdobi
            )
        )

    warning_df["CHYBEJICI_OBDOBI"] = (
        warning_df[config.sloupec_id]
        .map(chybejici_map)
    )

    warning_df["NAVIC_OBDOBI"] = (
        warning_df[config.sloupec_id]
        .map(navic_map)
    )

    logger.warning(
        "Nalezeno %s zaměstnanců s nekompletní historií.",
        len(nekompletni_set),
    )

    df_clean = (
        df[
            ~df[config.sloupec_id]
            .isin(nekompletni_set)
        ]
        .copy()
    )

    return df_clean, warning_df


# ==================================================
# TOP PERCENTILY
# ==================================================

def vytvor_top_percentily(
    df: pd.DataFrame, config: Config,
) -> pd.DataFrame:

    """
    Identifikuje extrémní nárůsty a poklesy mezd.

    Pro každé období jsou určeny:

    - horní percentil růstů,
    - dolní percentil poklesů.

    Parametry
    ---------
    df : pd.DataFrame
        Data obsahující vypočtené relativní změny.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Záznamy patřící do skupin:

        - top5_pct_INCREASE
        - top5_pct_DECREASE
    """

    vysledky: list[pd.DataFrame] = []

    for obdobi, skupina in df.groupby(config.sloupec_obdobi):

        data_mesic = skupina.dropna(
            subset=["Rel_zmena_%"]
        )

        if len(data_mesic) < 5:
            continue

        hranice_95 = (
            data_mesic["Rel_zmena_%"]
            .quantile(config.top_percentil)
        )

        hranice_05 = (
            data_mesic["Rel_zmena_%"]
            .quantile(config.bottom_percentil)
        )

        top_increase = (
            data_mesic[
                data_mesic["Rel_zmena_%"] >= hranice_95
            ]
            .assign(
                typ="top5_pct_INCREASE"
            )
        )

        top_decrease = (
            data_mesic[
                data_mesic["Rel_zmena_%"] <= hranice_05
            ]
            .assign(
                typ="top5_pct_DECREASE"
            )
        )

        vysledky.extend(
            [top_increase, top_decrease]
        )

    if not vysledky:
        return pd.DataFrame()

    return pd.concat(
        vysledky, ignore_index=True,
    )

def zapis_rel_zmena(
    writer: pd.ExcelWriter, df_export: pd.DataFrame,
) -> None:  
                
    """
    Zapíše hlavní list Rel_zmena.

    Parametry
    ---------
    writer : pd.ExcelWriter
        Otevřený ExcelWriter.

    df_export : pd.DataFrame
        Data s vypočtenými relativními změnami.

    Návratová hodnota
    -----------------
    None
    """

    df_export.to_excel(
        writer, sheet_name="Rel_zmena", index=False,
    )


def zapis_warning_mzdy(
    writer: pd.ExcelWriter, warning_mzdy: pd.DataFrame, config: Config
) -> None:

    """
    Zapíše list WARNING_MZDY0.

    Obsahuje záznamy vyřazené z důvodu
    nulové nebo chybějící mzdy.

    Parametry
    ---------
    writer : pd.ExcelWriter
        Otevřený ExcelWriter.

    warning_mzdy : pd.DataFrame
        Vyřazené záznamy.

    Návratová hodnota
    -----------------
    None
    """

    if warning_mzdy.empty:
        return

    warning_export = warning_mzdy.copy()

    warning_export[config.sloupec_obdobi] = (
        pd.to_datetime(
            warning_export[config.sloupec_obdobi],
            errors="coerce",
        )
        .dt.strftime("%m.%y")
    )

    warning_sloupce = [
        config.sloupec_id, config.sloupec_obdobi, "pracvz",
        "prijm", config.sloupec_hmzda, config.sloupec_hodnoty,
        config.sloupec_vedouci,
    ]

    warning_export[
        warning_sloupce
    ].to_excel(
        writer,
        sheet_name="WARNING_MZDY0",
        index=False,
    )

def zapis_warning_kompletnost(
    writer: pd.ExcelWriter, warning_kompletnost: pd.DataFrame, config: Config,
) -> None:

    """
    Zapíše list WARNING-kontrola kompletnosti.

    Obsahuje zaměstnance, kteří nemají
    kompletní historii ve sledovaných
    obdobích.

    Parametry
    ---------
    writer : pd.ExcelWriter
        Otevřený ExcelWriter.

    warning_kompletnost : pd.DataFrame
        Vyřazené záznamy.

    Návratová hodnota
    -----------------
    None
    """

    if warning_kompletnost.empty:
        return

    export_kompletnost = (
        warning_kompletnost
        .drop(
            columns=["Obdobi_sort"],
            errors="ignore",
        )
        .sort_values(
            [config.sloupec_id, config.sloupec_obdobi]
        )
    )

    export_kompletnost.to_excel(
        writer,
        sheet_name=(
            "WARNING-kontrola kompletnosti"
        ),
        index=False,
    )

def kontrola_odmen(df: pd.DataFrame, config: Config) -> None:
    """
    Zkontroluje konzistenci agregovaných a detailních složek odměn.

    Kontroly:
    - odmm == odm2
    - odmn == odm6
    - odmo == odm10

    Rozdíly vypisuje do konzole.
    """

    kontroly = [
        ("odmm", "odm2"),
        ("odmn", "odm6"),
        ("odmo", "odm10"),
    ]

    for souhrnny, detailni in kontroly:

        if souhrnny not in df.columns:
            logger.warning("Chybí sloupec %s", souhrnny)
            continue

        if detailni not in df.columns:
            logger.warning("Chybí sloupec %s", detailni)
            continue

        rozdily = df[
            df[souhrnny].fillna(0)
            != df[detailni].fillna(0)
        ]

        if not rozdily.empty:
            logger.warning(
                "%s vs %s: nalezeno %s rozdílů.",
                souhrnny, detailni, len(rozdily)
            )
            logger.warning(
                "\n%s",rozdily[[config.sloupec_id, config.sloupec_obdobi,
                                souhrnny, detailni,]
                                ].head(20)
                            )

        else:
            logger.info("%s vs %s: bez rozdílů.",
                        souhrnny, detailni,
                        )


def zapis_top5(
    writer: pd.ExcelWriter, df_top: pd.DataFrame, vystupni_sloupce: list[str],
) -> None:

    """
    Zapíše list Top5_pct.

    Obsahuje zaměstnance patřící
    do horního a dolního percentilu
    relativních změn.

    Parametry
    ---------
    writer : pd.ExcelWriter
        Otevřený ExcelWriter.

    df_top : pd.DataFrame
        Záznamy extrémních změn.

    vystupni_sloupce : list[str]
        Seznam exportovaných sloupců.

    Návratová hodnota
    -----------------
    None
    """

    if df_top.empty:

        pd.DataFrame(
            columns=vystupni_sloupce
        ).to_excel(
            writer, sheet_name="Top5_pct", index=False,
        )

        return

    (
        df_top
        .drop(columns=["typ"])
        .sort_values(
            "Rel_zmena_%",
            ascending=False,
        )
        .to_excel(
            writer, sheet_name="Top5_pct", index=False,
        )
    )

def zapis_obdobi_listy(
    writer: pd.ExcelWriter, df_export: pd.DataFrame, config: Config,
) -> None:

    """
    Vytvoří samostatné listy Obd_*.

    Každé období je exportováno
    do samostatného listu.

    Parametry
    ---------
    writer : pd.ExcelWriter
        Otevřený ExcelWriter.

    df_export : pd.DataFrame
        Data určená k exportu.

    Návratová hodnota
    -----------------
    None
    """

    for obdobi in sorted(
        df_export[config.sloupec_obdobi].unique()
    ):

        sheet_name = (
            f"Obd_{obdobi.replace('.', '_')}"
        )[:31]

        (
            df_export[
                df_export[config.sloupec_obdobi] == obdobi
            ]
            .sort_values(
                ["prijm", config.sloupec_id]
            )
            .to_excel(
                writer, sheet_name=sheet_name, index=False,
            )
        )

def priprav_inc_dec_export(
    df: pd.DataFrame, config: Config, vzestupne: bool,
) -> pd.DataFrame:
    """
    Připraví data pro export INC_* nebo DEC_* listu.
    """

    sloupce_k_odstraneni = (
        ["odmm", "odmn", "odmo", "odmr"]
        + [f"odm{i}" for i in range(1, 21)]
    )

    poradi_sloupcu = [
        config.sloupec_id, config.sloupec_obdobi, "pracvz",
        "prijm", "hmzda_predchozi", config.sloupec_hmzda,
        "hmzda_odmen_predchozi", config.sloupec_hodnoty, config.sloupec_vedouci,
        "Rel_zmena", "Rel_zmena_%",
    ]

    vysledek = (
        df
        .drop(
            columns=["typ", *sloupce_k_odstraneni],
            errors="ignore",
        )
        .sort_values(
            "Rel_zmena_%",
            ascending=vzestupne,
        )
    )

    return vysledek.reindex(
        columns=[
            c
            for c in poradi_sloupcu
            if c in vysledek.columns
        ]
    )

def zapis_inc_dec_listy(
    writer: pd.ExcelWriter, df_top: pd.DataFrame, config: Config,
) -> None:
    """
    Vytvoří listy INC_* a DEC_*.

    Pro každé období vytvoří seznam
    největších růstů a poklesů mezd.

    Parametry
    ---------
    writer : pd.ExcelWriter
        Otevřený ExcelWriter.

    df_top : pd.DataFrame
        Data extrémních změn.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    None
    """

    if df_top.empty:
        return

    for obdobi, data_obdobi in sorted(
        df_top.groupby(config.sloupec_obdobi),
        key=lambda x: x[0],
    ):

        inc = priprav_inc_dec_export(
            data_obdobi[
                data_obdobi["typ"]
                == "top5_pct_INCREASE"
            ],
            config, vzestupne=False,
        )

        dec = priprav_inc_dec_export(
            data_obdobi[
                data_obdobi["typ"]
                == "top5_pct_DECREASE"
            ],
            config, vzestupne=True,
        )

        inc.to_excel(
            writer,
            sheet_name=(
                f"INC_{obdobi.replace('.', '_')}"
            )[:31],
            index=False,
        )

        dec.to_excel(
            writer,
            sheet_name=(
                f"DEC_{obdobi.replace('.', '_')}"
            )[:31],
            index=False,
        )

def dopln_predchozi_hodnoty(
    df: pd.DataFrame, sloupec_id: str, sloupec_mzda: str,
    sloupec_hodnoty: str,
) -> pd.DataFrame:
    """
    Doplní předchozí hodnoty mzdy a mzdy včetně odměn.
    """

    df = df.copy()

    df["hmzda_predchozi"] = (
        df.groupby(sloupec_id)[sloupec_mzda]
        .shift(1)
    )

    df["hmzda_odmen_predchozi"] = (
        df.groupby(sloupec_id)[sloupec_hodnoty]
        .shift(1)
    )

    return df

def dopln_relativni_zmenu(
    df: pd.DataFrame, sloupec_id: str, sloupec_hodnoty: str,
) -> pd.DataFrame:
    """
    Doplní relativní změnu sledované hodnoty.
    """

    df = df.copy()

    df["Rel_zmena"] = (
        (
            df["hmzda_odmen"]
            - df["hmzda_odmen_predchozi"]
        )
        / df["hmzda_odmen_predchozi"]
    )

    df["Rel_zmena"] = (
        df["Rel_zmena"]
        .replace([np.inf, -np.inf], np.nan)
    )

    df["Rel_zmena_%"] = (
        df["Rel_zmena"] * 100
    ).round(2)

    return df

# ==================================================
# RELATIVNÍ ZMĚNA
# ==================================================

def vypocitej_relativni_zmenu(
    df: pd.DataFrame, config: Config
) -> pd.DataFrame:

    """
    Vypočte relativní změnu mzdy mezi obdobími.

    Výpočet probíhá samostatně pro každého
    zaměstnance na základě chronologického
    pořadí období.

    Parametry
    ---------
    df : pd.DataFrame
        Vstupní data.

    sloupec_id : str
        Identifikátor zaměstnance.

    sloupec_hodnoty : str
        Sloupec obsahující porovnávanou mzdu.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Data doplněná o sloupce:

        - Rel_zmena
        - Rel_zmena_%
    """

    df = (
        df.copy()
        .sort_values([config.sloupec_id, "Obdobi_sort"])
    )

    df = dopln_predchozi_hodnoty(
        df, config.sloupec_id, config.sloupec_hmzda, config.sloupec_hodnoty,
    )

    df = dopln_relativni_zmenu(
        df, config.sloupec_id, config.sloupec_hodnoty,
    )

    return df
                
def export_excel(
    df_export: pd.DataFrame, df_top: pd.DataFrame, warning_mzdy: pd.DataFrame,
    warning_kompletnost: pd.DataFrame, warning_pracvz: pd.DataFrame, df_odejiti: pd.DataFrame, 
    config: Config,
) -> None:

    """
    Vytvoří výsledný Excelový soubor.

    Exportuje:

    - Rel_zmena
    - WARNING_MZDY0
    - WARNING-kontrola kompletnosti
    - Top5_pct
    - Obd_*
    - INC_*
    - DEC_*

    Parametry
    ---------
    df_export : pd.DataFrame
        Hlavní dataset s relativními změnami.

    df_top : pd.DataFrame
        Zaměstnanci v extrémních percentilech.

    warning_mzdy : pd.DataFrame
        Vyřazené záznamy s neplatnými mzdami.

    warning_kompletnost : pd.DataFrame
        Vyřazené záznamy s neúplnou historií.

    Návratová hodnota
    -----------------
    None
    """
    try:
        with pd.ExcelWriter(
            config.soubor_vystup, engine="openpyxl",
        ) as writer:

            zapis_rel_zmena(
                writer, df_export,
            )

            zapis_warning_mzdy(
                writer, warning_mzdy, config
            )

            zapis_warning_kompletnost(
                writer, warning_kompletnost, config,
            )

            zapis_warning_pracvz(
                writer,warning_pracvz,
            )

            zapis_top5(
                writer, df_top, config.vystupni_sloupce,
            )

            zapis_obdobi_listy(
                writer, df_export, config,

            )

            zapis_inc_dec_listy(
                writer, df_top, config
            )

            if not df_odejiti.empty:
             

                df_odejiti.to_excel(
                    writer, sheet_name="odejiti", index=False,
                    )

    except PermissionError:
        logger.error(
            "Výstupní soubor je otevřen v Excelu."
            )
        raise

    logger.info(
        "Výstupní soubor vytvořen: %s",
        config.soubor_vystup,
    )

def priprav_warning_odchody(
    warning_kompletnost: pd.DataFrame, config: Config,
) -> pd.DataFrame:
    
    """
    Připraví kandidáty z upozornění na nekompletní historii.

    Záznamy jsou seřazeny chronologicky, doplněny
    o předchozí hodnoty mezd a relativní změny.
    Pro každého zaměstnance je následně ponechán
    poslední dostupný záznam.

    Parametry
    ---------
    warning_kompletnost : pd.DataFrame
        Záznamy z kontroly kompletnosti období.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Připravení kandidáti s označením původu
        „WARNING-kontrola kompletnosti“.
    """

    if warning_kompletnost.empty:
        return pd.DataFrame()

    warning_df = (
        warning_kompletnost
        .sort_values(
            [config.sloupec_id, "Obdobi_sort"]
        )
        .copy()
    )

    warning_df = dopln_predchozi_hodnoty(
        warning_df, config.sloupec_id, config.sloupec_hmzda,
        config.sloupec_hodnoty,
    )

    warning_df = dopln_relativni_zmenu(
        warning_df, config.sloupec_id, config.sloupec_hodnoty,
    )

    warning_df["puvod"] = ("WARNING-kontrola kompletnosti")

    warning_df = (
        warning_df
        .groupby(
            config.sloupec_id,
            as_index=False,
        )
        .last()
    )

    return warning_df

def priprav_inc_odchody(
    df_top: pd.DataFrame, config: Config,
) -> pd.DataFrame:
    """
    Připraví kandidáty z přehledů výrazných nárůstů mezd.

    Vybere zaměstnance zařazené mezi nejvyšší
    relativní růsty a doplní informaci o zdrojovém
    listu INC_*.

    Parametry
    ---------
    df_top : pd.DataFrame
        Data extrémních percentilů.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Kandidáti označení názvem zdrojového
        listu INC_*.
    """

    inc_df = (
    df_top[
        df_top["typ"]
        == "top5_pct_INCREASE"
    ]
    .copy()
    )

    if not inc_df.empty:
        inc_df["puvod"] = (
            "INC_"
            + inc_df[config.sloupec_obdobi]
            .str.replace(".", "_", regex=False)
        )

    return inc_df

def filtruj_odejite(
    kandidati: pd.DataFrame, odejiti: pd.DataFrame, config: Config,
) -> pd.DataFrame:
    """
    Ponechá pouze zaměstnance evidované jako odešlé.

    Parametry
    ---------
    kandidati : pd.DataFrame
        Kandidáti získaní z jednotlivých zdrojů.

    odejiti : pd.DataFrame
        Seznam odešlých zaměstnanců.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Kandidáti nacházející se v seznamu
        odešlých zaměstnanců.
    """

    odejiti_set = set(
    odejiti[config.sloupec_id]
    .dropna()
    )

    return kandidati[
        kandidati[config.sloupec_id]
        .isin(odejiti_set)
    ].copy()

def uprav_vystup_odchody(
    df: pd.DataFrame, config: Config,
) -> pd.DataFrame:
    """
    Uspořádá sloupce výsledného přehledu odešlých.

    Zachová pouze definované sloupce a nastaví
    jejich požadované pořadí pro export do Excelu.

    Parametry
    ---------
    df : pd.DataFrame
        Data určená k exportu.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Data s uspořádanými sloupci.
    """

    poradi_sloupcu = [
        config.sloupec_id, config.sloupec_obdobi,"pracvz",
        "prijm", "hmzda_predchozi", config.sloupec_hmzda,
        "hmzda_odmen_predchozi", config.sloupec_hodnoty, config.sloupec_vedouci,
        "Rel_zmena","Rel_zmena_%", "puvod",
        "CHYBEJICI_OBDOBI", "NAVIC_OBDOBI",
    ]

    return df.reindex(
        columns=[
            c
            for c in poradi_sloupcu
            if c in df.columns
        ]
    )


def vytvor_odejiti(
    warning_kompletnost: pd.DataFrame, df_top: pd.DataFrame, config: Config,
) -> pd.DataFrame:

    odejiti = nacti_odejite(config)

    if odejiti.empty:
        return pd.DataFrame()

    if config.sloupec_id not in odejiti.columns:
        raise ValueError(
            "Soubor odejiti.xlsx neobsahuje sloupec "
            f"'{config.sloupec_id}'."
        )

    warning_df = priprav_warning_odchody(
        warning_kompletnost,config,
    )

    inc_df = priprav_inc_odchody(
        df_top,config,
    )

    kandidati = pd.concat(
        [warning_df, inc_df], ignore_index=True, sort=False,
    )

    if kandidati.empty:
        return pd.DataFrame()

    vysledek = filtruj_odejite(
        kandidati, odejiti, config,
    )

    vysledek = uprav_vystup_odchody(
        vysledek, config,
    )

    logger.info(
        "Nalezeno %s odešlých zaměstnanců.",
        len(vysledek),
    )

    return vysledek

def zapis_warning_pracvz(
    writer: pd.ExcelWriter, warning_pracvz: pd.DataFrame,
) -> None:

    if warning_pracvz.empty:
        return

    warning_pracvz.to_excel(
        writer, sheet_name="WARNING-pracvz", index=False,
    )

# ==================================================
# BOXPLOT
# ==================================================

def vytvor_boxplot(
    df: pd.DataFrame, config: Config,
) -> None:
    """
    Vytvoří boxplot relativních změn mezd.

    Zpracovává pouze období definovaná
    v konfiguraci.

    Parametry
    ---------
    df : pd.DataFrame
        Data s vypočtenými relativními změnami.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    None
    """

    boxplot_data: list[pd.Series] = []
    platna_obdobi: list[str] = []
    popisky: list[str] = []

    for obdobi in config.obdobi_boxplot:

        hodnoty = (
            df.loc[
                df[config.sloupec_obdobi] == obdobi,
                "Rel_zmena_%",
            ]
            .dropna()
        )

        if hodnoty.empty:
            continue

        boxplot_data.append(hodnoty)
        platna_obdobi.append(obdobi)

        popisky.append(
            f"{obdobi}\n(n={len(hodnoty)})"
        )

    if not boxplot_data:
        logger.warning(
            "Pro boxplot nebyla nalezena žádná data."
        )
        return

    plt.figure(figsize=(10, 6))

    boxplot = plt.boxplot(
        boxplot_data, tick_labels=popisky, patch_artist=True,
        showfliers=False
    )

    for box in boxplot["boxes"]:
        box.set_facecolor("#8ecae6")

    mediany = [
        np.median(data)
        for data in boxplot_data
    ]

    prumery = [
            np.mean(data)
            for data in boxplot_data
        ]

    for i, (median, prumer) in enumerate(
        zip(mediany, prumery), start=1,
    ):
    
        plt.text(
            i, median, f"M: {median:.1f} %",
            ha="center", va="bottom", fontsize=9,
            fontweight="bold",
        )

        plt.text(
            i, prumer, f"Ø: {prumer:.1f} %",
            ha="center", va="top", fontsize=8,
            color="red",
        )

    plt.scatter(
        range(1, len(prumery) + 1),
        prumery, color="red", marker="D",
        label="Průměr", zorder=3,
    )

    plt.axhline(
        y=0, color="red", linestyle="--",
        linewidth=1, alpha=0.6,
    )

    celkem_zaznamu = sum(
        len(data)
        for data in boxplot_data
    )

    plt.title(
        "Relativní změna mezd včetně odměn (%)"
    )

    plt.suptitle(
        f"Celkem analyzováno {celkem_zaznamu:,} záznamů",
        fontsize=10,
    )

    plt.xlabel("Období")
    plt.ylabel("Relativní změna [%]")

    plt.grid(
        axis="y", linestyle="--", alpha=0.5,
    )

    plt.legend()

    plt.tight_layout()

    soubor_graf = (
        config.vstup.parent
        / f"boxplot_{config.vstup.stem}.png"
    )

    plt.savefig(
        soubor_graf, dpi=300, bbox_inches="tight",
    )

    plt.close()

    logger.info(
        "Graf uložen: %s", soubor_graf,
    )

    logger.info(
        "Boxplot vytvořen pro období: %s",
        ", ".join(platna_obdobi),
    )

def nacti_a_validuj_data(
    config: Config,
) -> pd.DataFrame:

    """
    Načte vstupní data a provede základní validace.

    Provádí:

    - načtení všech relevantních listů Excelu,
    - kontrolu neprázdného datasetu,
    - kontrolu přítomnosti povinných sloupců,
    - kontrolu konzistence odměn.

    Parametry
    ---------
    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    pd.DataFrame
        Načtená a základně validovaná data.

    Výjimky
    --------
    ValueError
        Pokud dataset neobsahuje žádná data
        nebo chybí povinné sloupce.

    FileNotFoundError
        Pokud vstupní soubor neexistuje.
    """

    df = nacti_data(config.soubor)

    if df.empty:
        raise ValueError(
            "Načtený soubor neobsahuje žádná data."
        )
    
    validuj_sloupce(df, config.povinne_sloupce,)

    # pracvz vždy jako text
    df["pracvz"] = (
        df["pracvz"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    kontrola_odmen(df, config)

    return df

def priprav_data(
    df: pd.DataFrame,config: Config,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame,
]:
    """
    Připraví data pro následnou analýzu.

    Provádí:

    - odstranění neplatných mezd,
    - převod a validaci období,
    - kontrolu požadovaných období,
    - kontrolu duplicit zaměstnanec–období,
    - kontrolu kompletnosti historie zaměstnanců.

    Parametry
    ---------
    df : pd.DataFrame
        Vstupní data.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]

    První položka:
        Očištěná data připravená k analýze.

    Druhá položka:
        Záznamy vyřazené z důvodu neplatné mzdy.

    Třetí položka:
        Záznamy vyřazené z důvodu nekompletní
        historie sledovaných období.

    Výjimky
    --------
    ValueError
        Pokud chybí požadovaná období nebo
        jsou nalezeny duplicitní kombinace
        zaměstnanec–období.
    """
    
    df, warning_mzdy = odstran_nulove_mzdy(df, config)

    df = priprav_obdobi(df, config)

    nalezena_obdobi = set(df[config.sloupec_obdobi])

    chybejici = (
        set(config.kontrolovana_obdobi)
        - nalezena_obdobi
    )

    if chybejici:
        raise ValueError(
            f"Chybí období: {sorted(chybejici)}"
        )

    duplicity = df[
        df.duplicated(
            subset=[config.sloupec_id, config.sloupec_obdobi,
            ],
            keep=False,
        )
    ]

    if not duplicity.empty:
        problematicke = (duplicity[[ config.sloupec_id, config.sloupec_obdobi,]]
                         .drop_duplicates()
                        )

        raise ValueError("Nalezeny duplicitní kombinace "  
                         f"zaměstnanec-období:\n{problematicke}")

    (
        df, warning_kompletnost,) = kontrola_kompletnosti(df, config,)

    warning_pracvz = kontrola_pracvz(df,config,)

    return (
        df, warning_mzdy, warning_kompletnost, warning_pracvz
    )

def proved_analyzu(
    df: pd.DataFrame,config: Config,
) -> tuple[pd.DataFrame, pd.DataFrame,]:
    """
    Provede vlastní analytické výpočty.

    Provádí:

    - výpočet relativních změn mzdy,
    - identifikaci extrémních růstů a poklesů,
    - přípravu dat pro export.

    Parametry
    ---------
    df : pd.DataFrame
        Očištěná a validovaná data.

    config : Config
        Konfigurace analýzy.

    Návratová hodnota
    -----------------
    tuple[pd.DataFrame, pd.DataFrame]

    První položka:
        Data doplněná o relativní změny mezd.

    Druhá položka:
        Zaměstnanci patřící do horního a dolního
        percentilu relativních změn.
    """
    
    df = vypocitej_relativni_zmenu(df, config,)

    df_top = vytvor_top_percentily(df, config,)

    return df, df_top

def vypis_souhrn(
    df_export: pd.DataFrame, warning_mzdy: pd.DataFrame, warning_kompletnost: pd.DataFrame,
    df_top: pd.DataFrame, df_odejiti: pd.DataFrame,
) -> None:
    """
    Vypíše souhrnné statistiky zpracování.
    """

    logger.info("=" * 60)
    logger.info("SOUHRN ZPRACOVÁNÍ")
    logger.info("=" * 60)

    logger.info(
        "Rel_zmena: %s řádků", len(df_export),
    )

    logger.info(
        "WARNING_MZDY0: %s řádků", len(warning_mzdy),
    )

    logger.info(
        "WARNING-kontrola kompletnosti: %s řádků", len(warning_kompletnost),
    )

    logger.info(
        "Top5_pct: %s řádků", len(df_top),
    )

    logger.info(
        "odejiti: %s řádků", len(df_odejiti),
    )

    logger.info("=" * 60)


# ==================================================
# MAIN
# ==================================================

def main() -> None:
    """
    Řídicí funkce celého zpracování.

    Koordinuje jednotlivé fáze analýzy:

    - načtení a validaci dat,
    - přípravu dat,
    - výpočet výsledků,
    - export výstupů,
    - vytvoření boxplotu.

    Návratová hodnota
    -----------------
    None
    """

    config = CONFIG

    try:

        df = nacti_a_validuj_data(config)

        (
            df, warning_mzdy, warning_kompletnost, warning_pracvz,) = priprav_data(df, config,)

        (
            df,df_top,) = proved_analyzu(df, config,)

        df_export = df[
            list(config.vystupni_sloupce)].copy()

        df_odejiti = vytvor_odejiti( warning_kompletnost, df_top, config,)

        logger.info(
            "df_odejiti rows: %s",
            len(df_odejiti),)


        export_excel( df_export, df_top, warning_mzdy, warning_kompletnost, warning_pracvz, df_odejiti, config)


        vytvor_boxplot(df, config)

    except Exception:
        logger.exception(
            "Zpracování skončilo chybou."
        )
        raise

    vypis_souhrn(
    df_export, warning_mzdy, warning_kompletnost,
    df_top, df_odejiti,
)

    


if __name__ == "__main__":
    main()