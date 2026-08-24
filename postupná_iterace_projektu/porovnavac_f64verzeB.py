"""

Porovnání výstupů PDNYV a F64.

Skript načte zdrojová data z Excel souborů,
porovná jejich obsah a vytvoří přehled rozdílů.

Zpracování zahrnuje:

- vyhledání shodných záznamů,
- vyhledání záznamů pouze v PDNYV,
- vyhledání záznamů pouze ve F64,
- doplnění informací z docházkových tabulek,
- vytvoření výsledného Excel souboru.

Výstup obsahuje samostatné listy se shodami
a rozdíly mezi porovnávanými systémy.

Verze:
    1.1

Datum vytvoření:
    2026-08-14

Požadavky:
    pandas
    openpyxl
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from time import perf_counter

import pandas as pd
from pandas.api.types import is_object_dtype




# Architektura               9,5/10
# Čitelnost                  9,7/10
# Dokumentace                9,5/10
# Logging                    9,5/10
# Python styl                9,5/10
# Robustnost                 9,5/10
# Udržovatelnost             9,6/10

# ---------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------
# KONSTANTY
# ---------------------------------------------------------------------

SOUBOR0_DOCHZARV2A = Path(
    "vstup_dochzarv2/07_DATA/dochzarv_07A_FILTER_vztahorg_NOT3.xlsx"
)

SOUBOR0_DOCHZARV2B = Path(
    "vstup_dochzarv2/07_DATA/dochzarv_07A_FILTER_vztahorg_3.xlsx"
)

JMENO_VYSLEDNEHO_SOUBORU = "vysledek_porovnani_"

OSCIS = "oscis"
DEN = "den"
AKTUAL_DEN = "aktual_den"
PRIJM = "prijm"
PRACVMES = "pracvmes"
PRACV = "pracv"

PPVDR = "ppvdr"

@dataclass(frozen=True)
class KonfiguraceFiltru:
    filtr: str
    soubor_pdnyv: Path
    soubor_f64: Path
    vystupni_adresar: Path

# ---------------------------------------------------------------------
# DEFINICE FILTRŮ
# ---------------------------------------------------------------------

KONFIGURACE = [
    KonfiguraceFiltru(
        filtr="OTECD",
        soubor_pdnyv=Path(
            "porovnavač/porovnavač_DROBY_OTECD/07_DATA/"
            "pdnyv_07A_26_droby_FILTER_OTECD.xlsx"
        ),
        soubor_f64=Path(
            "porovnavač/porovnavač_DROBY_OTECD/07_DATA/"
            "vystupf64_07A_OTECD.xlsx"
        ),
        vystupni_adresar=Path(
            "porovnavač/porovnavač_DROBY_OTECD"
        ),
    ),
    KonfiguraceFiltru(
        filtr="OCR",
        soubor_pdnyv=Path(
            "porovnavač/porovnavač_DROBY_OCR/07_DATA/"
            "pdnyv_07A_26_droby_FILTER_OCR.xlsx"
        ),
        soubor_f64=Path(
            "porovnavač/porovnavač_DROBY_OCR/07_DATA/"
            "vystupf64_07A_OCR.xlsx"
        ),
        vystupni_adresar=Path(
            "porovnavač/porovnavač_DROBY_OCR"
        ),
    ),
    KonfiguraceFiltru(
        filtr="NEMOC",
        soubor_pdnyv=Path(
            "porovnavač/porovnavač_NEMOC/07_DATA/"
            "pdnyv_07A_26_FILTER_nemoc_10_07.xlsx"
        ),
        soubor_f64=Path(
            "porovnavač/porovnavač_NEMOC/07_DATA/"
            "vystupf64_07A_NEMOC.xlsx"
        ),
        vystupni_adresar=Path(
            "porovnavač/porovnavač_NEMOC"
        ),
    ),
]

POVINNE_SLOUPCE_TAB1 = {
    OSCIS,
    DEN,
    PRIJM,
    PRACVMES,
}

POVINNE_SLOUPCE_TAB2 = {
    OSCIS,
    AKTUAL_DEN,
    PRIJM,
    PRACV,
}

POVINNE_SLOUPCE_TAB0 = {
    OSCIS,
    PPVDR,
}


def kontrola_souboru(cesta: Path) -> None:
    """
    Ověří existenci souboru.
    """
    if not cesta.exists():
        raise FileNotFoundError(f"Soubor neexistuje: {cesta}")

def validuj_prazdna_data(
    df: pd.DataFrame,
    nazev_tabulky: str,
) -> None:
    """
    Ověří, že tabulka obsahuje alespoň
    jeden datový řádek.

    Args:
        df: Kontrolovaný DataFrame.
        nazev_tabulky: Název tabulky.

    Raises:
        ValueError:
            Pokud tabulka neobsahuje žádná data.
    """

    if df.empty:
        raise ValueError(
            f"{nazev_tabulky}: tabulka neobsahuje žádná data"
        )

def validuj_duplicity(
    df: pd.DataFrame,
    klice: list[str],
    nazev_tabulky: str,
) -> None:

    """
    Ověří, že zadané klíčové sloupce
    neobsahují duplicitní kombinace hodnot.

    Args:
        df: Kontrolovaný DataFrame.
        klice: Seznam sloupců tvořících klíč.
        nazev_tabulky: Název tabulky.

    Raises:
        ValueError:
            Pokud jsou nalezeny duplicitní klíče.
    """

    duplicity_df = df[
        df.duplicated(
            subset=klice,
            keep=False,
        )
    ]

    if not duplicity_df.empty:
        ukazka = (
            duplicity_df[klice]
            .drop_duplicates()
            .head(5)
            .to_dict("records")
        )

        raise ValueError(
            f"{nazev_tabulky}: nalezeny "
            f"duplicitní klíče. "
            f"Ukázka: {ukazka}"
        )

def validuj_sloupce(
    df: pd.DataFrame,
    povinne_sloupce: set[str],
    nazev_tabulky: str,
) -> None:
    """
    Ověří přítomnost všech povinných sloupců.

    Args:
        df: Kontrolovaný DataFrame.
        povinne_sloupce: Množina povinných sloupců.
        nazev_tabulky: Název tabulky používaný v chybových hlášeních.

    Raises:
        ValueError:
            Pokud některý z povinných sloupců chybí.
    """
    
    chybejici = povinne_sloupce - set(df.columns)

    if chybejici:
        raise ValueError(
            f"{nazev_tabulky}: chybí sloupce {sorted(chybejici)}"
        )

def normalizuj_datove_typy(
    tab0a: pd.DataFrame,
    tab0b: pd.DataFrame,
    tab1: pd.DataFrame,
    tab2: pd.DataFrame,
) -> None:
    """
    Sjednotí datové typy klíčových sloupců.

    Funkce převádí identifikátory zaměstnanců OSCIS
    na textový typ a odstraňuje případné okolní mezery.
    Zároveň převádí sloupce obsahující datum na typ
    datetime, aby bylo možné provádět spolehlivé
    porovnávání a spojování dat pomocí merge.

    Args:
        tab0a:
            Docházková tabulka bez vztahorg = 3.
        tab0b:
            Docházková tabulka pro vztahorg = 3.
        tab1:
            Zdrojová tabulka PDNYV.
        tab2:
            Zdrojová tabulka F64.

    Raises:
        ValueError:
            Pokud některý z datových sloupců obsahuje
            neplatnou hodnotu, kterou nelze převést
            na datum.
    """

    for df in (tab0a, tab0b, tab1, tab2):
        df[OSCIS] = (
            df[OSCIS]
            .astype(str)
            .str.strip()
        )

    tab1[DEN] = pd.to_datetime(
        tab1[DEN],
        errors="raise",
    )

    tab2[AKTUAL_DEN] = pd.to_datetime(
        tab2[AKTUAL_DEN],
        errors="raise",
    )

def validuj_klice(
    df: pd.DataFrame,
    klice: list[str],
    nazev_tabulky: str,
) -> None:
    """
    Ověří, že klíčové sloupce neobsahují
    chybějící nebo prázdné hodnoty.

    Args:
        df: Kontrolovaný DataFrame.
        klice: Seznam klíčových sloupců.
        nazev_tabulky: Název tabulky.

    Raises:
        ValueError:
            Pokud je v některém klíči nalezena
            chybějící nebo prázdná hodnota.
    """

    for klic in klice:

        if df[klic].isna().any():
            pocet = df[klic].isna().sum()

            raise ValueError(
                f"{nazev_tabulky}: sloupec "
                f"{klic} obsahuje "
                f"{pocet} prázdných hodnot"
            )

        if is_object_dtype(df[klic]):

            prazdne = (
                df[klic]
                .astype(str)
                .str.strip()
                .eq("")
            )

            if prazdne.any():

                pocet = prazdne.sum()

                raise ValueError(
                    f"{nazev_tabulky}: sloupec "
                    f"{klic} obsahuje "
                    f"{pocet} prázdných řetězců"
                )

def nacti_excel(cesta: Path) -> pd.DataFrame:
    """
    Načte Excel soubor a vrátí jeho obsah jako DataFrame.

    Args:
        cesta: Cesta ke vstupnímu Excel souboru.

    Returns:
        DataFrame obsahující načtená data.

    Raises:
        FileNotFoundError:
            Pokud soubor neexistuje.
    """

    logger.info("Načítám %s", cesta)

    kontrola_souboru(cesta)

    df = pd.read_excel(cesta)

    logger.info(
        "%s načten (%s řádků, %s sloupců)",
        cesta.name,
        len(df),
        len(df.columns),
    )

    return df


def najdi_shodne(
    tab1: pd.DataFrame,
    tab2: pd.DataFrame,
) -> pd.DataFrame:
    """
    Vyhledá záznamy existující současně v obou tabulkách.

    Porovnání probíhá podle kombinace identifikátoru
    zaměstnance a dne.

    Args:
        tab1: Zdrojová tabulka PDNYV.
        tab2: Zdrojová tabulka F64.

    Returns:
        DataFrame obsahující pouze shodné záznamy.
    """

    return pd.merge(
        tab1,
        tab2,
        left_on=[OSCIS, DEN],
        right_on=[OSCIS, AKTUAL_DEN],
        how="inner",
    )


def najdi_pouze_v_pdnyv(
    tab1: pd.DataFrame,
    tab2: pd.DataFrame,
) -> pd.DataFrame:
    """
    Vyhledá záznamy, které existují pouze v tabulce PDNYV.

    Porovnání probíhá podle kombinace sloupců:
    oscis, den, prijm a pracvmes.

    Args:
        tab1: Zdrojová tabulka PDNYV.
        tab2: Zdrojová tabulka F64.

    Returns:
        DataFrame obsahující pouze záznamy nalezené
        v tabulce PDNYV.
    """

    vysledek = pd.merge(
        tab1,
        tab2,
        left_on=[OSCIS, DEN, PRIJM, PRACVMES],
        right_on=[OSCIS, AKTUAL_DEN, PRIJM, PRACV],
        how="left",
        indicator=True,
    )

    return vysledek[
        vysledek["_merge"] == "left_only"
    ][[OSCIS, DEN, PRIJM, PRACVMES]]


def najdi_pouze_ve_f64(
    tab1: pd.DataFrame,
    tab2: pd.DataFrame,
) -> pd.DataFrame:
    """
    Vyhledá záznamy, které existují pouze v tabulce F64.

    Porovnání probíhá podle kombinace sloupců:
    oscis, aktual_den, prijm a pracv.

    Args:
        tab1:
            Zdrojová tabulka PDNYV.
        tab2:
            Zdrojová tabulka F64.

    Returns:
        DataFrame obsahující pouze záznamy nalezené
        v tabulce F64.
    """

    vysledek = pd.merge(
        tab1,
        tab2,
        left_on=[OSCIS, DEN, PRIJM, PRACVMES],
        right_on=[OSCIS, AKTUAL_DEN, PRIJM, PRACV],
        how="right",
        indicator=True,
    )

    return vysledek[
        vysledek["_merge"] == "right_only"
    ][[OSCIS, AKTUAL_DEN, PRIJM, PRACV]]


def dopln_ppvdr(
    df: pd.DataFrame,
    tab0: pd.DataFrame,
    sloupec_oscis: str,
) -> pd.DataFrame:
    """
    Doplní do tabulky hodnotu PPVDR podle OSCIS.

    Args:
        df: DataFrame rozšiřovaný o sloupec PPVDR.
        tab0: Referenční tabulka obsahující OSCIS a PPVDR.
        sloupec_oscis: Název sloupce obsahujícího OSCIS.

    Returns:
        DataFrame rozšířený o sloupec PPVDR.
    """

    return pd.merge(
        df,
        tab0[[OSCIS, PPVDR]],
        left_on=sloupec_oscis,
        right_on=OSCIS,
        how="left",
    )


def uloz_vysledek(
    vystupni_soubor: Path,
    shodne: pd.DataFrame,
    pouze_pdnyv_a: pd.DataFrame,
    pouze_pdnyv_b: pd.DataFrame,
    pouze_f64_a: pd.DataFrame,
    pouze_f64_b: pd.DataFrame,
) -> None:
    """
    Uloží výsledky porovnání do Excel souboru.

    Vytvoří samostatné listy pro shodné záznamy
    a rozdíly mezi jednotlivými zdroji.

    Args:
        vystupni_soubor: Cesta k výslednému souboru.
        shodne: Shodné záznamy.
        pouze_pdnyv_a: Záznamy pouze v PDNYV.
        pouze_pdnyv_b: Záznamy pouze v PDNYV z externího zdroje.
        pouze_f64_a: Záznamy pouze ve F64.
        pouze_f64_b: Záznamy pouze ve F64 z externího zdroje.
    """

    logger.info("Ukládám %s", vystupni_soubor)

    with pd.ExcelWriter(
        vystupni_soubor,
        engine="openpyxl",
    ) as writer:
        shodne.to_excel(
            writer,
            sheet_name="SHODNE",
            index=False,
        )

        pouze_pdnyv_a.to_excel(
            writer,
            sheet_name="POUZE_pdnyv",
            index=False,
        )

        pouze_pdnyv_b.to_excel(
            writer,
            sheet_name="POUZE_pdnyv_POUZE_EXT",
            index=False,
        )

        pouze_f64_a.to_excel(
            writer,
            sheet_name="POUZE_f64",
            index=False,
        )

        pouze_f64_b.to_excel(
            writer,
            sheet_name="POUZE_f64_POUZE_EXT",
            index=False,
        )

    logger.info(
        "Výsledek uložen: %s",
        vystupni_soubor,
    )

def validuj_vstupni_data(
    tab0a: pd.DataFrame,
    tab0b: pd.DataFrame,
    tab1: pd.DataFrame,
    tab2: pd.DataFrame,
) -> None:
    """
    Provede kompletní validaci a normalizaci
    vstupních dat.

    Funkce ověřuje přítomnost dat, kontroluje
    povinné sloupce, sjednocuje datové typy,
    validuje klíčové sloupce a kontroluje
    duplicity klíčů.

    Args:
        tab0a:
            Docházková tabulka bez vztahorg = 3.
        tab0b:
            Docházková tabulka pro vztahorg = 3.
        tab1:
            Zdrojová tabulka PDNYV.
        tab2:
            Zdrojová tabulka F64.

    Raises:
        ValueError:
            Pokud některá z tabulek neobsahuje
            požadovaná data, povinné sloupce,
            obsahuje neplatné klíče nebo
            duplicitní záznamy.
    """

    validuj_prazdna_data(tab0a, "TAB0A")
    validuj_prazdna_data(tab0b, "TAB0B")
    validuj_prazdna_data(tab1, "TAB1")
    validuj_prazdna_data(tab2, "TAB2")

    validuj_sloupce(tab0a, POVINNE_SLOUPCE_TAB0, "TAB0A")
    validuj_sloupce(tab0b, POVINNE_SLOUPCE_TAB0, "TAB0B")
    validuj_sloupce(tab1, POVINNE_SLOUPCE_TAB1, "TAB1")
    validuj_sloupce(tab2, POVINNE_SLOUPCE_TAB2, "TAB2")

    normalizuj_datove_typy(
        tab0a,
        tab0b,
        tab1,
        tab2,
    )

    validuj_klice(tab0a, [OSCIS], "TAB0A")
    validuj_klice(tab0b, [OSCIS], "TAB0B")
    validuj_klice(tab1, [OSCIS, DEN], "TAB1")
    validuj_klice(tab2, [OSCIS, AKTUAL_DEN], "TAB2")

    validuj_duplicity(tab0a, [OSCIS], "TAB0A")
    validuj_duplicity(tab0b, [OSCIS], "TAB0B")
    validuj_duplicity(tab1, [OSCIS, DEN], "TAB1")
    validuj_duplicity(tab2, [OSCIS, AKTUAL_DEN], "TAB2")

def proved_porovnani(
    tab1: pd.DataFrame,
    tab2: pd.DataFrame,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    """
    Provede porovnání dat PDNYV a F64.

    Returns:
        (
            shodne,
            pouze_pdnyv,
            pouze_f64,
        )
    """

    shodne = najdi_shodne(
        tab1,
        tab2,
    )

    pouze_tab1 = najdi_pouze_v_pdnyv(
        tab1,
        tab2,
    )

    pouze_tab2 = najdi_pouze_ve_f64(
        tab1,
        tab2,
    )

    return (
        shodne,
        pouze_tab1,
        pouze_tab2,
    )

def dopln_vsechna_ppvdr(
    pouze_tab1: pd.DataFrame,
    pouze_tab2: pd.DataFrame,
    tab0a: pd.DataFrame,
    tab0b: pd.DataFrame,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    """
    Doplní hodnotu PPVDR do všech výstupních tabulek.

    Pro záznamy nalezené pouze v PDNYV i pouze ve F64
    provede napojení na obě docházkové tabulky a vrátí
    výsledné datové sady rozšířené o sloupec PPVDR.

    Args:
        pouze_tab1:
            Záznamy nalezené pouze v PDNYV.
        pouze_tab2:
            Záznamy nalezené pouze ve F64.
        tab0a:
            Docházková tabulka bez vztahorg = 3.
        tab0b:
            Docházková tabulka pro vztahorg = 3.

    Returns:
        N-tice obsahující:

        (
            pouze_pdnyv_a,
            pouze_pdnyv_b,
            pouze_f64_a,
            pouze_f64_b,
        )

        kde jednotlivé tabulky obsahují doplněné
        hodnoty PPVDR z odpovídající docházkové
        tabulky.
    """

    pouze_tab1a = dopln_ppvdr(
        pouze_tab1,
        tab0a,
        OSCIS,
    )

    pouze_tab1b = dopln_ppvdr(
        pouze_tab1,
        tab0b,
        OSCIS,
    )

    pouze_tab2a = dopln_ppvdr(
        pouze_tab2,
        tab0a,
        OSCIS,
    )

    pouze_tab2b = dopln_ppvdr(
        pouze_tab2,
        tab0b,
        OSCIS,
    )

    return (
        pouze_tab1a,
        pouze_tab1b,
        pouze_tab2a,
        pouze_tab2b,
    )


def porovnavac_f64(
    filtr: str,
    soubor0a: Path,
    soubor0b: Path,
    soubor1: Path,
    soubor2: Path,
    vystupni_adresar: Path,
) -> None:


    """
    Provede kompletní porovnání dat mezi PDNYV a F64.

    Workflow zpracování:

    1. načtení vstupních dat,
    2. validace a normalizace dat,
    3. vyhledání shodných a rozdílných záznamů,
    4. doplnění údajů PPVDR z docházkových tabulek,
    5. vytvoření výsledného Excel souboru,
    6. zápis statistik do logu.

    Args:
        filtr:
            Název zpracovávaného filtru.
        soubor0a:
            Docházková tabulka bez vztahorg = 3.
        soubor0b:
            Docházková tabulka pro vztahorg = 3.
        soubor1:
            Zdrojová tabulka PDNYV.
        soubor2:
            Zdrojová tabulka F64.
        vystupni_adresar:
            Adresář pro uložení výsledného souboru.

    Raises:
        FileNotFoundError:
            Pokud některý vstupní soubor neexistuje.
        ValueError:
            Pokud vstupní data neprojdou validací.
    """

    start = perf_counter()

    logger.info(
        "=================================================="
    )
    logger.info(
        "Zahajuji zpracování filtru %s",
        filtr,
    )

    logger.info(
        "[%s] Načítání vstupních dat",
        filtr,
    )

    tab0a = nacti_excel(soubor0a)
    tab0b = nacti_excel(soubor0b)
    tab1 = nacti_excel(soubor1)
    tab2 = nacti_excel(soubor2)

    logger.info(
        "[%s] Validace dat",
        filtr,
    )

    validuj_vstupni_data(
        tab0a,
        tab0b,
        tab1,
        tab2,
    )

    logger.info(
        "[%s] Vyhledávání rozdílů",
        filtr,
    )

    shodne, pouze_tab1, pouze_tab2 = (
        proved_porovnani(
            tab1,
            tab2,
        )
    )

    logger.info(
        "Shodné: %s",
        len(shodne),
    )
    logger.info(
        "Pouze PDNYV: %s",
        len(pouze_tab1),
    )
    logger.info(
        "Pouze F64: %s",
        len(pouze_tab2),
    )

    logger.info(
        "[%s] Doplňování PPVDR",
        filtr,
    )

    (
        pouze_tab1a,
        pouze_tab1b,
        pouze_tab2a,
        pouze_tab2b,
    ) = dopln_vsechna_ppvdr(
        pouze_tab1,
        pouze_tab2,
        tab0a,
        tab0b,
    )

    vystupni_soubor = (
        vystupni_adresar
        / f"{JMENO_VYSLEDNEHO_SOUBORU}{filtr}.xlsx"
    )

    uloz_vysledek(
        vystupni_soubor,
        shodne,
        pouze_tab1a,
        pouze_tab1b,
        pouze_tab2a,
        pouze_tab2b,
    )

    logger.info(
        "[%s] Hotovo | shodne=%s | pdnyv=%s | f64=%s",
        filtr,
        len(shodne),
        len(pouze_tab1),
        len(pouze_tab2),
    )

    logger.info(
        "[%s] Dokončeno za %.2f s",
        filtr,
        perf_counter() - start,
    )


def main() -> None:
    """
    Spustí zpracování všech definovaných filtrů.

    Postupně provede porovnání pro:

    - OTECD
    - OCR
    - NEMOC
    """

    uspesne = 0
    neuspesne = 0

    for konfigurace in KONFIGURACE:
        try:
            porovnavac_f64(
                filtr=konfigurace.filtr,
                soubor0a=SOUBOR0_DOCHZARV2A,
                soubor0b=SOUBOR0_DOCHZARV2B,
                soubor1=konfigurace.soubor_pdnyv,
                soubor2=konfigurace.soubor_f64,
                vystupni_adresar=konfigurace.vystupni_adresar,
            )

            uspesne += 1

        except Exception:
            neuspesne += 1

            logger.exception(
                "Zpracování filtru %s selhalo",
                konfigurace.filtr,
            )

    logger.info(
        "Dokončeno. Úspěšně: %s, Neúspěšně: %s",
        uspesne, neuspesne,
    )

    


if __name__ == "__main__":
    main()