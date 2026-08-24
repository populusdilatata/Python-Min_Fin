"""
Rozdělení docházkových dat do kategorií.

Program načte vstupní Excel soubor, omezí data na
definovaný rozsah řádků, provede normalizaci vybraných
sloupců a rozdělí záznamy do pěti skupin:

- dlouhodobě nemocní,
- dovolené, 
- absence, 
- pracanti bez záznamu,
- ostatní.

Před exportem kontroluje kvalitu dat ve vybraných
povinných sloupcích a případné nedostatky zapisuje
do logu formou varování.

Výstupem je Excel soubor obsahující samostatné listy
pro jednotlivé skupiny záznamů.
"""

# Architektura               9,8/10 |
# Čitelnost                  9,9/10 |
# Dokumentace                9,8/10 |
# Logging                    9,8/10 |
# Python styl                9,8/10 |
# Robustnost                 9,8/10 |
# Udržovatelnost             9,9/10 |

from dataclasses import dataclass
import logging
import pandas as pd
from pathlib import Path

logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s"
)

@dataclass(frozen=True)
class Config:
    vstupni_sloupce: tuple[str, ...]
    vystupni_sloupce: tuple[str, ...]
    povinne_sloupce: tuple[str, ...]
    
    dochazka_sloupce: tuple[str, ...]
    
    pracanti_stavy: tuple[str, ...]
    
    dochazka_soubor: Path
    prohledavany_soubor: Path
    
    output_dir: Path
    vystupni_soubor: Path
    
    vybrany_sloupec: str
    
    posledni_radek: int
    min_pocet_zaznamu: int

    export_filtry: dict[str, tuple[str, ...]]

@dataclass(frozen=True)
class Rozdeleni:
    nemocni: pd.DataFrame
    dovolene: pd.DataFrame
    absence: pd.DataFrame
    pracanti: pd.DataFrame
    ostatni: pd.DataFrame

CFG = Config(
    vstupni_sloupce=(
        "oscis", "den", "dt", "prijm",
        "duvod", "duvodt", "duvod2", "duvod2t",
        "pracvmes", "priche", "odche",
    ),

    vystupni_sloupce=(
        "oscis",  "den", "dt", "prijm",
        "duvod", "duvodt", "duvod2", "duvod2t",
        "pracvmes", "priche", "odche",
        "ppvdr", "cindrz",
    ),

    povinne_sloupce=(
        "oscis", "den", "dt", "prijm",
        "duvod", "duvodt",
    ),

    dochazka_sloupce=(
        "oscis",
        "ppvdr",
        "cindrz",
    ),

    pracanti_stavy=(
        "odpra",
        "služ",
    ),

    dochazka_soubor=Path(
        "vstup_dochzarv2/07_DATA/"
        "dochzarv.xlsx"
    ),

    prohledavany_soubor=Path(
        "porovnavač/2026-08-20_12-51/"
        "25_DNI_BEZ_PRUCHODU_07_26_12_51.xlsx"
    ),

    output_dir=Path(
        "porovnavač/2026-08-20_12-51/"
    ),

    vystupni_soubor=Path("25_DNI_A_VICE_07_26.xlsx"),

    vybrany_sloupec="duvodt",

    export_filtry={
    "nemoc": ("nemoc",),
    "dovol": ("dovol",),
    "absen": ("absen",),
    "lekar": ("lékař",),
    "indiv": ("indiv",),
    "droby": ("očr","placv","studv",),
    },

    posledni_radek=2245,
    min_pocet_zaznamu=25,
)

def nacti_data(
    soubor: Path,
    posledni_radek: int,
) -> pd.DataFrame:
    """
    Načte vstupní Excel soubor a ořízne data
    na požadovaný počet řádků.
    """

    logger.info("Načítám soubor %s", soubor)

    if not soubor.exists():
        raise FileNotFoundError(
            f"Soubor neexistuje: {soubor}"
        )

    try:
        df = pd.read_excel(soubor)

    except Exception as exc:
        raise ValueError(
            f"Nepodařilo se načíst Excel soubor: {soubor}"
        ) from exc

    df = df.iloc[:posledni_radek]

    logger.info(
        "Načteno %s řádků",
        len(df),
    )

    return df


def vyber_sloupce(
    df: pd.DataFrame,
    sloupce: tuple[str, ...],
) -> pd.DataFrame:
    """
    Ověří, že vstupní DataFrame obsahuje všechny
    požadované sloupce a ponechá pouze vybrané.

    Args:
        df:
            Načtená data ze vstupního Excel souboru.

        sloupce:
            Seznam požadovaných sloupců.

    Returns:
        DataFrame obsahující pouze vybrané sloupce.

    Raises:
        ValueError:
            Pokud chybí některý z požadovaných sloupců.
    """

    chybi = set(sloupce) - set(df.columns)

    if chybi:
        raise ValueError(
            f"Chybí sloupce: {', '.join(sorted(chybi))}"
        )

    return df[list(sloupce)].copy()


def normalizuj_data(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Provede normalizaci vybraných sloupců.

    Sloupec 'oscis' je převeden na text a
    očištěn od okolních mezer.

    Sloupec 'pracvmes' je převeden na textový
    formát bez desetinné části.

    Args:
        df:
            Vstupní DataFrame.

    Returns:
        Normalizovaný DataFrame.
    """

    df["oscis"] = normalizuj_oscis(
        df["oscis"]
    )


    # pracvmes jako string bez desetinné části
    df["pracvmes"] = (
        pd.to_numeric(
            df["pracvmes"],
            errors="coerce",
        )
        .astype("Int64")
        .astype(str)
        .replace("<NA>", "")
    )

    return df


def rozdel_data(
    df: pd.DataFrame, 
    cfg: Config,
) -> Rozdeleni:
    """
    Rozdělí data do cílových kategorií.

    Vytváří pět samostatných množin dat:

    - dlouhodobě nemocní,
    - dovolené,
    - absence,
    - pracanti bez záznamu,
    - ostatní.

    Args:
        df:
            Vstupní DataFrame obsahující
            normalizovaná data.

    Returns:
        Objekt Rozdeleni obsahující
        data dlouhodobě nemocných,
        pracantů bez záznamu a ostatních.
    """

    nemoc_df = (
        df[df[cfg.vybrany_sloupec] == "nemoc"]
        .copy()
    )

    dovol_df = (
            df[df[cfg.vybrany_sloupec] == "dovol"]
            .copy()
        )

    absen_df = (
                df[df[cfg.vybrany_sloupec] == "absen"]
                .copy()
            )

    pracanti_bez_zaznamu = df[
        df[cfg.vybrany_sloupec].isin(cfg.pracanti_stavy)
    ]

    counts = (
        pracanti_bez_zaznamu["oscis"]
        .value_counts()
    )

    pracanti = (
        counts[counts >= cfg.min_pocet_zaznamu]
        .index
    )

    pracanti_df = pracanti_bez_zaznamu[
        pracanti_bez_zaznamu["oscis"]
        .isin(pracanti)
    ].copy()

    ostatni_df = df[
        ~df.index.isin(nemoc_df.index)
        & ~df.index.isin(pracanti_df.index)
        & ~df.index.isin(dovol_df.index)
        & ~df.index.isin(absen_df.index)
    ].copy()

    pocet_nemocnych = len(nemoc_df)
    pocet_dovolenych = len(dovol_df)
    pocet_absenci = len(absen_df)
    pocet_pracantu = len(pracanti_df)
    pocet_ostatnich = len(ostatni_df)

    logger.info(
        "Rozdělení: nemocni=%s, na dovolené=%s, absence=%s, pracanti=%s, ostatni=%s",
        pocet_nemocnych,
        pocet_dovolenych,
        pocet_absenci,
        pocet_pracantu,
        pocet_ostatnich,
)

    return Rozdeleni(
        nemocni=nemoc_df,
        dovolene=dovol_df,
        absence=absen_df,
        pracanti=pracanti_df,
        ostatni=ostatni_df,
    )


def zkontroluj_povinne_udaje(
    df: pd.DataFrame, 
    nazev_listu: str, 
    cfg: Config,
) -> None:
    """
    Zkontroluje vyplnění povinných sloupců.

    Pro každý povinný sloupec vypočítá počet
    chybějících hodnot a zapíše je do logu
    jako varování.

    Současně vyhodnotí počet řádků, které
    obsahují alespoň jednu chybějící povinnou
    hodnotu.

    Args:
        df:
            Kontrolovaný DataFrame.

        nazev_listu:
            Název výstupního listu použitý
            v logovacích zprávách.
    """

    for sloupec in cfg.povinne_sloupce:

        chybi = df[sloupec].isna().sum()

        if chybi > 0:
            logger.warning(
                "[%s] Sloupec '%s' obsahuje %s chybějících hodnot",
                nazev_listu,
                sloupec,
                chybi,
            )

    radky_s_chybou = (
        df[list(cfg.povinne_sloupce)]
        .isna()
        .any(axis=1)
        .sum()
    )

    if radky_s_chybou > 0:
        logger.warning(
            "[%s] %s řádků obsahuje alespoň jednu chybějící povinnou hodnotu",
            nazev_listu,
            radky_s_chybou,
        )

def nacti_dochazku(
    soubor: Path,
) -> pd.DataFrame:
    """
    Načte doplňkový docházkový soubor.

    Args:
        soubor:
            Cesta k docházkovému souboru.

    Returns:
        Načtený DataFrame.
    """

    logger.info(
        "Načítám doplňkový soubor %s",
        soubor,
    )

    return pd.read_excel(soubor)

def validuj_dochazku(
    doch_df: pd.DataFrame,
    cfg: Config,
) -> pd.DataFrame:

    """
    Ověří strukturu docházkových dat
    a připraví je pro následné spojení.

    Kontroluje přítomnost všech
    požadovaných sloupců, vybere pouze
    potřebné sloupce a normalizuje
    identifikátor oscis.

    Args:
        doch_df:
            Načtená docházková data.

        cfg:
            Konfigurace aplikace.

    Returns:
        Připravený DataFrame vhodný
        pro spojení s hlavními daty.

    Raises:
        ValueError:
            Pokud chybí některý
            z požadovaných sloupců.
    """

    chybejici_sloupce = (
        set(cfg.dochazka_sloupce)
        - set(doch_df.columns)
    )

    if chybejici_sloupce:
        raise ValueError(
            "V doplňkovém souboru chybí sloupce: "
            f"{', '.join(sorted(chybejici_sloupce))}"
        )

    doch_df = (
        doch_df[list(cfg.dochazka_sloupce)]
        .copy()
    )

    doch_df["oscis"] = normalizuj_oscis(
        doch_df["oscis"]
    )

    logger.info(
        "Docházkový soubor obsahuje %s záznamů.",
        len(doch_df),
    )

    return doch_df

def normalizuj_oscis(
    series: pd.Series,
) -> pd.Series:

    """
    Normalizuje identifikátory oscis.

    Převede hodnoty na textový formát
    a odstraní okolní bílé znaky.

    Args:
        series:
            Vstupní sloupec oscis.

    Returns:
        Normalizovaný sloupec.
    """
    return (
        series
        .astype(str)
        .str.strip()
    )

def najdi_nejednoznacne(
    hlavni_df: pd.DataFrame,
    doch_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Vyhledá nejednoznačné identifikátory oscis.

    Za nejednoznačné jsou považovány záznamy,
    u kterých je v docházkových datech pro
    stejné oscis evidováno více různých
    hodnot cindrz.

    Nalezené záznamy jsou vráceny v samostatném
    DataFrame doplněném o sloupec
    konfliktni_cindrz.

    Args:
        hlavni_df:
            Hlavní DataFrame obsahující
            zpracovávaná data.

        doch_df:
            Docházkový DataFrame použitý
            pro kontrolu jednoznačnosti.

    Returns:
        DataFrame obsahující všechny řádky
        z hlavních dat, které patří
        k nejednoznačným oscis.

        Pokud nejsou nalezeny žádné konflikty,
        vrací prázdný DataFrame.
    """


    pocet_hodnot = (
        doch_df.groupby("oscis")["cindrz"]
        .nunique()
    )

    konfliktni_oscis = (
        pocet_hodnot[pocet_hodnot > 1]
        .index
    )

    if len(konfliktni_oscis) == 0:
        return pd.DataFrame()

    konfliktni_map = (
        doch_df.groupby("oscis")["cindrz"]
        .apply(
            lambda s: "; ".join(
                sorted(
                    s.dropna()
                    .astype(str)
                    .unique()
                )
            )
        )
    )

    nejednoznacne_df = hlavni_df[
        hlavni_df["oscis"]
        .isin(konfliktni_oscis)
    ].copy()

    nejednoznacne_df["konfliktni_cindrz"] = (
        nejednoznacne_df["oscis"]
        .map(konfliktni_map)
    )

    logger.warning(
        "Nalezeno %s nejednoznačných oscis.",
        len(konfliktni_oscis),
    )

    logger.warning(
        "Do listu 'nejednoznacne' bude zařazeno %s řádků.",
        len(nejednoznacne_df),
    )

    logger.info(
        "Audit nejednoznačností: %s konfliktů, %s dotčených řádků.",
        len(konfliktni_oscis),
        len(nejednoznacne_df),
    )

    return nejednoznacne_df

def zkontroluj_duplicitni_oscis(
    doch_df: pd.DataFrame,
) -> None:

    """
    Zkontroluje výskyt duplicitních oscis
    v docházkových datech.

    Duplicitní hodnoty mohou způsobit
    nejednoznačnosti při spojování dat.
    Funkce proto zapíše varování do logu
    včetně ukázky nalezených identifikátorů.

    Args:
        doch_df:
            Docházkový DataFrame určený
            ke kontrole.
    """

    duplicitni_oscis = (
        doch_df.loc[
            doch_df["oscis"].duplicated(),
            "oscis",
        ]
        .unique()
    )

    if len(duplicitni_oscis) > 0:
        logger.warning(
            "Nalezeno %s duplicitních oscis. "
            "Budou použity pouze první výskyty. "
            "Příklady: %s",
            len(duplicitni_oscis),
            ", ".join(
                map(str, duplicitni_oscis[:10])
            ),
        )

def spoj_dochazku(
    hlavni_df: pd.DataFrame,
    doch_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Provede spojení hlavních dat
    s docházkovými údaji.

    Args:
        hlavni_df:
            Hlavní DataFrame.

        doch_df:
            Docházkový DataFrame.

    Returns:
        Spojený DataFrame.
    """
    puvodni_pocet = len(doch_df)

    doch_df = doch_df.drop_duplicates(
        subset=["oscis"],
        keep="first",
    )

    odebrano = puvodni_pocet - len(doch_df)

    if odebrano:
        logger.info(
            "Odstraněno %s duplicitních řádků docházky.",
            odebrano,
        )

    vysledek = hlavni_df.merge(
        doch_df,
        on="oscis",
        how="left",
        validate="many_to_one",
    )

    nenalezeno = (
        vysledek["ppvdr"]
        .isna()
        .sum()
    )

    sparovano = len(vysledek) - nenalezeno

    uspesnost = (
        100 * sparovano / len(vysledek)
        if len(vysledek)
        else 0
    )

    logger.info(
        "Spárováno %s z %s řádků (%.1f %%).",
        sparovano,
        len(vysledek),
        uspesnost,
    )


    if nenalezeno:
        logger.warning(
            "Pro %s záznamů nebyla nalezena shoda podle oscis.",
            nenalezeno,
        )

    return vysledek

def over_konzistenci_rozdeleni(
    puvodni_df: pd.DataFrame,
    rozdeleni: Rozdeleni,
) -> None:
    """
    Ověří konzistenci rozdělení dat.

    Kontroluje, zda součet řádků ve všech
    kategoriích odpovídá počtu řádků
    původního DataFrame.

    Args:
        puvodni_df:
            Původní DataFrame před rozdělením.

        rozdeleni:
            Výsledek rozdělení dat.

    Raises:
        RuntimeError:
            Pokud součet řádků v kategoriích
            neodpovídá původnímu počtu řádků.
    """

    celkem = (
        len(rozdeleni.nemocni)
        +len(rozdeleni.dovolene)
        +len(rozdeleni.absence)
        + len(rozdeleni.pracanti)
        + len(rozdeleni.ostatni)
    )

    if celkem != len(puvodni_df):
        raise RuntimeError(
            f"Nekonzistentní rozdělení dat. "
            f"Původní={len(puvodni_df)}, "
            f"dovolené={len(rozdeleni.dovolene)}, "
            f"absence={len(rozdeleni.absence)}, "
            f"nemocni={len(rozdeleni.nemocni)}, "
            f"pracanti={len(rozdeleni.pracanti)}, "
            f"ostatni={len(rozdeleni.ostatni)}."
        )

    logger.info(
        "Konzistence rozdělení ověřena (%s řádků).",
        celkem,
    )

def vytvor_histogram_duvodu(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
) -> None:

    cetnosti = (
        df["duvodt"]
        .fillna("(prázdné)")
        .value_counts()
        .rename_axis("duvodt")
        .reset_index(name="pocet")
    )

    cetnosti.to_excel(
        writer,
        sheet_name="histogram_duvodt",
        index=False,
    )

    logger.info(
        "Vytvořen přehled četností duvodt (%s kategorií).",
        len(cetnosti),
    )


def uloz_data(
    vystupni_soubor: Path,
    nemoc_df: pd.DataFrame,
    dovolena_df: pd.DataFrame,
    absen_df: pd.DataFrame,
    pracanti_df: pd.DataFrame,
    ostatni_df: pd.DataFrame,
    nejednoznacne_df: pd.DataFrame,
    zdrojovy_df: pd.DataFrame
) -> None:
    """
    Uloží rozdělená data do Excel souboru.

    Každá kategorie je uložena do samostatného
    listu výsledného sešitu.

    Args:
        vystupni_soubor:
            Cesta k výslednému Excel souboru.

        nemoc_df:
            Data dlouhodobě nemocných.

        pracanti_df:
            Data pracantů bez záznamu.

        ostatni_df:
            Všechna ostatní data.
    """


    with pd.ExcelWriter(vystupni_soubor) as writer:
        nemoc_df.to_excel(
            writer,
            sheet_name="dlouhodobe_nemocni",
            index=False,
        )

        dovolena_df.to_excel(
                    writer,
                    sheet_name="dovolena",
                    index=False,
                )

        absen_df.to_excel(
                            writer,
                            sheet_name="absence",
                            index=False,
                        )

        pracanti_df.to_excel(
            writer,
            sheet_name="pracanti_bez_zaznamu",
            index=False,
        )

        ostatni_df.to_excel(
            writer,
            sheet_name="ostatni",
            index=False,
        )

        if not nejednoznacne_df.empty:
            nejednoznacne_df.to_excel(
                writer,
                sheet_name="nejednoznacne",
                index=False,
            )
        vytvor_histogram_duvodu(
            writer,
            zdrojovy_df,
        )

    logger.info(
        "Výstup uložen do %s",
        vystupni_soubor,
    )

def exportuj_filtrovane_soubory(
    df: pd.DataFrame,
    cfg: Config,
) -> None:
    """
    Vytvoří samostatné Excel soubory
    pro vybrané skupiny hodnot ve sloupci
    'duvodt'.
    """

    for nazev, hodnoty in cfg.export_filtry.items():

        filtr_df = df[
            df["duvodt"].isin(hodnoty)
        ].copy()

        if filtr_df.empty:
            logger.info(
                "Filtr %s neobsahuje žádná data.",
                nazev,
            )
            continue

        vystup = (
            cfg.output_dir
            / f"pdnyv_07_26_FILTER_{nazev}_12_51.xlsx"
        )

        filtr_df.to_excel(
            vystup,
            index=False,
        )

        logger.info(
            "Exportován soubor %s (%s řádků).",
            vystup.name,
            len(filtr_df),
        )


def main() -> None:

    cfg = CFG
    logger.info(
        "=== Zahajuji zpracování docházkových dat ==="
    )

    try:

        df = nacti_data(
            cfg.prohledavany_soubor,
            cfg.posledni_radek,
        )

        df = vyber_sloupce(
            df,
            cfg.vstupni_sloupce,
        )

        df = normalizuj_data(df)

        doch_df = nacti_dochazku(
            cfg.dochazka_soubor,
        )

        doch_df = validuj_dochazku(
            doch_df,
            cfg,
        )

        zkontroluj_duplicitni_oscis(
            doch_df,
        )

        nejednoznacne_df = najdi_nejednoznacne(
            df, doch_df,
        )

        df = spoj_dochazku(
            df, doch_df,
        )


        rozdeleni = rozdel_data(df, cfg)

        over_konzistenci_rozdeleni(
            df, 
            rozdeleni,
        )

        nemoc_df = rozdeleni.nemocni[list(cfg.vystupni_sloupce)]
        dovolena_df = rozdeleni.dovolene[list(cfg.vystupni_sloupce)]
        absence_df = rozdeleni.absence[list(cfg.vystupni_sloupce)]
        pracanti_df = rozdeleni.pracanti[list(cfg.vystupni_sloupce)]
        ostatni_df = rozdeleni.ostatni[list(cfg.vystupni_sloupce)]

        

        zkontroluj_povinne_udaje(
            nemoc_df,"dlouhodobe_nemocni", cfg,
        )

        zkontroluj_povinne_udaje(
            pracanti_df,"pracanti_bez_zaznamu", cfg,
        )

        zkontroluj_povinne_udaje(
            ostatni_df, "ostatni", cfg,
        )

        logger.info(
            "Připraveno k exportu: "
            "nemocni=%s, pracanti=%s, ostatni=%s, "
            "nejednoznacne=%s",
            len(nemoc_df),
            len(pracanti_df),
            len(ostatni_df),
            len(nejednoznacne_df),
        )

        exportuj_filtrovane_soubory(
            df, cfg,
        )

        uloz_data(
            cfg.output_dir / cfg.vystupni_soubor,
            nemoc_df,  dovolena_df, absence_df, 
            pracanti_df, ostatni_df, nejednoznacne_df, df
        )

        logger.info(
            "=== Zpracování úspěšně dokončeno ==="
        )

    except PermissionError as exc:
        logger.error(
            "Přístup k souboru byl odepřen: %s",
            exc,)

    except FileNotFoundError as exc:
        logger.error(
        "Soubor nebyl nalezen: %s",
        exc,
        )

    except Exception:
        logger.exception(
        "Neočekávaná chyba při zpracování dat."
        )


if __name__ == "__main__":
    main()