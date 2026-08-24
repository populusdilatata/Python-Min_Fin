"""
Skript pro rozdělení souboru F60 podle hodnot ve sloupci neo2dr.

Vstup:
    vstup_f60/07_DATA/f60.xlsx

Logika:
    Záznamům je přiřazena kategorie podle hodnoty ve sloupci neo2dr.
    Každá kategorie je následně exportována do samostatného XLSX souboru.

Výstup:
    vstup_f60/f60_FILTER_<KATEGORIE>.xlsx

Kategorie:
    LEKAR
    PLACV
    INDIV
    ABSEN
    DOVOL
    SVZOZ
    STUDV
    DROBY

Autor: Majda Tomáš
Verze: 1.0
"""


import logging
from pathlib import Path

import pandas as pd

# ================================

# Architektura               8,9/10
# Čitelnost                  9,2/10
# Dokumentace                9,7/10
# Logging                    9,4/10
# Python styl                8,9/10
# Robustnost                 9,0/10
# Udržovatelnost             9,3/10

# ============================================================================
# KONFIGURACE
# ============================================================================

ROZDELOVANY_SOUBOR = "vstup_f60/07_DATA/f60.xlsx"
OUTPUT_DIR = "vstup_f60"

SLOUPEC_KATEGORIE = "neo2dr"
DATUM_OD = "neo2za"
DATUM_DO = "neo2ko"

# Mapování hodnot neo2dr na exportní kategorie.
#
# LEKAR = Lékaři
# PLACV = Placené volno
# INDIV = Individuální agenda
# ABSEN = Absence
# DOVOL = Dovolená
# SVZOZ = Služební volno zařízení osobních záležitostí
# STUDV = Studijní volno
# DROBY = všechny neklasifikované záznamy

MAP_KATEGORIE = {
    "ABSEN": [75, 85],
    "DOVOL": [95],
    "INDIV": [158, 52],
    "LEKAR": [140, 77, 139, 82, 160],
    "PLACV": [143, 78, 147, 137, 163, 150],
    
    "STUDV": [181],    
    "SVZOZ": [151],    
}

# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler("f60.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


# ============================================================================
# FUNKCE
# ============================================================================

def prirad_kategorii(hodnota):
    """
    Přiřadí kategorii na základě hodnoty ve sloupci neo2dr.

    Funkce porovná zadanou hodnotu s definovanými
    kategoriemi v MAP_KATEGORIE. Pokud hodnota není
    nalezena v žádné kategorii, vrátí kategorii DROBY.

    Parameters
    ----------
    hodnota : int
        Hodnota ze sloupce neo2dr.

    Returns
    -------
    str
        Název nalezené kategorie nebo hodnota "DROBY".
    """

    for kategorie, hodnoty in MAP_KATEGORIE.items():
        if hodnota in hodnoty:
            return kategorie

    return "DROBY"



def validuj_sloupce(df):
    """
    Ověří, že vstupní DataFrame obsahuje
    všechny povinné sloupce.

    Parameters
    ----------
    df : pandas.DataFrame
        Načtená data ze zdrojového souboru.

    Raises
    ------
    ValueError
        Pokud některý povinný sloupec chybí.
    """

    povinne_sloupce = {
        SLOUPEC_KATEGORIE,
        DATUM_OD,
        DATUM_DO,
    }

    chybejici = povinne_sloupce - set(df.columns)

    if chybejici:
        raise ValueError(
            f"Chybí povinné sloupce: {', '.join(chybejici)}"
        )


def formatuj_datum(df, sloupec):
    """
    Převede datumový sloupec do formátu DD.MM.RRRR.

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame obsahující zpracovávaná data.

    sloupec : str
        Název datumového sloupce.

    Returns
    -------
    pandas.DataFrame
        DataFrame s naformátovaným datumovým sloupcem.
    """

    if sloupec in df.columns:
        df[sloupec] = pd.to_datetime(
            df[sloupec],
            errors="coerce"
        ).dt.strftime("%d.%m.%Y")

    return df


def exportuj_kategorie(df, output_dir):
    """
    Rozdělí data podle kategorií a exportuje je
    do samostatných XLSX souborů.

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame obsahující sloupec 'kategorie'.

    output_dir : str
        Cílový adresář pro export souborů.

    Returns
    -------
    None

    Příklady vytvořených souborů
    ----------------------------
    f60_FILTER_LEKAR.xlsx
    f60_FILTER_PLACV.xlsx
    f60_FILTER_ABSEN.xlsx
    """

    pocet_souboru = 0

    for kategorie, data in df.groupby("kategorie"):

        logger.info(
            "Exportuji kategorii %s (%s řádků)",
            kategorie,
            len(data)
        )

        df_export = data.copy()

        df_export = formatuj_datum(df_export, DATUM_OD)
        df_export = formatuj_datum(df_export, DATUM_DO)

        vystupni_soubor = (
            Path(output_dir)
            / f"f60_FILTER_{kategorie}.xlsx"
        )

        try:
            df_export.to_excel(
                vystupni_soubor,
                index=False
            )

        except Exception:
            logger.exception(
                "Export kategorie %s selhal.",
                kategorie
            )
            raise

        if not vystupni_soubor.exists():
            raise RuntimeError(
                f"Soubor nebyl vytvořen: {vystupni_soubor}"
            )

        logger.info(
            "Vytvořen soubor: %s",
            vystupni_soubor
        )

        pocet_souboru += 1

    logger.info(
        "Export dokončen. Vytvořeno %s souborů.",
        pocet_souboru
    )

def validuj_droby(df):
    """
    Ověří, zda byly nalezeny hodnoty neo2dr,
    které nejsou definovány v MAP_KATEGORIE.

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame obsahující sloupce neo2dr a kategorie.

    Returns
    -------
    None

    Notes
    -----
    Pokud jsou nalezeny neznámé hodnoty, jsou
    zalogovány jako WARNING a zařazeny do
    kategorie DROBY.
    """

    maska_droby = df["kategorie"] == "DROBY"

    pocet_droby = maska_droby.sum()

    nezname_hodnoty = (
        df.loc[
            maska_droby,
            SLOUPEC_KATEGORIE
        ]
        .dropna()
        .unique()
    )

    if pocet_droby > 0:

        logger.warning(
            "Do kategorie DROBY bylo zařazeno %s záznamů.",
            pocet_droby
        )

        logger.warning(
            "Neznámé hodnoty neo2dr: %s",
            ", ".join(sorted(map(str, nezname_hodnoty)))
        )

    else:

        logger.info(
            "Všechny hodnoty neo2dr byly úspěšně zařazeny "
            "do definovaných kategorií."
        )

def main():
    """
    Hlavní vstupní bod aplikace.

    Provede:

    1. Načtení vstupního Excel souboru.
    2. Validaci povinných sloupců.
    3. Přiřazení kategorií podle neo2dr.
    4. Vytvoření výstupního adresáře.
    5. Export dat do samostatných souborů.
    6. Zápis průběhu zpracování do logu.

    Raises
    ------
    Exception
        Jakákoliv neošetřená chyba vzniklá během
        načítání, transformace nebo exportu dat.
    """

    logger.info("Spouštím zpracování F60")
    logger.info("Vstupní soubor: %s", ROZDELOVANY_SOUBOR)

    if not Path(ROZDELOVANY_SOUBOR).exists():
        raise FileNotFoundError(
            f"Vstupní soubor neexistuje: {ROZDELOVANY_SOUBOR}"
        )

    try:

        df = pd.read_excel(
            ROZDELOVANY_SOUBOR,
            engine="openpyxl"
        )

        if df.empty:
            raise ValueError(
                "Vstupní soubor neobsahuje žádná data."
            )

        logger.info(
            "Načteno %s řádků",
            len(df)
        )

        validuj_sloupce(df)

        df["kategorie"] = (
            df[SLOUPEC_KATEGORIE]
            .apply(prirad_kategorii)
        )

        validuj_droby(df)

        logger.info("Rozložení kategorií:")
        logger.info(
            "\n%s",
            df["kategorie"].value_counts()
        )

        Path(OUTPUT_DIR).mkdir(
            parents=True,
            exist_ok=True
        )

        exportuj_kategorie(df, OUTPUT_DIR)

        logger.info("Skript byl úspěšně dokončen.")

    except Exception:
        logger.exception(
            "Při zpracování došlo k neočekávané chybě."
        )
        raise


if __name__ == "__main__":
    main()