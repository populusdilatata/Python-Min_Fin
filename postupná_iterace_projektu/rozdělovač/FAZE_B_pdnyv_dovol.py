import pandas as pd
from pathlib import Path

import logging


# Architektura               9,5/10
# Čitelnost                  9,6/10
# Dokumentace                9,5/10
# Logging                    9.2/10
# Python styl                9,3/10
# Robustnost                 9,2/10
# Udržovatelnost             9,6/10

"""
Porovnávač exportů docházky.

Účel:
    Sloučí data ze dvou excelových souborů do jednoho výstupu.

Postup:
    1. Načte první soubor a ponechá pouze řádky,
       kde je vyplněn sloupec 'duvod2t'.
    2. Načte druhý soubor a ponechá pouze řádky,
       kde má sloupec 'duvodt' hodnotu 'puldo'.
    3. Oba datasety spojí.
    4. Výsledek seřadí podle sloupce 'oscis'.
    5. Uloží do nového XLSX souboru.

Vstupy:
    pdnyv_06_26_FILTER_odpra_11_12.xlsx
    pdnyv_06_26_FILTER_dovol_11_12.xlsx

Výstup:
    pdnyv_06_26_FILTER_PULDOVOL_11_12.xlsx

Autor:
    Majda Tomáš

Verze:
    2.0

Datum vytvoření:
    2026-07-14

Požadavky:
    pandas
    openpyxl

"""
# ===== Konfigurace vstupních souborů =====

INPUT_FILE_1 = "porovnavač/2026-08-05_11-46/pdnyv_07_26_FILTER_odpra_11_46.xlsx"
# Sloupec obsahující důvod nepřítomnosti z prvního exportu
ZADANY_SLOUPEC_1 = "duvod2t"

INPUT_FILE_2 = "porovnavač/2026-08-05_11-46/pdnyv_07A_26_FILTER_dovol_11_46.xlsx"
# Sloupec obsahující důvod nepřítomnosti
ZADANY_SLOUPEC_2 = "duvodt"
# Osobní číslo zaměstnance
ZADANY_SLOUPEC_3 = "oscis"

OUTPUT_DIR = "porovnavač/2026-08-04_11-46"
CATEGORY = "PULDOVOL"
FILTER_VALUE = "puldo"

# ===== Logging =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    
    handlers=[
        logging.FileHandler(
            "porovnavac.log", 
            encoding="utf-8"
            ),
        logging.StreamHandler()
    ]

)

# ===== Pomocné funkce =====
def over_sloupec(
    df: pd.DataFrame,
    sloupec: str
) -> None:
    """
    Ověří existenci sloupce v DataFrame.
    """

    if sloupec not in df.columns:
        raise ValueError(
            f"Chybí požadovaný sloupec: '{sloupec}'"
        )


def nacti_excel(cesta: str) -> pd.DataFrame:
    
    """
    Načte Excel soubor.
    """

    logging.info("Načítám soubor %s", cesta)

    
    if not Path(cesta).exists():
        
        logging.error(
        "Soubor neexistuje: %s",
        cesta
        )

        raise FileNotFoundError(
        f"Soubor neexistuje: {cesta}"
        )


    try:
        df = pd.read_excel(cesta, engine="openpyxl")
    except Exception as e:
        logging.error("Chyba při načítání %s: %s", cesta, e)
        raise

    logging.info("Načteno %d řádků", len(df))

    return df

def filtr_notna(df: pd.DataFrame, sloupec: str) -> pd.DataFrame:
        
    """
    Ponechá pouze řádky s vyplněnou hodnotou.
    """

    over_sloupec(df, sloupec)

    vysledek = df[df[sloupec].notna()]

    logging.info(
        "Filtr NOT NULL ve sloupci '%s' vrátil %d řádků",
        sloupec,
        len(vysledek)
    )

    
    if vysledek.empty:
        logging.warning(
            "Filtr NOT NULL nevrátil žádné záznamy."
        )


    return vysledek

def filtr_hodnota(df: pd.DataFrame, sloupec: str, hodnota: str) -> pd.DataFrame:
    
    """
    Filtruje konkrétní hodnotu.
    """

    over_sloupec(df, sloupec)

    vysledek = df[df[sloupec] == hodnota]

    logging.info(
        "Filtr '%s = %s' vrátil %d řádků",
        sloupec,
        hodnota,
        len(vysledek)
    )

    
    if vysledek.empty:
        logging.warning(
            "Filtr '%s = %s' nevrátil žádné záznamy.",
            sloupec,
            hodnota
        )


    return vysledek

def sluc_datasety(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    """
    Sloučí dva datasety.
    """

    vysledek = pd.concat([df1, df2], ignore_index=True)

    logging.info(
        "Výsledný dataset obsahuje %d řádků",
        len(vysledek)
    )

    return vysledek

def uloz_vysledek(df: pd.DataFrame, vystupni_cesta: Path, radici_sloupec: str) -> None:
    """
    Seřadí a uloží výsledek.
    """

    over_sloupec(df, radici_sloupec)

    
    logging.info(
        "Řadím podle sloupce '%s'",
        radici_sloupec
    )


    df_sorted = df.sort_values(by=radici_sloupec)

    df_sorted.to_excel(vystupni_cesta, index=False)

    logging.info(
        "Výsledek uložen do %s",
        vystupni_cesta
    )


def main() -> None:
    
    """
    Hlavní řídicí funkce programu.
    """


    df1 = nacti_excel(INPUT_FILE_1)   
    df1 = filtr_notna(
        df1,
        ZADANY_SLOUPEC_1
    )

    df2 = nacti_excel(INPUT_FILE_2)    
    df2 = filtr_hodnota(
        df2,
        ZADANY_SLOUPEC_2,
        FILTER_VALUE 
    )


    vysledek = sluc_datasety(df1, df2)
    
    filename = f"pdnyv_06_26_FILTER_{CATEGORY}_11_12.xlsx" 
    Path(OUTPUT_DIR).mkdir(
        parents=True,
        exist_ok=True
    ) 
    path = Path(OUTPUT_DIR) / filename    

    uloz_vysledek(
        vysledek,
        path,
        ZADANY_SLOUPEC_3
    )

    
if __name__ == "__main__":
    main()
