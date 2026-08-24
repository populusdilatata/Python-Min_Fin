"""
Rozdělení vstupního Excel souboru na více výstupních souborů
na základě hodnoty ve sloupci prndr.
"""

from pathlib import Path
import logging
from typing import Dict, List

import pandas as pd
# Architektura               9,0/10 
# Čitelnost                  9,0/10 |
# Dokumentace                9,0/10 |
# Logging                    8,5/10 |
# Python styl                9,0/10 |
# Robustnost                 8,0/10 |
# Udržovatelnost             9,5/10 |

# ============================================================================
# KONFIGURACE
# ============================================================================

INPUT_FILE = Path("vstup_f64/07_DATA/f64A.xlsx")
OUTPUT_DIR = Path("vstup_f64")
BASE_FILENAME = "f64"

FILTER_COLUMN = "prndr"
DATE_COLUMNS = ["prnza", "prnko"]

FILTERS: Dict[str, List[int]] = {
    "NEMOC": [100, 101],
    "OTECD": [16],
    "OCR": [108],
}


# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

LOGGER = logging.getLogger(__name__)


# ============================================================================
# FUNKCE
# ============================================================================

def validate_columns(
    dataframe: pd.DataFrame,
    required_columns: list[str]
) -> None:
    """
    Ověří existenci požadovaných sloupců v DataFrame.

    Args:
        dataframe:
            Kontrolovaný DataFrame.

        required_columns:
            Seznam povinných sloupců.

    Raises:
        ValueError:
            Pokud některý sloupec chybí.
    """

    missing_columns = [
        column
        for column in required_columns
        if column not in dataframe.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Chybí požadované sloupce: {missing_columns}"
        )


def load_excel(file_path: Path) -> pd.DataFrame:
    """
    Načte vstupní Excel soubor.

    Args:
        file_path:
            Cesta ke zdrojovému souboru.

    Returns:
        Načtený DataFrame.
    """

    LOGGER.info("Načítám soubor: %s", file_path)

    dataframe = pd.read_excel(
        file_path,
        engine="openpyxl"
    )

    LOGGER.info(
        "Načteno %s řádků",
        len(dataframe)
    )

    return dataframe


def format_date_columns(
    dataframe: pd.DataFrame,
    columns: list[str]
) -> pd.DataFrame:
    """
    Převede datumové sloupce do formátu dd.mm.rrrr.

    Args:
        dataframe:
            DataFrame určený ke zpracování.

        columns:
            Seznam datumových sloupců.

    Returns:
        Upravený DataFrame.
    """

    dataframe = dataframe.copy()

    for column in columns:
        dataframe[column] = pd.to_datetime(
            dataframe[column],
            errors="coerce"
        )

        dataframe[column] = dataframe[column].dt.strftime(
            "%d.%m.%Y"
        )

    return dataframe


def save_dataframe(
    dataframe: pd.DataFrame,
    output_file: Path
) -> None:
    """
    Uloží DataFrame do Excel souboru.

    Args:
        dataframe:
            DataFrame určený k uložení.

        output_file:
            Výstupní soubor.
    """

    dataframe.to_excel(
        output_file,
        index=False
    )

    LOGGER.info(
        "Uložen soubor %s (%s řádků)",
        output_file,
        len(dataframe)
    )


def process_filter(
    source_dataframe: pd.DataFrame,
    filter_name: str,
    filter_values: list[int]
) -> None:
    """
    Vyfiltruje data podle definovaných hodnot
    a uloží výsledek do samostatného souboru.

    Args:
        source_dataframe:
            Zdrojový DataFrame.

        filter_name:
            Název filtru.

        filter_values:
            Hodnoty použité pro filtraci.
    """

    filtered_dataframe = source_dataframe[
        source_dataframe[FILTER_COLUMN].isin(filter_values)
    ].copy()

    LOGGER.info(
        "%s: nalezeno %s řádků",
        filter_name,
        len(filtered_dataframe)
    )

    filtered_dataframe = format_date_columns(
        filtered_dataframe,
        DATE_COLUMNS
    )

    output_file = (
        OUTPUT_DIR /
        f"{BASE_FILENAME}_FILTER_{filter_name}.xlsx"
    )

    save_dataframe(
        filtered_dataframe,
        output_file
    )


def main() -> None:
    """
    Hlavní vstupní bod aplikace.
    """

    LOGGER.info("Start zpracování")

    dataframe = load_excel(INPUT_FILE)

    validate_columns(
        dataframe,
        [FILTER_COLUMN, *DATE_COLUMNS]
    )

    for filter_name, filter_values in FILTERS.items():
        process_filter(
            dataframe,
            filter_name,
            filter_values
        )

    LOGGER.info("Zpracování dokončeno")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        LOGGER.exception("Program byl ukončen chybou")
        raise