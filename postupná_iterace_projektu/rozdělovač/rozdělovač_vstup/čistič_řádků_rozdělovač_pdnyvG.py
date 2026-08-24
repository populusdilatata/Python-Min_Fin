import pandas as pd
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass
import logging

from time import perf_counter
import xlsxwriter


# Architektura               9,2/10 
# Čitelnost                  9,5/10 |
# Dokumentace                9,0/10 |
# Logging                    9,3/10 |
# Python styl                9,2/10 |
# Robustnost                 8,9/10 |
# Udržovatelnost             9,4/10 |

# ===== nastavení =====
@dataclass(frozen=True)
class Config:
    limit_oscis: int = 90000
    arbiter: int = 901099000
    fau: int = 905098000
    #===========================
    saturday: int = 5
    sunday: int = 6
    min_empty_priche_count: int = 25
    #===========================
    zpracovavany_mesic: str ="06_26"
    #=========================== 
    input_file: Path  = Path("vstup_pdnyv/pdnyv_06.xlsx")
    base_output_dir: Path  = Path("porovnavač")
CFG = Config()


@dataclass(frozen=True)
class Columns:
    den: str = "den"
    oscis: str = "oscis"
    pracvmes: str = "pracvmes"
    duvodt: str = "duvodt"
    priche: str = "priche"
    odche: str = "odche"
    kategorie: str = "kategorie"

COL = Columns()


HOLIDAYS = {
    1: [1],          # Nový rok
    5: [1, 8],       # Svátek práce, Den vítězství
    7: [5, 6],       # Cyril a Metoděj, Jan Hus
    9: [28],         # Den české státnosti
    10: [28],        # Vznik ČSR
    11: [17],        # Den boje za svobodu
    12: [24, 25, 26] # Vánoce
}


# ======================================================
# LOGGING
# ======================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)

logger = logging.getLogger(__name__)



EXCEL_EPOCH = datetime(1899, 12, 30)
# ===== sloupce =====
TIME_COLS = frozenset({
    "priche",
    "odche",
    "prichv",
    "odchv",
    "prichsm",
    "odchsm",
    "odprac",
    "odpracz",
    "dfond"
}
)

CATEGORY_MAP: dict[str, str] = {
    "dovol": "dovol",
    "puldo": "dovol",

    "odpra": "odpra",
    "služ": "odpra",

    "absen": "absen",
    "nemoc": "nemoc",
    "indiv": "indiv",
    "lékař": "lékař"
}



# ===== složky =====
def create_output_directory() -> tuple[Path, str]:
    
    """
    Vytvoří výstupní adresář pro aktuální běh programu.

    Název adresáře obsahuje datum a čas spuštění,
    což umožňuje uchovávat výsledky jednotlivých běhů odděleně.

    Returns:
        tuple[Path, str]:
            output_dir - cesta k výstupní složce
            time_suffix - čas použitý v názvech souborů
    """
    now = datetime.now()

    run_timestamp = now.strftime("%Y-%m-%d_%H-%M")
    time_suffix = now.strftime("%H_%M")

    base_output_dir = Path(CFG.base_output_dir)
    base_output_dir.mkdir(exist_ok=True)

    output_dir = base_output_dir / run_timestamp
    output_dir.mkdir(exist_ok=True)

    return output_dir, time_suffix


# ===== načtení =====
def load_data()-> pd.DataFrame:
    """
    Načte vstupní Excel soubor do DataFrame.

    Současně převede sloupec s datem na pandas datetime
    a neplatné hodnoty nahradí NaT.

    Returns:
        pd.DataFrame:
            Načtená a základně připravená data.
    """
    print(xlsxwriter.__version__)
    try:
        df = pd.read_excel(
            CFG.input_file,
            engine="openpyxl"
                        )
    except Exception:
        logger.exception(
                        "Nepodařilo se načíst %s",
                        CFG.input_file
                        )
        raise

    


    df[COL.den] = pd.to_datetime( df[COL.den], errors="coerce")
    return df

# ===== bezpečný převod času =====
def normalize_time_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Převede časové sloupce typu timedelta na datetime.

    Jako základ je použito datum EXCEL_EPOCH,
    ke kterému se přičítá časová složka.

    Args:
        df: Zdrojový DataFrame.

    Returns:
        pd.DataFrame:
            Upravený DataFrame.
    """
    for col in TIME_COLS:

        if col not in df.columns:
            continue

        
        if pd.api.types.is_timedelta64_dtype(df[col]):
            df[col] = EXCEL_EPOCH + df[col]


    return df

def apply_main_filter(df: pd.DataFrame) -> pd.DataFrame:    
    """
    Aplikuje hlavní filtr záznamů.

    Odstraňuje vybrané organizační jednotky
    a záznamy s OSCIS nad nastaveným limitem.

    Args:
        df: Zdrojový DataFrame.

    Returns:
        pd.DataFrame:
            Filtrovaný DataFrame.
    """
    filter_mask = (~df[COL.pracvmes].isin([
                                                CFG.arbiter,
                                                CFG.fau
                                                ])
                    ) & (
                        df[COL.oscis] < CFG.limit_oscis
                        )

    return df.loc[filter_mask]

def get_empty_priche(df: pd.DataFrame) -> pd.DataFrame:
   """
    Vybere záznamy bez vyplněného času příchodu.

    Pokud sloupec 'priche' ve vstupních datech
    neexistuje, vrací prázdný DataFrame.

    Args:
        df: Zdrojový DataFrame.

    Returns:
        pd.DataFrame:
            Záznamy s prázdným příchodem.
    """
   if COL.priche not in df.columns:
        return pd.DataFrame()
   return df.loc[df[COL.priche].isna()]


def has_valid_time(df_input: pd.DataFrame) -> pd.DataFrame:    
    """
    Odstraní záznamy bez příchodu nebo odchodu.

    Zachová pouze řádky, které obsahují
    vyplněné oba časové údaje.

    Args:
        df_input: Vstupní DataFrame.

    Returns:
        pd.DataFrame:
            Filtrovaná data.
    """
    return df_input.dropna(subset=[COL.priche, COL.odche])

def filter_frequent_empty_priche(df_prichv_empty: pd.DataFrame) -> pd.DataFrame:
    """
    Vybere zaměstnance s častým výskytem
    prázdného času příchodu.

    Ponechá pouze OSCIS, které se vyskytují
    vícekrát než určuje konfigurace.

    Args:
        df_prichv_empty: Záznamy s prázdným příchodem.

    Returns:
        pd.DataFrame:
            Vyfiltrovaná data.
    """
    if df_prichv_empty.empty:
        return pd.DataFrame()

    counts = (
        df_prichv_empty[COL.oscis]
        .value_counts()
    )

    valid_values = (
        counts[counts > CFG.min_empty_priche_count]
        .index
    )

    return df_prichv_empty.loc[
        df_prichv_empty[COL.oscis]
        .isin(valid_values)
    ]

def get_holidays(df: pd.DataFrame) -> pd.DataFrame:
    """
    Vyhledá záznamy spadající na státní svátky.

    Svátky jsou definovány ve slovníku HOLIDAYS,
    kde klíčem je měsíc a hodnotou seznam dní.

    Args:
        df: Zdrojový DataFrame.

    Returns:
        pd.DataFrame:
            Záznamy odpovídající svátkům.
    """



    holiday_dates = {
        (month, day)
        for month, days in HOLIDAYS.items()
        for day in days
    }

    
    mask = [
        (month, day) in holiday_dates
        for month, day in zip
            (
            df[COL.den].dt.month,
            df[COL.den].dt.day
            )
           ]                


    return df.loc[mask]

    
def get_weekends(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Rozdělí víkendové záznamy.

    Vrací samostatně soboty, neděle
    a jejich společné spojení.

    Args:
        df: Zdrojový DataFrame.

    Returns:
        tuple:
            (soboty, neděle, víkend)
    """
    df_saturday = df.loc[df[COL.den].dt.weekday == CFG.saturday]
    df_sunday = df.loc[df[COL.den].dt.weekday == CFG.sunday]

    df_vikend = pd.concat([
        df_saturday,
        df_sunday
    ])

    return (
        df_saturday,
        df_sunday,
        df_vikend
    )
# ===== export =====
def build_filename(name: str) -> str:
    """
    Vytvoří standardizovaný název exportního souboru.

    Všechny exporty používají jednotný formát
    názvu podle zpracovávaného období.

    Args:
        name: Název typu exportu.

    Returns:
        str:
            Název souboru bez přípony.
    """
    return f"pdnyv_vystup_{CFG.zpracovavany_mesic}_{name}"

def save_with_footer(df_input: pd.DataFrame, base_filename: str, output_dir: Path, time_suffix: str) -> None:
    """
    Uloží DataFrame do Excel souboru.

    Současně upraví formát datumových a časových
    sloupců tak, aby byly správně zobrazeny
    v Microsoft Excelu.

    Args:
        df_input: DataFrame k exportu.
        base_filename: Základ názvu souboru.
    """

    
    if df_input.empty:
            logger.warning( "Přeskočeno: %s",base_filename,)
            return

    df_copy = df_input.copy()

    df_copy[COL.den] = (
            df_copy[COL.den]
            .dt.strftime("%d.%m.%Y")
        )

    filename = f"{base_filename}_{time_suffix}.xlsx"
    filepath = output_dir / filename

    logger.info(                
                "Soubor %s obsahuje %d řádků a %d sloupců",                
                base_filename,
                len(df_copy),
                len(df_copy.columns)
            )
  
    t0 = perf_counter()
    with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:    

        df_copy.to_excel(
            writer,
            index=False
        )

        #ws = writer.book.active

        
        sheet_name = "Sheet1"
        ws = writer.sheets[sheet_name]

        """
        for col_idx, col_name in enumerate(df_copy.columns, 1):
            if col_name in TIME_COLS:
                #for row in range(2, ws.max_row + 1):
                max_row = len(df_copy)
                for row in range(2, max_row + 1):
                    ws.cell(
                            row=row,
                            column=col_idx
                            ).number_format = "HH:mm"
        """
        time_format = writer.book.add_format({'num_format': 'hh:mm'})

        for col_idx, col_name in enumerate(df_copy.columns):
            if col_name in TIME_COLS:
                ws.set_column(col_idx, col_idx, 12, time_format)
        
    t1 = perf_counter()
    logger.debug(
                "Export do Excelu: za %.2f s",
                t1 - t0,
                )
     

    t2 = perf_counter()
    logger.debug(
            "Formátování buněk za %.2f s",
            t2 - t1
                )


    t3 = perf_counter()
    logger.debug(
            "Uložení souboru za %.2f s",
            t3 - t2
                )
    

def export_weekends(df_saturday: pd.DataFrame, df_sunday: pd.DataFrame, output_dir: Path, time_suffix: str ) -> None:
    
    """
    Exportuje víkendové záznamy.

    Vytváří samostatný export sobot,
    neděl a jejich společný soubor.

    Args:
        df_saturday: Záznamy za soboty.
        df_sunday: Záznamy za neděle.
    """


    logger.info("Export víkendů")
    df_saturday_valid = has_valid_time(df_saturday)

    save_with_footer(
        df_saturday_valid,
        build_filename("soboty"),
        output_dir,
        time_suffix,
    )

    df_sunday_valid = has_valid_time(
        df_sunday
    )

    save_with_footer(
        df_sunday_valid,
        build_filename("nedele"),
        output_dir, 
        time_suffix,
    )

    df_vikend_valid = pd.concat(
        [
            df_saturday_valid,
            df_sunday_valid
        ]
    )

    save_with_footer(
        df_vikend_valid,
        build_filename("vikendy"), 
        output_dir, 
        time_suffix,
    )


def export_categories(df_clean: pd.DataFrame, output_dir: Path, time_suffix: str)-> None:
    """
    Rozdělí data podle kategorií důvodů.

    Pro každou nalezenou kategorii vytvoří
    samostatný exportní Excel soubor.

    Args:
        df_clean: DataFrame obsahující sloupec kategorie.
    """
    logger.info("Exporty rozdělených do kategorií")
    for name, df_part in (df_clean.groupby(COL.kategorie)):
        save_with_footer(
            df_part.drop(
                columns=[COL.kategorie]
            ),
            f"pdnyv_{CFG.zpracovavany_mesic}_FILTER_{name}", 
            output_dir, 
            time_suffix,
        )
def export_holidays(df_svatky: pd.DataFrame, output_dir: Path, time_suffix: str) -> None:
    """
    Exportuje záznamy spadající na státní svátky.

    Před exportem odstraní řádky,
    které nemají vyplněný příchod a odchod.

    Args:
        df_svatky: Záznamy za svátky.
    """

    logger.info("Export svátků")
    save_with_footer(
        has_valid_time(df_svatky),
        build_filename("svatky"), 
        output_dir,
        time_suffix,
    )


def export_empty_priche(df_prichv_empty_filtered: pd.DataFrame, output_dir: Path, time_suffix: str)-> None:
    """
    Exportuje zaměstnance s chybějícím příchodem.

    Do exportu jsou zahrnuti pouze uživatelé,
    kteří překročili nastavený limit výskytů.

    Args:
        df_prichv_empty_filtered:
            Vyfiltrované záznamy.
    """

    logger.info("Export prázdných příchodů")
    save_with_footer(
        df_prichv_empty_filtered,
        build_filename("FILTER_prichv_EMPTY"), output_dir, 
        time_suffix,
    )

def prepare_clean_dataframe(df: pd.DataFrame, df_vikend: pd.DataFrame, df_svatky: pd.DataFrame, df_prichv_empty_filtered: pd.DataFrame) -> pd.DataFrame:

    excluded_idx = (
        set(df_vikend.index)
        | set(df_svatky.index)
        | set(df_prichv_empty_filtered.index)
    )

    df_clean = df.drop(index=excluded_idx).copy()

    df_clean.loc[:, COL.kategorie] = (df_clean[COL.duvodt].map(CATEGORY_MAP).fillna("droby"))

    return df_clean

def main():
    """
    Hlavní řídicí funkce programu.

    Zajišťuje načtení dat, aplikaci filtrů,
    přípravu exportních datasetů a vytvoření
    všech výsledných Excel souborů.

    Program začíná i končí právě zde.

    
    """
                           
    output_dir, time_suffix = (create_output_directory())

    

    df = load_data()

    df = normalize_time_columns(df)

    df = apply_main_filter(df)
    
    df_prichv_empty = get_empty_priche(df)

    
    df_prichv_empty_filtered = (filter_frequent_empty_priche(df_prichv_empty))


    df_svatky = get_holidays(df)

    (
        df_saturday,
        df_sunday,
        df_vikend
    ) = get_weekends(df)

    logger.info("Exporty")
    start = perf_counter()
           
    export_empty_priche(df_prichv_empty_filtered, output_dir, time_suffix)
    df_clean = prepare_clean_dataframe(df, df_vikend, df_svatky, df_prichv_empty_filtered)

    export_holidays(df_svatky, output_dir, time_suffix)
    export_weekends(df_saturday, df_sunday, output_dir, time_suffix)

    export_categories(df_clean, output_dir, time_suffix)

    logger.info("Celkový čas exportu: %.2f s",
                perf_counter() - start)

    logger.info("Hotovo.")
    logger.info("Soubory jsou ve složce: %s",
                output_dir)


if __name__ == "__main__":
    main()
