import calendar
import pandas as pd
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass
import logging

from time import perf_counter


# Architektura               9,5/10 
# Čitelnost                  9,7/10 |
# Dokumentace                9,1/10 |
# Logging                    9,3/10 |
# Python styl                9,4/10 |
# Robustnost                 9,5/10 |
# Udržovatelnost             9,6/10 |

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
    #=========================== 
    zpracovavany_rok: int = 2026
    zpracovavany_mesic_cislo: int = 6
CFG = Config()


@dataclass(frozen=True)
class Columns:
    den: str = "den"
    oscis: str = "oscis"
    prijm: str = "prijm"
    pracvmes: str = "pracvmes"
    duvodt: str = "duvodt"
    priche: str = "priche"
    odche: str = "odche"
    kategorie: str = "kategorie"
    dt: str = "dt"


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
VALID_DT_VALUES = {
    "Po",
    "Út",
    "St",
    "Čt",
    "Pá",
    "So",
    "Ne",
}

WEEKDAY_MAP = {
    0: "Po",
    1: "Út",
    2: "St",
    3: "Čt",
    4: "Pá",
    5: "So",
    6: "Ne",
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



EXCEL_EPOCH = datetime(2000, 1, 1)
# ===== sloupce =====
DESIRED_COLS = [
    "oscis",
    "den",
    "prijm",
    "dt",
    "duvod",
    "duvodt",
    "duvod2",
    "duvod2t",
    "pracvmes",
    "priche",
    "odche"
]
TIME_COLS = frozenset({
    "priche",
    "odche",
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

REQUIRED_COLUMNS = set(DESIRED_COLS)


def validate_required_columns(df: pd.DataFrame) -> None:
    """
    Ověří přítomnost všech povinných sloupců.
    """
    missing = REQUIRED_COLUMNS - set(df.columns)

    if missing:
        raise ValueError(
            "Ve vstupním souboru chybí povinné sloupce: "
            f"{', '.join(sorted(missing))}"
        )


def keep_desired_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Vrátí DataFrame obsahující pouze sloupce definované v DESIRED_COLS."""
    cols = [col for col in df.columns if col in DESIRED_COLS]
    return df[cols]


def validate_oscis(df: pd.DataFrame) -> None:
    """
    Ověří, že OSCIS obsahuje pouze pětimístná čísla.
    """
    
    invalid_rows = df.loc[
        ~df[COL.oscis].between(10000, 99999)
    ]

    if not invalid_rows.empty:
        logger.warning(
            "Nalezeno %d neplatných OSCIS.",
            len(invalid_rows)
        )

        logger.warning(
            "Neplatné hodnoty OSCIS: %s",
            sorted(invalid_rows[COL.oscis].unique())
        )


def validate_dates_in_month(df: pd.DataFrame) -> None:
    """
    Ověří, že všechna data spadají
    do zadaného měsíce.
    """

    last_day = calendar.monthrange(
        CFG.zpracovavany_rok,
        CFG.zpracovavany_mesic_cislo
    )[1]

    date_from = pd.Timestamp(
        CFG.zpracovavany_rok,
        CFG.zpracovavany_mesic_cislo,
        1
    )

    date_to = pd.Timestamp(
        CFG.zpracovavany_rok,
        CFG.zpracovavany_mesic_cislo,
        last_day
    )

    invalid_dates = df.loc[
        (df[COL.den] < date_from)
        |
        (df[COL.den] > date_to)
    ]

    if not invalid_dates.empty:

        logger.warning(
            "Nalezeno %d záznamů mimo povolené období.",
            len(invalid_dates)
        )

        logger.warning(
            "Neplatná data: %s",
            sorted(
                invalid_dates[COL.den]
                .dt.strftime("%d.%m.%Y")
                .unique()
            )
        )


def validate_dt_column(df: pd.DataFrame) -> None:
    """
    Ověří obsah sloupce DT.
    """

    invalid_values = (
        df[COL.dt]
        .dropna()
        .loc[
            ~df[COL.dt].isin(VALID_DT_VALUES)
        ]
        .unique()
    )

    if len(invalid_values) > 0:

        logger.warning(
            "Sloupec DT obsahuje nepovolené hodnoty: %s",
            sorted(map(str, invalid_values))
        )

def validate_string(df):
    counter=0

    for idx, row in df.iterrows():
        
        if pd.isna(row["pracvmes"]):
            #logging.warning(
            #"Řádek %s neprošel validací - chybí pracvmes",
            #idx
            #)
            counter=counter+1
            continue

        text = str(int(row["pracvmes"]))

        if len(text) != 9:
            
            logging.warning(
                "Řádek %s neprošel validací, pracvmes=%s, PROTOŽE DÉLKA NESEDÍ",
                idx,
                text
            )
            continue


        if not text.startswith("9"):
            logging.warning(
                "Řádek %s neprošel validací, pracvmes=%s, NEZAČÍNÁ DEVÍTKOU",
                idx,
                text
            )
            continue

        kombinace = text[1:4]
        nasledujici = text[4:7]

        povolene = {
            "010": {"740", "760", "900", "990", "001", "000"},
            "020": {"130", "590", "700", "900", "000"},

            "030": {"290", "410", "720", "780", "790", "000"},
            "040": {"470", "520", "170", "260", "750", "521", "000"},
            "050": {"320", "390", "150", "180", "240", "280", "730", "260", "980", "000"},
            "060": {"370", "690", "110", "120", "140", "190","200", "000"},
            "070": {"350", "360", "530", "550", "580", "270", "000"},
            "080": {"300", "660", "230", "900", "000"},


        }

        posledni=text[-1]

        if kombinace in povolene and nasledujici in povolene[kombinace] and posledni in {"0", "8", "9", "1"}:
            continue
        else:
            logging.warning(
                "Řádek %s neprošel validací, pracvmes=%s",
                idx,
                text
                )
            
    logging.warning(
            "Celkem %d neprošel validací - chybí pracvmes",
            counter
        )

def validate_input_data(df: pd.DataFrame) -> None:
    """
    Spustí všechny vstupní validace.
    """

    validate_required_columns(df)

    validate_oscis(df)

    validate_dates_in_month(df)

    validate_dt_column(df)

    validate_date_dt_consistency(df)

    validate_string(df)

def validate_date_dt_consistency(df: pd.DataFrame) -> None:
    """
    Ověří, že sloupec DT odpovídá skutečnému dni
    vypočtenému ze sloupce DEN.
    """

    expected_dt = (
        df[COL.den]
        .dt.weekday
        .map(WEEKDAY_MAP)
    )

    invalid_rows = df.loc[
        expected_dt != df[COL.dt]
    ]

    if invalid_rows.empty:
        logger.info(
            "Kontrola DEN × DT: bez nesrovnalostí."
        )
        return

    logger.warning(
        "Nalezeno %d neshod mezi DEN a DT.",
        len(invalid_rows)
    )

    for _, row in invalid_rows.iterrows():

        expected = WEEKDAY_MAP[
            row[COL.den].weekday()
        ]

        logger.warning(
            "Řádek %s | den=%s | dt=%s | očekáváno=%s",
            row.name,
            row[COL.den].strftime("%d.%m.%Y"),
            row[COL.dt],
            expected
        )

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

    Returns:f
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

    """
    mask = [
        (month, day) in holiday_dates
        for month, day in zip
            (
            df[COL.den].dt.month,
            df[COL.den].dt.day
            )
           ]                
    """

    mask = (pd.MultiIndex.from_arrays([df[COL.den].dt.month, df[COL.den].dt.day]).isin(holiday_dates))

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

def save_excel(df_input: pd.DataFrame, base_filename: str, output_dir: Path, time_suffix: str) -> None:
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
    with pd.ExcelWriter(filepath, engine="xlsxwriter", datetime_format="hh:mm",) as writer:    

        df_copy.to_excel(
            writer,
            index=False
        )
        
        
    t1 = perf_counter()
    logger.debug(
                "Export do Excelu: za %.2f s",
                t1 - t0,
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

    save_excel(
        df_saturday_valid,
        build_filename("soboty"),
        output_dir,
        time_suffix,
    )

    df_sunday_valid = has_valid_time(
        df_sunday
    )

    save_excel(
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

    save_excel(
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
    logger.info("Exporty rozdělené do kategorií")
    for name, df_part in (df_clean.groupby(COL.kategorie)):
        save_excel(
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
    save_excel(
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
    save_excel(
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

    validate_input_data(df)

    df_desired = keep_desired_cols(df)

    df = normalize_time_columns(df_desired)

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
