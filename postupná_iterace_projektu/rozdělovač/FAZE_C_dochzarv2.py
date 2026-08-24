import logging
import pandas as pd
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass

# Architektura               9,5/10 
# Čitelnost                  9,4/10 |
# Dokumentace                9,4/10 |
# Logging                    9,7/10 |
# Python styl                9,3/10 |
# Robustnost                 9,5/10 |
# Udržovatelnost             9,7/10 |

# ===== nastavení =====
INPUT_FILE: Path = Path("vstup_dochzarv2/07_DATA/dochzarv_07A.xlsx")
@dataclass(frozen=True)
class Config:
    limit_oscis: int = 90000
    arbiter: int = 901099000
    fau: int = 905098000
    year: int = 2026
    month: int = 7
    relation_filter: int = 3
CFG = Config()
# ===== sloupce =====
@dataclass(frozen=True)
class Columns:
    start: str = "zapl"    
    end: str = "kopl"
    person_id: str = "oscis"
    work_place: str = "pracv"
    relation_to_org: str = "vztahorg"
COL = Columns()
# ===== další globální proměnné =====
TIME_COLS: list[str] = ["zapl", "kopl"]
ONE_DAY: pd.Timedelta = pd.Timedelta(days=1)

# ===== logging =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)

logger: logging.Logger = logging.getLogger(__name__)

# ===== složky =====
BASE_OUTPUT_DIR: Path = Path("vstup_dochzarv2")

def validate_data_quality(df: pd.DataFrame) -> None:
    
    """
    Ověří základní kvalitu vstupních dat.

    Kontroluje:

    - chybějící datum nástupu,
    - otevřené pracovní poměry,
    - intervaly s koncem před začátkem,
    - neplatné hodnoty OSCIS.
    """

    missing_start = df[COL.start].isna().sum()

    if missing_start:
        logger.warning(
            "Chybí datum nástupu u %s záznamů",
            missing_start,
        )

    missing_end = df[COL.end].isna().sum()

    if missing_end:
        logger.info(
            "Otevřený pracovní poměr u %s záznamů",
            missing_end,
        )

    invalid_intervals = (
        df[COL.start].notna()
        & df[COL.end].notna()
        & (df[COL.end] < df[COL.start])
    )

    invalid_count = invalid_intervals.sum()

    if invalid_count:
        logger.warning(
            "Nalezeno %s intervalů s koncem před začátkem",
            invalid_count,
        )
    
    invalid_oscis = (df[COL.person_id] <= 0).sum()

    if invalid_oscis:
        logger.warning(
            "Nalezeno %s neplatných OSCIS",
            invalid_oscis,
        )

   

def validate_columns(df: pd.DataFrame) -> None:
    
    """
    Ověří strukturu vstupních dat.

    Kontroluje přítomnost všech sloupců
    potřebných pro další zpracování.

    Parametry:
        df:
            DataFrame načtený ze vstupního
            Excel souboru.

    Výjimky:
        ValueError:
            Pokud chybí alespoň jeden
            povinný sloupec.
    """

    required_columns: set[str] = {
        COL.start,
        COL.end,
        COL.person_id,
        COL.work_place,
        COL.relation_to_org,
    }

    missing = required_columns - set(df.columns)

    if missing:
        
        logger.error(
                "Chybí povinné sloupce: %s",
                ", ".join(sorted(missing))
                    )

        raise ValueError(
            f"Chybí sloupce: {', '.join(sorted(missing))}"
        )
    logger.info(
        "Nalezeno všech %s povinných sloupců",
        len(required_columns),
                )


# ===== načtení =====
def load_data(input_file: Path) -> pd.DataFrame:
    
    
    """
    Načte vstupní soubor docházky.

    Parametry:
        input_file:
            Cesta ke zdrojovému Excel souboru.

    Návratová hodnota:
        DataFrame se vstupními daty.
    """


    try:
        df = pd.read_excel(
            input_file,
            engine="openpyxl",
        )

    except FileNotFoundError:
        logger.error(
            "Soubor nebyl nalezen: %s",
            input_file,
        )
        raise

    except Exception:
        logger.exception(
            "Neočekávaná chyba při načítání %s",
            input_file,
        )
        raise

    validate_columns(df)
    


    df[COL.start] = pd.to_datetime(
        df[COL.start],
        format="%d.%m.%Y",
        errors="coerce",
    )

    df[COL.end] = pd.to_datetime(
        df[COL.end],
        format="%d.%m.%Y",
        errors="coerce",
    )

    validate_data_quality(df)

    logger.info(
        "Načteno %s záznamů",
        len(df),
    )

    return df

# ===== filtr =====
def apply_filters(df : pd.DataFrame, min_start: pd.Timestamp, max_end: pd.Timestamp) -> pd.DataFrame:
    
    """
    Aplikuje základní filtry.

    Odstraňuje:
    - záznamy pracovišť Arbiter a FAÚ,
    - OSCIS >= limit_oscis,
    - záznamy začínající po sledovaném období.

    Současně omezí interval pracovního poměru
    na sledovaný kalendářní měsíc.
    """


    valid_pracv = ~df[COL.work_place].isin([CFG.arbiter, CFG.fau])
    valid_oscis = df[COL.person_id] < CFG.limit_oscis
    valid_date = (df[COL.start].isna()| 
              (df[COL.start] <= max_end)
             )

    before = len(df)

    df = df[
            valid_pracv &
            valid_oscis &
            valid_date
            ].copy()
    
    df[COL.start] = df[COL.start].clip(lower=min_start)

    df[COL.end] = (df[COL.end].fillna(max_end).clip(upper=max_end))

    after = len(df)

    logger.info(
        "Filtrace odstranila %s záznamů (%s → %s)",
        before - after,
        before,
        after,
    )
    return df



def split_by_relation(df: pd.DataFrame, relation_value: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    

    """
    Rozdělí data podle hodnoty vztahorg.

    Parametry:
        relation_value:
            Hodnota vztahorg použitá
            pro rozdělení datasetu.

    Vrací dvojici:
        (
            vztahorg == relation_value,
            vztahorg != relation_value,
        )
    """

   
    mask = (
        df[COL.relation_to_org]
        == relation_value
    )

    
    matching = df[mask].copy()
    others = df[~mask].copy()

    logger.info(
        "Rozdělení vztahorg: %s = %s řádků, ostatní = %s řádků",
        relation_value,
        len(matching),
        len(others),
                )


    return (
        matching,
        others,
    )



def find_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    
    """
    Vyhledá zaměstnance,
    kteří mají více než jeden záznam.

    Duplicity jsou určeny
    podle osobního čísla (OSCIS).
    """

    duplicates = df[
                df.duplicated(
                subset=[COL.person_id],
                keep=False,
                     )
                ]
    
    logger.info(
            "Nalezeno %s duplicitních záznamů",
            len(duplicates),
                )

    return duplicates




def spojit_zaznamy(df: pd.DataFrame) -> pd.DataFrame:
    
    """
    Spojuje navazující intervaly
    pracovních poměrů stejného zaměstnance.

    Pokud následující interval začíná
    bezprostředně po skončení předchozího,
    budou oba intervaly sloučeny do jednoho.

    """
    
    if df.empty:
            logger.info(
            "Spojování přeskočeno - dataset je prázdný"
                )
            return df.copy()

    
    
    before = len(df)
    spojeno = 0


    df = df.sort_values(
        [COL.person_id, COL.start]
    ).copy()

    vysledek = []

    for oscis, skupina in df.groupby(COL.person_id):

        aktualni = skupina.iloc[0].copy()

        for _, radek in skupina.iloc[1:].iterrows():

            navazuje = (
                aktualni[COL.end] + ONE_DAY
                == radek[COL.start]
            )

            if navazuje:
                spojeno += 1
                aktualni[COL.end] = max(
                    aktualni[COL.end],
                    radek[COL.end]
                )
            else:
                vysledek.append(aktualni)
                aktualni = radek.copy()

        vysledek.append(aktualni)
    
    result = pd.DataFrame(vysledek)
    
    logger.info(
        "Spojování dokončeno (%s → %s řádků, %s spojení)",
        before,
        len(result),
        spojeno,
    )


    return result


# ===== export =====
def save_excel(df_input: pd.DataFrame, base_filename: str, output_dir: Path, time_suffix: str) ->None:

    """
    Uloží DataFrame do Excelu.

    Pro datumové sloupce nastaví
    formát dd.mm.yyyy.
    Prázdné datasety neexportuje.
    """

    if df_input.empty:
        logger.warning(
            "Export přeskočen - dataset je prázdný (%s)",
            base_filename,
        )

        return

    filepath = output_dir / f"{base_filename}_{time_suffix}.xlsx"

    logger.info(
        "Exportuji %s (%s řádků)",
        filepath.name,
        len(df_input),
        )

    with pd.ExcelWriter(filepath, engine="openpyxl",) as writer:

        df_input.to_excel(
            writer, index=False, sheet_name="Data",
                         )

        ws = writer.sheets["Data"]

        for col_idx, col_name in enumerate( df_input.columns, start=1,):
            if col_name not in TIME_COLS:
                continue

            for column in ws.iter_cols(
                min_col=col_idx,
                max_col=col_idx,
                min_row=2,
            ):
                for cell in column:
                    cell.number_format = "dd.mm.yyyy"

    logger.info(
        "Export dokončen: %s",
        filepath,
        )

    
def main():
    
    """
    Hlavní řídicí funkce.

    Provádí:
    1. Načtení dat.
    2. Filtraci.
    3. Spojování intervalů.
    4. Vyhledání duplicit.
    5. Rozdělení podle vztahorg.
    6. Export výsledků.
    """    
    
    min_start = pd.Timestamp(CFG.year, CFG.month, 1)
    max_end = min_start + pd.offsets.MonthEnd(0)

    
    output_dir = (BASE_OUTPUT_DIR / datetime.now().strftime("%Y-%m-%d_%H-%M"))
    output_dir.mkdir(parents=True, exist_ok=True)

    time_suffix = datetime.now().strftime("%H_%M")

    logger.info("Načítání dat...")
    df = load_data(INPUT_FILE)

    logger.info("Filtrování...")
    df = apply_filters(df, min_start, max_end)

    
    logger.info("Spojování intervalů...")
    df_spojene = spojit_zaznamy(
        df
    )

    duplicity = find_duplicates( df_spojene)
    vztahorg_3, vztahorg_ostatni = (split_by_relation(df_spojene, CFG.relation_filter))

    # ===== exporty =====
    logger.info("Export filtrovaných záznamů")

    exports = [
        (duplicity, "duplicity"),
        (vztahorg_3, "FILTER_vztahorg"),
        (vztahorg_ostatni, "FILTER_vztahorg_NOT3",
        ),
               ]

    zpracovavany_mesic = f"{CFG.month:02}_{str(CFG.year)[2:]}"
    for frame, suffix in exports:
        save_excel(
            frame,
            f"dochzarv2_{zpracovavany_mesic}_{suffix}",
            output_dir,
            time_suffix,
        )


    # ===== výpis =====
    
    logger.info("===== Souhrn =====")
    logger.info("Po filtraci: %s řádků", len(df))
    logger.info("Duplicity: %s řádků", len(duplicity))
    logger.info("Vztahorg=%s: %s řádků",
                CFG.relation_filter,
                len(vztahorg_3))
    logger.info("Vztahorg!=%s: %s řádků",
                CFG.relation_filter,
                len(vztahorg_ostatni))

    
    logger.info("Hotovo")
    logger.info("Soubory jsou ve složce: %s", output_dir)


    
if __name__ == "__main__":
    main()

