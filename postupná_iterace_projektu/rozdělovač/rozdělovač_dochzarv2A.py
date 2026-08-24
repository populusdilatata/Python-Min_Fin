import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from pathlib import Path
from dataclasses import dataclass

# Architektura               7,0/10 
# Čitelnost                  8,0/10 |
# Dokumentace                5,0/10 |
# Logging                    3,0/10 |
# Python styl                7,0/10 |
# Robustnost                 6,0/10 |
# Udržovatelnost             8/10 |

# ===== nastavení =====
input_file = "vstup_dochzarv2/06_DATA/dochzarv_06B.xlsx"

@dataclass(frozen=True)
class Config:
    limit_oscis: int = 90000
    arbiter: int = 901099000
    fau: int = 905098000
    year = 2026
    month = 6
CFG = Config()
# ===== sloupce =====
@dataclass(frozen=True)
class Columns:
    start = "zapl"    
    end = "kopl"
    person_id = "oscis"
    work_place = "pracv"
    relation_to_org = "vztahorg"
COL = Columns()
# ===== další globální proměnné =====
time_cols = ["zapl", "kopl"]
ONE_DAY = pd.Timedelta(days=1)
# ===== složky =====
base_output_dir = Path("vstup_dochzarv2")

# ===== načtení =====
def load_data(input_file: str) -> pd.DataFrame:
    df = pd.read_excel(input_file, engine="openpyxl")
    df[COL.start] = pd.to_datetime(df[COL.start], format="%d.%m.%Y", errors="coerce", )

    return df

# ===== filtr =====
def apply_filters(df : pd.DataFrame, min_start: pd.Timestamp, max_end: pd.Timestamp) -> pd.DataFrame:
    # Odfiltruj nevyhovující záznamy podle typu pracovního vztahu,
    # hodnoty OSCIS 
    # data zániku pracovního poměru.

    valid_pracv = ~df[COL.work_place].isin([CFG.arbiter, CFG.fau])
    valid_oscis = df[COL.person_id] < CFG.limit_oscis
    valid_date = (df[COL.start].isna()| 
              (df[COL.start] <= max_end)
             )

    df = df[valid_pracv & valid_oscis & valid_date].copy()
    # Omez interval na sledované období:
    # - datum začátku posuň nejdříve na MIN_START
    # - chybějící datum konce nahraď MAX_END
    # - datum konce posuň nejpozději na MAX_END
    
    df[COL.start] = df[COL.start].clip(lower=min_start)

    df[COL.end] = (df[COL.end].fillna(max_end).clip(upper=max_end))

    return df


def split_by_relation(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    # ===== filtr vztahorg = 3 =====
    #if COL.relation_to_org in df_spojene.columns:
    #df_vztahorg_FILTER = df_spojene[df_spojene[COL.relation_to_org]==3]
    #df_vztahorg_OSTATNI = df_spojene[df_spojene[COL.relation_to_org]!=3]

    if COL.relation_to_org not in df.columns:
        return (
            pd.DataFrame(),
            pd.DataFrame(),
        )

    return (
        df[df[COL.relation_to_org] == 3],
        df[df[COL.relation_to_org] != 3],
    )


def find_duplicates(df: pd.DataFrame) -> pd.DataFrame:

    return df[
        df.duplicated(
            subset=[COL.person_id],
            keep=False,
                     )
             ]



def spojit_zaznamy(df: pd.DataFrame, potvrdit) ->pd.DataFrame:
    df = df.sort_values([COL.person_id, COL.start]).copy()

    for oscis, skupina in df.groupby(COL.person_id):

        idxs = list(skupina.index)

        # Vytvoř dvojice po sobě jdoucích indexů (aktuální, následující)
        for idx1, idx2 in zip(idxs, idxs[1:]):
        
        
            od1 = df.loc[idx1, COL.start]
            do1 = df.loc[idx1, COL.end]

            od2 = df.loc[idx2, COL.start]
            do2 = df.loc[idx2, COL.end]

            if pd.isna(do1) or pd.isna(od2):
                continue

            if do1 + ONE_DAY == od2:

                print("\n--------------------------------")
                print(f"Osobní číslo: {oscis}")
                print(f"1. {od1.date()} - {do1.date()}")
                print(f"2. {od2.date()} - {do2.date()}")
                print("--------------------------------")

                if not potvrdit():
                    return df

    return df


def potvrdit():
    return input("Spojit záznamy? (S/N): ").upper() == "S"


# ===== export =====
def save_with_footer(df_input: pd.DataFrame, base_filename: str, output_dir: Path, time_suffix: str) ->None:

    if df_input.empty:
        print(f"Přeskočeno: {base_filename}")
        return

    filepath = output_dir / f"{base_filename}_{time_suffix}.xlsx"

    with pd.ExcelWriter(filepath, engine="openpyxl",) as writer:

        df_input.to_excel(
            writer, index=False, sheet_name="Data",
                         )

        ws = writer.sheets["Data"]

        for col_idx, col_name in enumerate( df_input.columns, start=1,):
            if col_name not in time_cols:
                continue

            for column in ws.iter_cols(
                min_col=col_idx,
                max_col=col_idx,
                min_row=2,
            ):
                for cell in column:
                    cell.number_format = "dd.mm.yyyy"

    
def main():
    
    
    min_start = pd.Timestamp(CFG.year, CFG.month, 1)
    max_end = min_start + pd.offsets.MonthEnd(0)

    
    output_dir = (base_output_dir / datetime.now().strftime("%Y-%m-%d_%H-%M"))
    output_dir.mkdir(parents=True, exist_ok=True)

    time_suffix = datetime.now().strftime("%H_%M")

    print("Načítání dat...")
    df = load_data(input_file)

    print("Filtrování...")
    df = apply_filters(df, min_start, max_end)

    
    print("Spojování...")
    df_spojene = spojit_zaznamy(
        df,
        lambda: True,
    )

    duplicity = find_duplicates( df_spojene)
    vztahorg_3, vztahorg_ostatni = (split_by_relation(df_spojene))

    # ===== exporty =====
    print("Export filtrovaných záznamů")
    exports = [
        (duplicity, "duplicity"),
        (vztahorg_3, "FILTER_vztahorg"),
        (vztahorg_ostatni, "FILTER_vztahorg_NOT3",
        ),
               ]

    zpracovavany_mesic = f"{CFG.month:02}_{str(CFG.year)[2:]}"
    for frame, suffix in exports:
        save_with_footer(
            frame,
            f"dochzarv2_{zpracovavany_mesic}_{suffix}",
            output_dir,
            time_suffix,
        )


    # ===== výpis =====
    print()
    print("Hotovo.")
    print(f"Soubory jsou ve složce: {output_dir}")

    
if __name__ == "__main__":
    main()

