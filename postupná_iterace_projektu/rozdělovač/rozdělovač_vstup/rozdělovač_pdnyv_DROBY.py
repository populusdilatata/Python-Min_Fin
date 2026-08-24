import pandas as pd 
import os

ROZDELOVANY_SOUBOR = "vystupy/2026-07-03_09-46/pdnyv_meziclen_05_droby_09_46.xlsx"
output_dir = "vystupy/2026-07-03_09-46/"
base_filename="pdnyv_meziclen_05_droby"
SLOUPEC = "duvodt"
ZACHOVANE_SLOUPCE = ["oscis","den","prijm", "pracvmes", "duvod", "duvodt"]

# --- načtení ---
df_soubor = pd.read_excel(ROZDELOVANY_SOUBOR, engine="openpyxl")
#print(df_soubor)
# --- rozdělení dle duvodt ---
# --- duvodt = OCR : očr
# --- duvodt = STUDV: STUDV
# --- duvodt = OTECD : OTECD
# --- duvodt = SVZOZ : SVZOZ
# --- duvodt = PLACV : PLACV
# --- duvodt = CERNV : CERNV
# --- duvodt = HOMOF : homof


df_FILTER_OCR = df_soubor[df_soubor[SLOUPEC] == "očr"]
df_FILTER_STUDV = df_soubor[df_soubor[SLOUPEC] == "studv"]
df_FILTER_OTECD = df_soubor[df_soubor[SLOUPEC] == "otecd"]
df_FILTER_SVZOZ = df_soubor[df_soubor[SLOUPEC] == "svzoz"]
df_FILTER_PLACV= df_soubor[df_soubor[SLOUPEC] == "placv"]
df_FILTER_CERNV = df_soubor[df_soubor[SLOUPEC] == "cernv"]
df_FILTER_HOMOF= df_soubor[df_soubor[SLOUPEC] == "homof"]

# --- uložení do jednotlivých souborů ---
filter_suffix="OCR"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_OCR = df_FILTER_OCR[ZACHOVANE_SLOUPCE]
df_FILTER_OCR.to_excel(filepath, index=False)

filter_suffix="STUDV"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_STUDV = df_FILTER_STUDV[ZACHOVANE_SLOUPCE]
df_FILTER_STUDV.to_excel(filepath, index=False)

filter_suffix="OTECD"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_OTECD = df_FILTER_OTECD[ZACHOVANE_SLOUPCE]
df_FILTER_OTECD.to_excel(filepath, index=False)

filter_suffix="SVZOZ"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_SVZOZ = df_FILTER_SVZOZ[ZACHOVANE_SLOUPCE]
df_FILTER_SVZOZ.to_excel(filepath, index=False)

filter_suffix="PLACV"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_PLACV = df_FILTER_PLACV[ZACHOVANE_SLOUPCE]
df_FILTER_PLACV.to_excel(filepath, index=False)

filter_suffix="CERNV"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_CERNV = df_FILTER_CERNV[ZACHOVANE_SLOUPCE]
df_FILTER_CERNV.to_excel(filepath, index=False)

filter_suffix="HOMOF"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_HOMOF = df_FILTER_HOMOF[ZACHOVANE_SLOUPCE]
df_FILTER_HOMOF.to_excel(filepath, index=False)