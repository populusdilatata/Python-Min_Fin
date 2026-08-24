import pandas as pd 
import os
# Architektura               5,0/10 
# Čitelnost                  8,0/10 |
# Dokumentace                5,0/10 |
# Logging                    1/10 |
# Python styl                6,0/10 |
# Robustnost                 4,0/10 |
# Udržovatelnost             4,0/10 |


ROZDELOVANY_SOUBOR = "vstup_f64/07_DATA/f64A.xlsx"
output_dir = "vstup_f64"
base_filename="f64"
SLOUPEC = "prndr"
ZADANY_SLOUPEC_1="prnza"
ZADANY_SLOUPEC_2="prnko"

# --- načtení ---
df_soubor = pd.read_excel(ROZDELOVANY_SOUBOR, engine="openpyxl")
#print(df_soubor)
# --- rozdělení dle prndr ---
# --- prndr = 100, 101 : NEMOC
# --- prndr = 16 : OTECD
# --- prndr = 108 : OCR

df_FILTER_NEMOC = df_soubor[df_soubor[SLOUPEC].isin([100, 101])]
df_FILTER_OTECD = df_soubor[df_soubor[SLOUPEC] == 16]
df_FILTER_OCR = df_soubor[df_soubor[SLOUPEC] == 108]

# --- uložení do jednotlivých souborů ---
filter_suffix="NEMOC"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_NEMOC[ZADANY_SLOUPEC_1] = df_FILTER_NEMOC[ZADANY_SLOUPEC_1].dt.strftime("%d.%m.%Y")
df_FILTER_NEMOC[ZADANY_SLOUPEC_2] = df_FILTER_NEMOC[ZADANY_SLOUPEC_2].dt.strftime("%d.%m.%Y")
df_FILTER_NEMOC.to_excel(filepath, index=False)

filter_suffix="OTECD"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_OTECD[ZADANY_SLOUPEC_1] = df_FILTER_OTECD[ZADANY_SLOUPEC_1].dt.strftime("%d.%m.%Y")
df_FILTER_OTECD[ZADANY_SLOUPEC_2] = df_FILTER_OTECD[ZADANY_SLOUPEC_2].dt.strftime("%d.%m.%Y")
df_FILTER_OTECD.to_excel(filepath, index=False)

filter_suffix="OCR"
filename = f"{base_filename}_FILTER_{filter_suffix}.xlsx"
filepath = os.path.join(output_dir, filename)
df_FILTER_OCR[ZADANY_SLOUPEC_1] = df_FILTER_OCR[ZADANY_SLOUPEC_1].dt.strftime("%d.%m.%Y")
df_FILTER_OCR[ZADANY_SLOUPEC_2] = df_FILTER_OCR[ZADANY_SLOUPEC_2].dt.strftime("%d.%m.%Y")
df_FILTER_OCR.to_excel(filepath, index=False)