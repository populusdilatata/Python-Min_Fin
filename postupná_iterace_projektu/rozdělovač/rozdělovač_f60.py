import pandas as pd 
import os

ROZDELOVANY_SOUBOR = "vstup_f60/07_DATA/f60.xlsx"
output_dir = "vstup_f60"
base_filename="f60"
SLOUPEC = "neo2dr"
ZADANY_SLOUPEC_1="neo2za"
ZADANY_SLOUPEC_2="neo2ko"

# --- načtení ---
df_soubor = pd.read_excel(ROZDELOVANY_SOUBOR, engine="openpyxl")
#print(df_soubor)
# --- rozdělení dle neo2dr ---
# --- neo2dr = 140, 77, 139, 82, 160 : LEKAR
# --- neo2dr = 143,78, 147, 137, 163, 150: PLACV
# --- neo2dr = 158, 52 : INDIV
# --- neo2dr = 75, 85 : ABSEN
# --- neo2dr = 95 : DOVOL
# --- neo2dr = 151 : SVZOZ
# --- neo2dr = 181 : STUDV
# --- neo2dr = VSE OSTATNI DROBY


map_kategorie = {
    "LEKAR": [140, 77, 139, 82, 160],
    "PLACV": [143, 78, 147, 137, 163, 150],  
    "INDIV": [158, 52],
    "ABSEN": [75, 85],
    "DOVOL": [95],
    "SVZOZ": [151],  
    "STUDV": [181],
}

def prirad_kategorii(hodnota):
    for kat, hodnoty in map_kategorie.items():
        if hodnota in hodnoty:
            return kat
    return "DROBY"

df_soubor["kategorie"] = df_soubor["neo2dr"].apply(prirad_kategorii)


os.makedirs(output_dir, exist_ok=True)

for kat, data in df_soubor.groupby("kategorie"):
 
    df_export = data.copy()    
    df_export[ZADANY_SLOUPEC_1] = df_export[ZADANY_SLOUPEC_1].dt.strftime("%d.%m.%Y")
    df_export[ZADANY_SLOUPEC_2] = df_export[ZADANY_SLOUPEC_2].dt.strftime("%d.%m.%Y")

    
    filename = f"f60_FILTER_{kat}.xlsx"
    path = os.path.join(output_dir, filename)
    df_export.to_excel(path, index=False)


    
    

