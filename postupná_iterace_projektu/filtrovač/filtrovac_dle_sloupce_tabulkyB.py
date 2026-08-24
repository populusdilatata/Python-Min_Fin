import pandas as pd

FILTROVANY_SOUBOR="dochzarv2.xlsx"
FILTRAT="f60_FILTER_LEKAR20.xlsx"
SLOUPEC="oscis"

# --- načtení ---
hlavni = pd.read_excel(FILTROVANY_SOUBOR)
ids = pd.read_excel(FILTRAT)

# --- sjednocení typů ---
hlavni[SLOUPEC] = hlavni[SLOUPEC].astype(str)
ids[SLOUPEC] = ids[SLOUPEC].astype(str)

# --- 1) odstranění duplicit v HLAVNÍ tabulce ---
# zachová se první výskyt každého ID
hlavni_unique = hlavni.drop_duplicates(subset=SLOUPEC, keep="first")

# --- 2) filtrování podle seznamu ID ---
vysledek = hlavni_unique[hlavni_unique[SLOUPEC].isin(ids[SLOUPEC])]

# --- výstup ---
print(vysledek)
vysledek.to_excel("dochzarv2_FILTER_LEKARB.xlsx", index=False)