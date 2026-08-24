import pandas as pd

FILTROVANY_SOUBOR="dochzarv2.xlsx"
FILTRAT="f60_FILTER_LEKAR20.xlsx"
SLOUPEC="oscis"


hlavni = pd.read_excel(FILTROVANY_SOUBOR, engine="openpyxl")
ids = pd.read_excel(FILTRAT, engine="openpyxl")

seznam_id = ids[SLOUPEC]

vysledek = hlavni[hlavni[SLOUPEC].isin(seznam_id)]

print(vysledek)
vysledek.to_excel("dochzarv2_FILTER_LEKAR.xlsx", index=False)