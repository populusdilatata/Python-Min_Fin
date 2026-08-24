import pandas as pd

FILTROVANY_SOUBOR = "vstup_dochzarv2/07_DATA/dochzarv_07A.xlsx"
FILTRAT = "vstup_f60/07A_DATA/f60_FILTER_PLACV.xlsx"
SLOUPEC = "oscis"
VYBRANE_SLOUPCE = ["oscis", "prijm", "neo2dr","appvdr","ppvdr"]

# --- načtení ---
hlavni = pd.read_excel(FILTROVANY_SOUBOR, engine="openpyxl")
ids = pd.read_excel(FILTRAT, engine="openpyxl")

# --- sjednocení typů ---
hlavni[SLOUPEC] = hlavni[SLOUPEC].astype(str)
ids[SLOUPEC] = ids[SLOUPEC].astype(str)

# --- 1) odstranění duplicit v HLAVNÍ tabulce ---
hlavni_unique = hlavni.drop_duplicates(subset=SLOUPEC, keep="first")

# --- 2) napojení (zachová duplicity z filtrátu) ---
vysledek = ids.merge(hlavni_unique, on=SLOUPEC, how="left")
#print(vysledek.head)
vysledek = vysledek[VYBRANE_SLOUPCE]
# --- 3) vytvoření rozdilového sloupce ---
vysledek["rozdil"]=vysledek["appvdr"]-vysledek["ppvdr"]
# --- 4) pouze nenulové řadky ---
nenulove = vysledek[vysledek["rozdil"] != 0]
print(nenulove)

# --- 5) vytvoření podmínky pro logický test ---

def kontrola(radek):
    if radek["appvdr"] in [210, 200] and radek["neo2dr"] in [143, 147, 150, 160]:
        return "splněno"
    elif radek["appvdr"] == 101 and radek["neo2dr"] in [78, 137, 150, 160]:
        return "splněno"
    else:
        return "nesplněno"

vysledek["stav"] = vysledek.apply(kontrola, axis=1)

# --- 6) pouze řadky s nesplněno ---
print("Výpis výsledků logického testu")
nenulove2 = vysledek[vysledek["stav"] == "nesplněno"]
print(nenulove2)



# --- výstup ---
#print(vysledek)
with pd.ExcelWriter("dochzarv2_06_FILTER_PLACV.xlsx", engine="openpyxl") as writer:
    nenulove.to_excel(writer, sheet_name="rozdilf60(appvdr)_vs_dochzarv(ppvdr)", index=False)
    nenulove2.to_excel(writer, sheet_name="nesplňuje logickou podmínku", index=False)


#vysledek.to_excel("dochzarv2_FILTER_PLACV.xlsx", index=False)