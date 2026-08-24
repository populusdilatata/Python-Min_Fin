import pandas as pd

# ================================
#  NASTAVENÍ (SEM DOSADÍŠ NÁZVY)
# ================================

soubor1 = "pdnyv_meziclanek_dovol_17_05.xlsx"
soubor2 = "vystupf60_dovol.xlsx"

sloupec1_tab1 = "oscis"
sloupec2_tab1 = "den"

sloupec1_tab2 = "Osobní číslo"
sloupec2_tab2 = "aktual_den"

# ================================
# NAČTENÍ DAT
# ================================

tab1 = pd.read_excel(soubor1)
tab2 = pd.read_excel(soubor2)

# ================================
# POROVNÁNÍ
# ================================

# 1) shodné
shodne = pd.merge(
    tab1,
    tab2,
    left_on=[sloupec1_tab1, sloupec2_tab1],
    right_on=[sloupec1_tab2, sloupec2_tab2],
    how="inner"
)

# 2) pouze v TAB1
pouze_tab1 = pd.merge(
    tab1,
    tab2,
    left_on=[sloupec1_tab1, sloupec2_tab1],
    right_on=[sloupec1_tab2, sloupec2_tab2],
    how="left",
    indicator=True
)

pouze_tab1 = pouze_tab1[
    pouze_tab1["_merge"] == "left_only"
][[sloupec1_tab1, sloupec2_tab1]]

# 3) pouze v TAB2
pouze_tab2 = pd.merge(
    tab1,
    tab2,
    left_on=[sloupec1_tab1, sloupec2_tab1],
    right_on=[sloupec1_tab2, sloupec2_tab2],
    how="right",
    indicator=True
)

pouze_tab2 = pouze_tab2[
    pouze_tab2["_merge"] == "right_only"
][[sloupec1_tab2, sloupec2_tab2]]

# ================================
# ULOŽENÍ
# ================================

with pd.ExcelWriter("vysledek_porovnani.xlsx", engine="openpyxl") as writer:
    shodne.to_excel(writer, sheet_name="SHODNE", index=False)
    pouze_tab1.to_excel(writer, sheet_name="POUZE_TAB1", index=False)
    pouze_tab2.to_excel(writer, sheet_name="POUZE_TAB2", index=False)

print("Hotovo: vysledek_porovnani.xlsx")