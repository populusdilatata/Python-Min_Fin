import pandas as pd
from openpyxl import load_workbook

# ========= NASTAVENÍ =========
soubor = "Výplaty/vyplaty_zam2.xlsx"

# název sloupce s osobním číslem
sloupec_id = "oscis"

# název sloupce, pro který chcete počítat změnu
sloupec_hodnoty = "cmzda"
# =============================

xls = pd.ExcelFile(soubor)

data = []

for sheet in xls.sheet_names:

    if sheet == "Rel_zmena":
        continue

    df = pd.read_excel(xls, sheet_name=sheet)

    data.append(df)

df = pd.concat(data, ignore_index=True)

# převod období typu 05.26 -> datum pro správné řazení
print(df["obdobi"].unique())
print(df.head())

df["Obdobi_sort"] = pd.to_datetime(
    df["obdobi"],
    format = "%m.%y"
)


df = df.sort_values(
    [sloupec_id, "Obdobi_sort"]
)

# relativní změna vůči předchozímu období

# přepsání zobrazovaného formátu

df["obdobi"] = (df["Obdobi_sort"]
                .dt.strftime("%m.%Y"))

df = df.sort_values([sloupec_id, "Obdobi_sort"])

df = df.drop(columns=["Obdobi_sort"])
df["Rel_zmena"] = (
    df.groupby(sloupec_id)[sloupec_hodnoty]
      .pct_change()
)

df["Rel_zmena_%"] = (
    df["Rel_zmena"] * 100
).round(2)

test = df[df["oscis"] == 10347]

print(test[["oscis", "obdobi", "cmzda"]])


# zápis do existujícího souboru
with pd.ExcelWriter(
    soubor,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Rel_zmena",
        index=False
    )

print("Hotovo.")