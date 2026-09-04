import pandas as pd
import matplotlib.pyplot as plt

# ========= NASTAVENÍ =========

soubor = "Výplaty/vyplaty_zam4.xlsx"

# název sloupce s osobním číslem
sloupec_id = "oscis"

# název sloupce, pro který chcete počítat změnu
sloupec_hodnoty = "hmzda_odmen"

# období pro boxploty
obdobi_boxplot = ["05.2026", "06.2026", "07.2026"]

# =============================


# ==================================================
# NAČTENÍ DAT
# ==================================================

xls = pd.ExcelFile(soubor)

data = []

for sheet in xls.sheet_names:

    if sheet in ["Rel_zmena", "Top5_pct"]:
        continue

    df_sheet = pd.read_excel(
        xls,
        sheet_name=sheet
    )

    data.append(df_sheet)

# spojení všech listů
df = pd.concat(data, ignore_index=True)

print("Nalezená období:")
print(df["obdobi"].unique())

print("\nUkázka dat:")
print(df.head())


# ==================================================
# VÝPOČET RELATIVNÍ ZMĚNY
# ==================================================

# převod období typu 05.26 -> datum
df["Obdobi_sort"] = pd.to_datetime(
    df["obdobi"],
    format="%m.%y"
)

# seřazení dle zaměstnance a období
df = df.sort_values(
    [sloupec_id, "Obdobi_sort"]
)

# relativní změna vůči předchozímu období
df["Rel_zmena"] = (
    df.groupby(sloupec_id)[sloupec_hodnoty]
      .pct_change()
)

df["Rel_zmena_%"] = (
    df["Rel_zmena"] * 100
).round(2)

# převod zpět na MM.RRRR
df["obdobi"] = (
    df["Obdobi_sort"]
      .dt.strftime("%m.%Y")
)

# odstranění pomocného sloupce
df = df.drop(columns=["Obdobi_sort"])


# ==================================================
# KONTROLA JEDNOHO ZAMĚSTNANCE
# ==================================================

test_id = 10347

test = df[df[sloupec_id] == test_id]

print(f"\nKontrola zaměstnance {test_id}:")
print(
    test[
        [
            sloupec_id,
            "obdobi",
            sloupec_hodnoty,
            "Rel_zmena_%"
        ]
    ]
)


# ==================================================
# TOP 5 % NÁRŮSTŮ / POKLESŮ POMOCÍ PERCENTILŮ
# ==================================================

vysledky_top = []

print("\n")
print("=" * 80)
print("TOP 5 % NÁRŮSTŮ A POKLESŮ")
print("=" * 80)

for obdobi, skupina in df.groupby("obdobi"):

    data_mesic = skupina.dropna(
        subset=["Rel_zmena_%"]
    ).copy()

    if len(data_mesic) == 0:
        continue

    hranice_95 = data_mesic["Rel_zmena_%"].quantile(
        0.95
    )

    hranice_05 = data_mesic["Rel_zmena_%"].quantile(
        0.05
    )

    top_increase = data_mesic[
        data_mesic["Rel_zmena_%"] >= hranice_95
    ].copy()

    top_decrease = data_mesic[
        data_mesic["Rel_zmena_%"] <= hranice_05
    ].copy()

    print("\n" + "=" * 60)
    print(f"OBDOBÍ: {obdobi}")

    print(
        f"95. percentil: {hranice_95:.2f} %"
    )

    print(
        f"5. percentil: {hranice_05:.2f} %"
    )

    print(
        f"\nPočet zaměstnanců v TOP růstu: {len(top_increase)}"
    )

    print(
        f"Počet zaměstnanců v TOP poklesu: {len(top_decrease)}"
    )

    print("\nTOP 5 % NÁRŮSTŮ")

    print(
        top_increase[
            [
                sloupec_id,
                sloupec_hodnoty,
                "Rel_zmena_%"
            ]
        ]
        .sort_values(
            "Rel_zmena_%",
            ascending=False
        )
        .to_string(index=False)
    )

    print("\nOsobní čísla:")

    print(
        top_increase[sloupec_id]
        .astype(int)
        .tolist()
    )

    print("\nTOP 5 % POKLESŮ")

    print(
        top_decrease[
            [
                sloupec_id,
                sloupec_hodnoty,
                "Rel_zmena_%"
            ]
        ]
        .sort_values(
            "Rel_zmena_%"
        )
        .to_string(index=False)
    )

    print("\nOsobní čísla:")

    print(
        top_decrease[sloupec_id]
        .astype(int)
        .tolist()
    )

    for _, row in top_increase.iterrows():

        vysledky_top.append({
            "obdobi": obdobi,
            "typ": "top5_pct_INCREASE",
            sloupec_id: row[sloupec_id],
            sloupec_hodnoty: row[sloupec_hodnoty],
            "Rel_zmena_%": row["Rel_zmena_%"],
            "hranice_percentilu": round(
                hranice_95,
                2
            )
        })

    for _, row in top_decrease.iterrows():

        vysledky_top.append({
            "obdobi": obdobi,
            "typ": "top5_pct_DECREASE",
            sloupec_id: row[sloupec_id],
            sloupec_hodnoty: row[sloupec_hodnoty],
            "Rel_zmena_%": row["Rel_zmena_%"],
            "hranice_percentilu": round(
                hranice_05,
                2
            )
        })

df_top = pd.DataFrame(vysledky_top)


# ==================================================
# ULOŽENÍ DO EXCELU
# ==================================================

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

    df_top.to_excel(
        writer,
        sheet_name="Top5_pct",
        index=False
    )

print("\nList 'Rel_zmena' byl uložen.")
print("List 'Top5_pct' byl uložen.")


# ==================================================
# BOXPLOTY RELATIVNÍCH ZMĚN
# ==================================================

boxplot_data = []

for obd in obdobi_boxplot:

    hodnoty = (
        df.loc[
            df["obdobi"] == obd,
            "Rel_zmena_%"
        ]
        .dropna()
    )

    boxplot_data.append(hodnoty)

    print(
        f"{obd}: {len(hodnoty)} hodnot"
    )

plt.figure(figsize=(10, 6))

plt.boxplot(
    boxplot_data,
    tick_labels=obdobi_boxplot
)

plt.title(
    "Relativní změna mzdy (%) podle období"
)

plt.xlabel("Období")
plt.ylabel("Relativní změna [%]")

plt.grid(
    axis="y",
    linestyle="--",
    alpha=0.5
)

plt.tight_layout()

soubor_graf = (
    "Výplaty/boxplot_rel_zmena_05_06_07_2026.png"
)

plt.savefig(
    soubor_graf,
    dpi=300,
    bbox_inches="tight"
)

plt.show()

print(
    f"\nBoxplot uložen do souboru: {soubor_graf}"
)