import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# ==================================================
# NASTAVENÍ
# ==================================================

soubor = "Výplaty/vyplaty_zam4.xlsx"

sloupec_id = "oscis"
sloupec_hodnoty = "hmzda_odmen"
sloupec_vedouci = "appvdr"

obdobi_boxplot = [
    "05.2026",
    "06.2026",
    "07.2026"
]

# ==================================================
# NAČTENÍ DAT
# ==================================================

xls = pd.ExcelFile(soubor)

data = []

for sheet in xls.sheet_names:

    if sheet.startswith("INC_"):
        continue

    if sheet.startswith("DEC_"):
        continue

    if sheet.startswith("Obd_"):
        continue

    if sheet in ["Rel_zmena", "Top5_pct"]:
        continue

    print(f"Načítám list: {sheet}")

    df_sheet = pd.read_excel(
        xls,
        sheet_name=sheet
    )

    data.append(df_sheet)

if len(data) == 0:
    raise ValueError(
        "Nebyla nalezena žádná vstupní data."
    )

df = pd.concat(
    data,
    ignore_index=True
)

print("\nPočet načtených řádků:", len(df))

# ==================================================
# KONTROLA SLOUPCŮ
# ==================================================

povinne_sloupce = [
    sloupec_id,
    sloupec_hodnoty,
    sloupec_vedouci,
    "obdobi"
]

for sloupec in povinne_sloupce:

    if sloupec not in df.columns:

        raise ValueError(
            f"Ve vstupních datech chybí sloupec: "
            f"{sloupec}"
        )

# ==================================================
# VÝPOČET RELATIVNÍ ZMĚNY
# ==================================================

df["Obdobi_sort"] = pd.to_datetime(
    df["obdobi"],
    format="%m.%y",
    errors="coerce"
)

df = df.dropna(
    subset=["Obdobi_sort"]
)

df = df.sort_values(
    [sloupec_id, "Obdobi_sort"]
)

df["Rel_zmena"] = (
    df.groupby(sloupec_id)[sloupec_hodnoty]
      .pct_change()
)

# odstranění inf hodnot
df["Rel_zmena"] = (
    df["Rel_zmena"]
    .replace(
        [np.inf, -np.inf],
        np.nan
    )
)

df["Rel_zmena_%"] = (
    df["Rel_zmena"] * 100
).round(2)

df["obdobi"] = (
    df["Obdobi_sort"]
      .dt.strftime("%m.%Y")
)

df = df.drop(
    columns=["Obdobi_sort"]
)

print("\nPopis Rel_zmena_%")
print(
    df["Rel_zmena_%"]
    .describe()
)

# ==================================================
# TOP 5 %
# ==================================================

vysledky_top = []

print("\n")
print("=" * 70)
print("TOP 5 % NÁRŮSTŮ A POKLESŮ")
print("=" * 70)

for obdobi, skupina in df.groupby("obdobi"):

    data_mesic = skupina.copy()

    data_mesic = data_mesic.dropna(
        subset=["Rel_zmena_%"]
    )

    if len(data_mesic) < 5:

        print(
            f"{obdobi}: přeskočeno "
            f"(málo záznamů)"
        )

        continue

    hranice_95 = (
        data_mesic["Rel_zmena_%"]
        .quantile(0.95)
    )

    hranice_05 = (
        data_mesic["Rel_zmena_%"]
        .quantile(0.05)
    )

    top_increase = data_mesic[
        data_mesic["Rel_zmena_%"]
        >= hranice_95
    ].copy()

    top_decrease = data_mesic[
        data_mesic["Rel_zmena_%"]
        <= hranice_05
    ].copy()

    print("\n" + "=" * 50)
    print(f"Období: {obdobi}")

    print(
        f"95. percentil: "
        f"{hranice_95:.2f}"
    )

    print(
        f"5. percentil: "
        f"{hranice_05:.2f}"
    )

    print(
        f"TOP růsty: "
        f"{len(top_increase)}"
    )

    print(
        f"TOP poklesy: "
        f"{len(top_decrease)}"
    )

    for _, row in top_increase.iterrows():

        vysledky_top.append({
            "obdobi": obdobi,
            "typ": "top5_pct_INCREASE",
            sloupec_id: row[sloupec_id],
            sloupec_vedouci: row[sloupec_vedouci],
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
            sloupec_vedouci: row[sloupec_vedouci],
            sloupec_hodnoty: row[sloupec_hodnoty],
            "Rel_zmena_%": row["Rel_zmena_%"],
            "hranice_percentilu": round(
                hranice_05,
                2
            )
        })

# ==================================================
# DATAFRAME TOP
# ==================================================

df_top = pd.DataFrame(
    vysledky_top,
    columns=[
        "obdobi",
        "typ",
        sloupec_id,
        sloupec_vedouci,
        sloupec_hodnoty,
        "Rel_zmena_%",
        "hranice_percentilu"
    ]
)

print("\nPočet TOP záznamů:", len(df_top))

# ==================================================
# ULOŽENÍ DO EXCELU
# ==================================================

with pd.ExcelWriter(
    soubor,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:

    # hlavní list
    df.to_excel(
        writer,
        sheet_name="Rel_zmena",
        index=False
    )

    # souhrnný TOP list
    df_top.to_excel(
        writer,
        sheet_name="Top5_pct",
        index=False
    )

    # ------------------------------------------
    # LISTY OBDOBÍ
    # ------------------------------------------

    for obdobi in sorted(
        df["obdobi"].dropna().unique()
    ):

        nazev = (
            "Obd_" +
            obdobi.replace(".", "_")
        )[:31]

        df_obd = df[
            df["obdobi"] == obdobi
        ]

        df_obd.to_excel(
            writer,
            sheet_name=nazev,
            index=False
        )

    # ------------------------------------------
    # LISTY INC / DEC
    # ------------------------------------------

    if not df_top.empty:

        for obdobi in sorted(
            df_top["obdobi"].dropna().unique()
        ):

            inc = df_top[
                (df_top["obdobi"] == obdobi)
                &
                (
                    df_top["typ"]
                    == "top5_pct_INCREASE"
                )
            ]

            dec = df_top[
                (df_top["obdobi"] == obdobi)
                &
                (
                    df_top["typ"]
                    == "top5_pct_DECREASE"
                )
            ]

            inc_sheet = (
                "INC_" +
                obdobi.replace(".", "_")
            )[:31]

            dec_sheet = (
                "DEC_" +
                obdobi.replace(".", "_")
            )[:31]

            inc.to_excel(
                writer,
                sheet_name=inc_sheet,
                index=False
            )

            dec.to_excel(
                writer,
                sheet_name=dec_sheet,
                index=False
            )

print("\nExcel byl úspěšně aktualizován.")

# ==================================================
# BOXPLOT
# ==================================================

boxplot_data = []
platna_obdobi = []

for obd in obdobi_boxplot:

    hodnoty = (
        df.loc[
            df["obdobi"] == obd,
            "Rel_zmena_%"
        ]
        .dropna()
    )

    if len(hodnoty) > 0:

        boxplot_data.append(
            hodnoty
        )

        platna_obdobi.append(
            obd
        )

if len(boxplot_data) > 0:

    plt.figure(
        figsize=(10, 6)
    )

    plt.boxplot(
        boxplot_data,
        tick_labels=platna_obdobi
    )

    plt.title(
        "Relativní změna hmzda_odmen (%)"
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
        "Výplaty/"
        "boxplot_rel_zmena.png"
    )

    plt.savefig(
        soubor_graf,
        dpi=300,
        bbox_inches="tight"
    )

    plt.show()

    print(
        f"\nGraf uložen: {soubor_graf}"
    )

else:

    print(
        "\nPro boxplot nejsou "
        "k dispozici žádná data."
    )