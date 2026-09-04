import logging
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ==================================================
# NASTAVENÍ
# ==================================================

soubor = "Výplaty/vyplaty_zam4.xlsx"

sloupec_id = "oscis"
sloupec_hodnoty = "hmzda_odmen"
sloupec_vedouci = "appvdr"

obdobi_boxplot = [
    "05.26",
    "06.26",
    "07.26"
]

# povinná období zaměstnance
kontrolovana_obdobi = [
    "04.26",
    "05.26",
    "06.26"
]

# ==================================================
# VÝSTUPNÍ SOUBOR
# ==================================================

vstup = Path(soubor)

soubor_vystup = (
    vstup.parent /
    f"RESULT_{vstup.name}"
)

# ==================================================
# LOGGER
# ==================================================

logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

logger = logging.getLogger(__name__)

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

    if sheet in [
        "Rel_zmena",
        "Top5_pct",
        "WARNING_MZDY0", 
        "WARNING-kontrola kompletnosti"
    ]:
        continue

    print(f"Načítám list: {sheet}")

    df_sheet = pd.read_excel(
        xls,
        sheet_name=sheet
    )

    data.append(df_sheet)

if not data:
    raise ValueError(
        "Nebyla nalezena žádná vstupní data."
    )

df = pd.concat(
    data,
    ignore_index=True
)

# sjednocení formátu období na MM.YY

if pd.api.types.is_datetime64_any_dtype(df["obdobi"]):

    df["obdobi"] = (
        pd.to_datetime(df["obdobi"])
        .dt.strftime("%m.%y")
    )

else:
    """
    df["obdobi"] = (
        pd.to_datetime(
            df["obdobi"],
            errors="coerce"
        )
        .dt.strftime("%m.%y")
        .fillna(
            df["obdobi"].astype(str)
        )
    )
    """
    df["obdobi"] = (
            df["obdobi"]
            .astype(str)
            .str.strip()
            )

    df["Obdobi_sort"] = pd.to_datetime(
                        df["obdobi"],
                        format="%m.%y",
                        errors="coerce")


df["obdobi"] = (
    df["obdobi"]
    .astype(str)
    .str.strip()
)

print(
    f"\nPočet načtených řádků: {len(df)}"
)

# ==================================================
# KONTROLA POVINNÝCH SLOUPCŮ
# ==================================================

povinne_sloupce = [
    "oscis",
    "obdobi",
    "pracvz",
    "prijm",
    "hmzda",
    "hmzda_odmen",
    "appvdr"
]

for sloupec in povinne_sloupce:

    if sloupec not in df.columns:

        raise ValueError(
            f"Chybí povinný sloupec: {sloupec}"
        )

# ==================================================
# NULOVÉ MZDY - VYŘAZENÍ ZE ZPRACOVÁNÍ
# ==================================================

warning_mzdy = df[
    (df["hmzda"].isna())
    |
    (df["hmzda"] == 0)
    |
    (df["hmzda_odmen"].isna())
    |
    (df["hmzda_odmen"] == 0)
].copy()

if not warning_mzdy.empty:

    logger.warning(
        "Nalezeno %s záznamů s hmzda=0/NULL "
        "nebo hmzda_odmen=0/NULL. "
        "Budou vyřazeny ze zpracování.",
        len(warning_mzdy)
    )

    logger.warning(
        "Dotčení zaměstnanci: %s",
        warning_mzdy["oscis"]
        .astype(str)
        .tolist()
    )

# vyřazení ze zpracování

df = df[
    ~(df["hmzda"].isna())
    &
    (df["hmzda"] != 0)
    &
    ~(df["hmzda_odmen"].isna())
    &
    (df["hmzda_odmen"] != 0)
].copy()

print(
    f"Počet řádků po odstranění nulových mezd: {len(df)}"
)

# ==================================================
# PŘÍPRAVA OBDOBÍ
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

# ==================================================
# KONTROLA KOMPLETNOSTI OBDOBÍ
# ==================================================

warning_kompletnost = []

for oscis, skupina in df.groupby("oscis"):
    existujici_obdobi = set(
        skupina["Obdobi_sort"]
            .dt.strftime("%m.%y")
                )

    """
    existujici_obdobi = set(
        skupina["obdobi"]
        .astype(str)
        .str.strip()
    )
    """

    chybejici = [
        obd
        for obd in kontrolovana_obdobi
        if obd not in existujici_obdobi
    ]

    if len(chybejici) > 0:

        temp = skupina.copy()

        temp["CHYBEJICI_OBDOBI"] = (
            ", ".join(chybejici)
        )

        temp["obdobi"] = (
            pd.to_datetime(temp["Obdobi_sort"])
            .dt.strftime("%m.%y")
            )

        warning_kompletnost.append(temp)

if warning_kompletnost:

    warning_kompletnost = pd.concat(
        warning_kompletnost,
        ignore_index=True
    )

    nekompletni_oscis = (
        warning_kompletnost["oscis"]
        .unique()
    )

    logger.warning(
        "Nalezeno %s zaměstnanců s nekompletní historií.",
        len(nekompletni_oscis)
    )

    logger.warning(
        "Osobní čísla: %s",
        list(nekompletni_oscis)
    )

    print(
        f"\nNekompletních zaměstnanců: "
        f"{len(nekompletni_oscis)}"
    )

    print(
        f"Vyřazeno řádků: "
        f"{len(warning_kompletnost)}"
    )

    # ----------------------------------------------
    # VYŘAZENÍ ZE ZPRACOVÁNÍ
    # ----------------------------------------------

    df = df[
        ~df["oscis"].isin(
            nekompletni_oscis
        )
    ].copy()

    print(
        f"Po kontrole kompletnosti "
        f"zůstalo {len(df)} řádků."
    )

else:

    warning_kompletnost = pd.DataFrame()

    print(
        "\nVšichni zaměstnanci mají kompletní historii."
    )
# ==================================================
# RELATIVNÍ ZMĚNA
# ==================================================

df["Rel_zmena"] = (
    df.groupby(sloupec_id)[sloupec_hodnoty]
      .pct_change()
)

df["Rel_zmena"] = (
    df["Rel_zmena"]
      .replace([np.inf, -np.inf], np.nan)
)

df["Rel_zmena_%"] = (
    df["Rel_zmena"] * 100
).round(2)

"""
df["obdobi"] = (
    df["Obdobi_sort"]
      .dt.strftime("%m.%y")
)
"""

df = df.drop(
    columns=["Obdobi_sort"]
)

print("\nStatistika Rel_zmena_%")

print(
    df["Rel_zmena_%"]
    .describe()
)

# ==================================================
# VÝSTUPNÍ SLOUPCE
# ==================================================

vystupni_sloupce = [
    "oscis",
    "obdobi",
    "pracvz",
    "prijm",
    "hmzda",
    "hmzda_odmen",
    "appvdr",
    "Rel_zmena",
    "Rel_zmena_%"
]

df_export = df[
    vystupni_sloupce
].copy()

# ==================================================
# KONTROLA ZAMĚSTNANCE
# ==================================================

test_id = 10347

test = df[
    df["oscis"] == test_id
]

if not test.empty:

    print(
        f"\nKontrola zaměstnance {test_id}:"
    )

    print(
        test[
            [
                "oscis",
                "obdobi",
                "hmzda_odmen",
                "Rel_zmena",
                "Rel_zmena_%"
            ]
        ]
    )

# ==================================================
# TOP 5 % NÁRŮSTŮ A POKLESŮ
# ==================================================

vysledky_top = []

print("\n" + "=" * 80)
print("TOP 5 % NÁRŮSTŮ A POKLESŮ")
print("=" * 80)

for obdobi, skupina in df.groupby("obdobi"):

    data_mesic = skupina.dropna(
        subset=["Rel_zmena_%"]
    ).copy()

    if len(data_mesic) < 5:

        print(
            f"{obdobi}: přeskočeno (málo dat)"
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
        data_mesic["Rel_zmena_%"] >= hranice_95
    ].copy()

    top_decrease = data_mesic[
        data_mesic["Rel_zmena_%"] <= hranice_05
    ].copy()

    print("\n" + "=" * 60)
    print(f"OBDOBÍ: {obdobi}")

    print(
        f"95. percentil: {hranice_95:.2f}%"
    )

    print(
        f"5. percentil: {hranice_05:.2f}%"
    )

    print(
        f"TOP růsty: {len(top_increase)}"
    )

    print(
        f"TOP poklesy: {len(top_decrease)}"
    )

    print(
        "Osobní čísla TOP růstů:",
        top_increase["oscis"]
        .astype(str)
        .tolist()
    )

    print(
        "Osobní čísla TOP poklesů:",
        top_decrease["oscis"]
        .astype(str)
        .tolist()
    )

    for _, row in top_increase.iterrows():

        vysledky_top.append({
            "oscis": row["oscis"],
            "obdobi": row["obdobi"],
            "pracvz": row["pracvz"],
            "prijm": row["prijm"],
            "hmzda": row["hmzda"],
            "hmzda_odmen": row["hmzda_odmen"],
            "appvdr": row["appvdr"],
            "Rel_zmena": row["Rel_zmena"],
            "Rel_zmena_%": row["Rel_zmena_%"],
            "typ": "top5_pct_INCREASE"
        })

    for _, row in top_decrease.iterrows():

        vysledky_top.append({
            "oscis": row["oscis"],
            "obdobi": row["obdobi"],
            "pracvz": row["pracvz"],
            "prijm": row["prijm"],
            "hmzda": row["hmzda"],
            "hmzda_odmen": row["hmzda_odmen"],
            "appvdr": row["appvdr"],
            "Rel_zmena": row["Rel_zmena"],
            "Rel_zmena_%": row["Rel_zmena_%"],
            "typ": "top5_pct_DECREASE"
        })

df_top = pd.DataFrame(vysledky_top)

print(
    f"\nPočet TOP záznamů: {len(df_top)}"
)

# ==================================================
# EXPORT DO EXCELU
# ==================================================

with pd.ExcelWriter(
    soubor_vystup,
    engine="openpyxl"
) as writer:

    # ----------------------------------------------
    # RELATIVNÍ ZMĚNY
    # ----------------------------------------------

    df_export.to_excel(
        writer,
        sheet_name="Rel_zmena",
        index=False
    )

    # ----------------------------------------------
    # NULOVÉ MZDY
    # ----------------------------------------------

    if not warning_mzdy.empty:

        warning_sloupce = [
            "oscis",
            "obdobi",
            "pracvz",
            "prijm",
            "hmzda",
            "hmzda_odmen",
            "appvdr"
        ]

        warning_mzdy[
            warning_sloupce
        ].to_excel(
            writer,
            sheet_name="WARNING_MZDY0",
            index=False
        )

        # ----------------------------------------------
    # KONTROLA KOMPLETNOSTI
    # ----------------------------------------------

    if not warning_kompletnost.empty:

        export_kompletnost = (
            warning_kompletnost
            .drop(columns=["Obdobi_sort"], errors="ignore")
            .sort_values(["oscis", "obdobi"])
            )

        export_kompletnost.to_excel(
            writer,
            sheet_name="WARNING-kontrola kompletnosti",
            index=False
        )

    # ----------------------------------------------
    # SOUHRNNÝ TOP LIST
    # ----------------------------------------------

    if not df_top.empty:

        (
            df_top
            .drop(columns=["typ"])
            .sort_values(
                "Rel_zmena_%",
                ascending=False
            )
            .to_excel(
                writer,
                sheet_name="Top5_pct",
                index=False
            )
        )

    else:

        pd.DataFrame(
            columns=vystupni_sloupce
        ).to_excel(
            writer,
            sheet_name="Top5_pct",
            index=False
        )

    # ----------------------------------------------
    # LISTY OBDOBÍ
    # ----------------------------------------------

    for obdobi in sorted(
        df_export["obdobi"].unique()
    ):

        nazev = (
            "Obd_" +
            obdobi.replace(".", "_")
        )[:31]

        (
            df_export[
                df_export["obdobi"] == obdobi
            ]
            .sort_values(
                ["prijm", "oscis"]
            )
            .to_excel(
                writer,
                sheet_name=nazev,
                index=False
            )
        )

    # ----------------------------------------------
    # INC / DEC LISTY
    # ----------------------------------------------

    if not df_top.empty:

        for obdobi in sorted(
            df_top["obdobi"].unique()
        ):

            inc = (
                df_top[
                    (df_top["obdobi"] == obdobi)
                    &
                    (
                        df_top["typ"]
                        == "top5_pct_INCREASE"
                    )
                ]
                .drop(columns=["typ"])
                .sort_values(
                    "Rel_zmena_%",
                    ascending=False
                )
            )

            dec = (
                df_top[
                    (df_top["obdobi"] == obdobi)
                    &
                    (
                        df_top["typ"]
                        == "top5_pct_DECREASE"
                    )
                ]
                .drop(columns=["typ"])
                .sort_values(
                    "Rel_zmena_%",
                    ascending=True
                )
            )

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

print(
    f"\nVýstupní soubor vytvořen:"
)
print(
    soubor_vystup
)

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
        vstup.parent /
        f"boxplot_{vstup.stem}.png"
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
        "\nPro boxplot nebyla nalezena žádná data."
    )