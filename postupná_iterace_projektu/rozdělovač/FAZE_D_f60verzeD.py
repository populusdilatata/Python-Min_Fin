"""
Skript pro rozdělení souboru F60 podle hodnot ve sloupci neo2dr.

Účel:
    Rozdělit vstupní soubor F60 do samostatných XLSX souborů
    podle definovaného mapování hodnot neo2dr.

Vstup:
    vstup_f60/07A_DATA/f60.xlsx

Povinné sloupce:
    neo2dr
    neo2za
    neo2ko
    neo2pl

Výstup:
    vstup_f60/f60_FILTER_<KATEGORIE>.xlsx

Neznámé hodnoty:
    Hodnoty neo2dr, které nejsou uvedeny v MAP_KATEGORIE,
    budou zařazeny do kategorie DROBY.

Autor: Majda Tomáš
Verze: 3.0
"""

from dataclasses import dataclass
from pathlib import Path
import logging

import pandas as pd

# ================================

# Architektura               9,5/10
# Čitelnost                  9,5/10
# Dokumentace                10/10
# Logging                    9,5/10
# Python styl                9,5/10
# Robustnost                 9,6/10
# Udržovatelnost             9,6/10

# ============================================================================
# KONSTANTY
# ============================================================================

SLOUPEC_VYSTUPNI_KATEGORIE = "kategorie"
KATEGORIE_DROBY = "DROBY"

# ============================================================================
# MAPOVÁNÍ KATEGORIÍ
# ============================================================================
# Hodnoty neo2dr jsou mapovány na výstupní kategorie.
# Každá hodnota smí být použita pouze jednou.
# Hodnoty neuvedené v mapování jsou zařazeny do DROBY.

MAP_KATEGORIE: dict[str, list[int]] = {
    "ABSEN": [75, 85],
    "DOVOL": [95],
    "INDIV": [158, 52],
    "LEKAR": [140, 77, 139, 82, 160],
    "PLACV": [143, 78, 147, 137, 163, 150],
    "STUDV": [181],
    "SVZOZ": [151],
}

# ============================================================================
# KONFIGURACE
# ============================================================================

@dataclass(frozen=True)
class Config:
    """
    Konfigurace zpracování F60.

    Attributes
    ----------
    vstupni_soubor
        Cesta ke vstupnímu XLSX souboru.

    output_dir
        Adresář pro export výsledků.

    sloupec_kategorie
        Sloupec obsahující kódy neo2dr.

    sloupec_datum_od
        Sloupec začátku období.

    sloupec_datum_do
        Sloupec konce období.
    """
    vstupni_soubor: Path
    output_dir: Path
    sloupec_kategorie: str
    sloupec_datum_od: str
    sloupec_datum_do: str
    sloupec_rozdeleni_dovolena: str


# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(
            "f60.log",
            encoding="utf-8",
        ),
        logging.StreamHandler(),
    ],
)

LOGGER = logging.getLogger(__name__)

# ============================================================================
# FUNKCE
# ============================================================================


def vytvor_lookup_kategorii() -> dict[int, str]:
    """
    Vytvoří mapu:
        hodnota neo2dr -> kategorie.

    Současně ověří duplicity.

    Raises
    ------
    ValueError
        Pokud je některá hodnota namapována
        do více kategorií.
    """

    lookup: dict[int, str] = {}

    for kategorie, hodnoty in MAP_KATEGORIE.items():

        for hodnota in hodnoty:

            if hodnota in lookup:
                raise ValueError(
                    f"Hodnota {hodnota} je namapována "
                    f"vícekrát ({lookup[hodnota]}, "
                    f"{kategorie})."
                )

            lookup[hodnota] = kategorie

    return lookup

def validuj_a_normalizuj_kategorii(
    df: pd.DataFrame,
    sloupec: str,
    strict: bool = False,
) -> None:
    """
    Ověří a sjednotí hodnoty ve sloupci neo2dr.

    Příklady převodu:
        75      -> 75
        "75"    -> 75
        "75.0"  -> 75
        " 75 "  -> 75

    Neplatné hodnoty:
        "ABC"
        "X75"
        "***"

    budou převedeny na <NA>.

    Parameters
    ----------
    df
        Zdrojový DataFrame.
    sloupec
        Název sloupce s hodnotami neo2dr.
    strict
        Pokud True, nalezení neplatných hodnot
        vyvolá výjimku.

    Raises
    ------
    ValueError
        Pokud sloupec neexistuje nebo jsou nalezeny
        neplatné hodnoty v režimu strict.
    """

    if sloupec not in df.columns:
        raise ValueError(
            f"Sloupec '{sloupec}' neexistuje."
        )

    puvodni_hodnoty = df[sloupec].copy()

    hodnoty = pd.to_numeric(
        puvodni_hodnoty.astype(str).str.strip(),
        errors="coerce",
    )

    neplatne_maska = (
        puvodni_hodnoty.notna()
        & hodnoty.isna()
    )

    neplatnych = int(neplatne_maska.sum())

    if neplatnych:
        ukazka = (
            puvodni_hodnoty[neplatne_maska]
            .astype(str)
            .unique()[:10]
        )

        LOGGER.warning(
            "Sloupec %s obsahuje %s neplatných hodnot.",
            sloupec,
            neplatnych,
        )

        LOGGER.warning(
            "Ukázka neplatných hodnot: %s",
            ", ".join(ukazka),
        )

        if strict:
            raise ValueError(
                f"Ve sloupci '{sloupec}' byly nalezeny "
                f"neplatné hodnoty."
            )

    df[sloupec] = hodnoty.astype("Int64")

    LOGGER.info(
        "Sloupec %s byl normalizován na typ Int64.",
        sloupec,
    )


def validuj_sloupce(
    df: pd.DataFrame,
    config: Config,
) -> None:
    """
    Ověří existenci povinných sloupců.

    Raises
    ------
    ValueError
        Pokud některý z povinných sloupců chybí.
    """

    povinne_sloupce = {
        config.sloupec_kategorie,
        config.sloupec_datum_od,
        config.sloupec_datum_do,
        config.sloupec_rozdeleni_dovolena,
    }

    chybejici = povinne_sloupce - set(df.columns)

    if chybejici:
        raise ValueError(
            f"Chybí povinné sloupce: "
            f"{', '.join(sorted(chybejici))}"
        )


def validuj_droby(
    df: pd.DataFrame,
    config: Config,
) -> None:
    """
    Zaloguje neznámé hodnoty.
    """

    maska_droby = (
        df[SLOUPEC_VYSTUPNI_KATEGORIE]
        == KATEGORIE_DROBY
    )

    pocet_droby = int(maska_droby.sum())

    if pocet_droby == 0:

        LOGGER.info(
            "Všechny hodnoty byly úspěšně zařazeny."
        )

        return

    nezname_hodnoty = (
        df.loc[
            maska_droby,
            config.sloupec_kategorie,
        ]
        .dropna()
        .unique()
    )

    LOGGER.warning(
        "Do kategorie DROBY bylo zařazeno %s záznamů.",
        pocet_droby,
    )

    LOGGER.warning(
        "Neznámé hodnoty neo2dr: %s",
        ", ".join(sorted(map(str, nezname_hodnoty))),
    )


def formatuj_datumy(
    df: pd.DataFrame,
    sloupce: list[str],
) -> None:
    """
    Ověří existenci datumových sloupců,
    zvaliduje hodnoty a převede datumy
    do formátu DD.MM.RRRR.

    Raises
    ------
    ValueError
        Pokud některý z datumových sloupců neexistuje.
    """

    chybejici = [
        sloupec
        for sloupec in sloupce
        if sloupec not in df.columns
    ]

    if chybejici:
        raise ValueError(
            "Chybí datumové sloupce: "
            f"{', '.join(chybejici)}"
        )

    for sloupec in sloupce:

        puvodni_hodnoty = df[sloupec]

        datumy = pd.to_datetime(
            puvodni_hodnoty,
            errors="coerce",
        )

        neplatne_maska = (
            puvodni_hodnoty.notna()
            & datumy.isna()
        )

        neplatnych = int(
            neplatne_maska.sum()
        )

        if neplatnych:

            LOGGER.warning(
                "Sloupec %s obsahuje %s neplatných datumů.",
                sloupec,
                neplatnych,
            )

            LOGGER.warning(
                "Ukázka neplatných hodnot: %s",
                ", ".join(
                    map(
                        str,
                        puvodni_hodnoty[
                            neplatne_maska
                        ]
                        .astype(str)
                        .unique()[:10]
                    )
                ),
            )

        df[sloupec] = datumy.dt.strftime(
            "%d.%m.%Y"
        )

    LOGGER.info(
        "Datumové sloupce byly úspěšně zpracovány."
    )


def exportuj_kategorie(
    df: pd.DataFrame,
    config: Config,
) -> None:
    """
    Exportuje jednotlivé kategorie.

    Kategorie DOVOL je navíc rozdělena podle
    hodnoty neo2pl:

        neo2pl = 0 -> DOVOLENA
        neo2pl = 2 -> PULDOVOL
        ostatní   -> DOVOL_NEPOVOL
    """

    pocet_souboru = 0

    for kategorie, data in df.groupby(
        SLOUPEC_VYSTUPNI_KATEGORIE,
        sort=True,
    ):

        # -------------------------------------------------
        # SPECIÁLNÍ ZPRACOVÁNÍ DOVOLENÉ (neo2dr = 95)
        # -------------------------------------------------

        if kategorie == "DOVOL":

            rozdeleni = {
                "DOVOLENA":
                    data[
                        data[
                            config.sloupec_rozdeleni_dovolena
                        ] == 0
                    ],

                "PULDOVOL":
                    data[
                        data[
                            config.sloupec_rozdeleni_dovolena
                        ] == 2
                    ],

                "DOVOL_NEPOVOL":
                    data[
                        ~data[
                            config.sloupec_rozdeleni_dovolena
                        ].isin([0, 2])
                    ],
            }

            for nazev, subset in rozdeleni.items():

                if subset.empty:
                    continue

                vystupni_soubor = (
                    config.output_dir
                    / f"f60_FILTER_{nazev}.xlsx"
                )

                LOGGER.info(
                    "Exportuji %s (%s řádků)",
                    nazev,
                    len(subset),
                )

                subset.to_excel(
                    vystupni_soubor,
                    index=False,
                )

                LOGGER.info(
                    "Vytvořen soubor: %s",
                    vystupni_soubor,
                )

                pocet_souboru += 1

            continue

        # -------------------------------------------------
        # STANDARDNÍ EXPORT OSTATNÍCH KATEGORIÍ
        # -------------------------------------------------

        vystupni_soubor = (
            config.output_dir
            / f"f60_FILTER_{kategorie}.xlsx"
        )

        LOGGER.info(
            "Exportuji %s (%s řádků)",
            kategorie,
            len(data),
        )

        data.to_excel(
            vystupni_soubor,
            index=False,
        )

        LOGGER.info(
            "Vytvořen soubor: %s",
            vystupni_soubor,
        )

        pocet_souboru += 1

    LOGGER.info(
        "Export dokončen. Vytvořeno %s souborů.",
        pocet_souboru,
    )


def main() -> None:

    config = Config(
        vstupni_soubor=Path(
            "vstup_f60/07_DATA/f60.xlsx"
        ),
        output_dir=Path("vstup_f60"),
        sloupec_kategorie="neo2dr",
        sloupec_datum_od="neo2za",
        sloupec_datum_do="neo2ko",
        sloupec_rozdeleni_dovolena= "neo2pl",
    )

    lookup_kategorii = vytvor_lookup_kategorii()

    LOGGER.info("Spouštím zpracování F60")

    if not config.vstupni_soubor.exists():

        raise FileNotFoundError(
            f"Soubor neexistuje: "
            f"{config.vstupni_soubor}"
        )

    try:

        df = pd.read_excel(
            config.vstupni_soubor,
            engine="openpyxl",
        )

        if df.empty:

            raise ValueError(
                "Vstupní soubor neobsahuje data."
            )

        LOGGER.info(
            "Načteno %s řádků.",
            len(df),
        )

        validuj_a_normalizuj_kategorii(
            df,
            config.sloupec_kategorie,
            strict=False,
        )

        validuj_sloupce(
            df,
            config,
        )

        df[SLOUPEC_VYSTUPNI_KATEGORIE] = (
            df[config.sloupec_kategorie]
            .map(lookup_kategorii)
            .fillna(KATEGORIE_DROBY)
        )

        validuj_droby(
            df,
            config,
        )

        formatuj_datumy(
            df,
            [
                config.sloupec_datum_od,
                config.sloupec_datum_do,
            ],
        )

        LOGGER.info(
            "Rozložení kategorií:\n%s",
            df[
                SLOUPEC_VYSTUPNI_KATEGORIE
            ].value_counts(),
        )

        config.output_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        exportuj_kategorie(
            df,
            config,
        )

        LOGGER.info(
            "Skript byl úspěšně dokončen."
        )

    except Exception:

        LOGGER.exception(
            "Při zpracování došlo k neočekávané chybě."
        )

        raise


if __name__ == "__main__":
    main()