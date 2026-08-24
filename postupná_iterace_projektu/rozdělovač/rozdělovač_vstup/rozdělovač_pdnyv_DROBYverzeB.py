from dataclasses import dataclass, field
from pathlib import Path
import logging

import pandas as pd

# Architektura               9,5/10 
# Čitelnost                  9,4/10 |
# Dokumentace                9,9/10 |
# Logging                    9,5/10 |
# Python styl                9,5/10 |
# Robustnost                 9,5/10 |
# Udržovatelnost             9,6/10 |


"""
Rozdělení Excelového souboru podle hodnot ve sloupci `duvodt`.

Skript:
1. Načte vstupní Excel soubor.
2. Ověří přítomnost požadovaných sloupců.
3. Vyfiltruje záznamy podle definovaných kategorií.
4. Každou kategorii uloží do samostatného Excel souboru.

Výstupní soubory jsou ukládány do adresáře OUTPUT_DIR
ve formátu:

    <BASE_FILENAME>_FILTER_<KATEGORIE>.xlsx

    Vstup:
        pdnyv_06_26_FILTER_droby_12_35.xlsx

    Výstupy:
        pdnyv_06_26_droby_FILTER_CERNV.xlsx
        pdnyv_06_26_droby_FILTER_HOMOF.xlsx
        pdnyv_06_26_droby_FILTER_OCR.xlsx
        pdnyv_06_26_droby_FILTER_OTECD.xlsx
        pdnyv_06_26_droby_FILTER_PLACV.xlsx
        pdnyv_06_26_droby_FILTER_STUDV.xlsx        
        pdnyv_06_26_droby_FILTER_SVZOZ.xlsx

Autor: Majda Tomáš

Datum vytvoření:
    2026-07-20

Požadavky:
    pandas
    openpyxl
    xlsxwriter

"""



# ==========================================
# Konfigurace
# ==========================================


@dataclass(frozen=True)
class Config:
    # Zdrojový Excel soubor určený k rozdělení.
    input_file: Path = Path(
        "porovnavač/2026-07-17_12-35/"
        "pdnyv_06_26_FILTER_droby_12_35.xlsx"
        )
    # Adresář pro ukládání výsledných souborů.
    output_dir: Path = Path(
        "porovnavač/2026-07-17_12-35"
        )
    # Základ názvu generovaných souborů
    base_filename: str = "pdnyv_06_26_droby"
    # Sloupec používaný pro filtrování kategorií.
    filter_column: str = "duvodt"
    #export_columns: list[str]
    # Mapování:
    # název výstupního souboru -> hodnota ve sloupci duvodt
    categories: dict[str, str] = field(
        default_factory=lambda:
        {
        "OCR": ["očr"],       # Ošetřování člena rodiny
        "STUDV": ["studv"],   # Studium
        "OTECD": ["otecd"],   # Otcovská dovolená
        "SVZOZ": ["svzoz"],   # Služební volno k zařízení osobních záležitostí
        "PLACV": ["placv", "svpro"],   # Placené volno
        "CERNV": ["cernv"],   # Čerpání náhradního volna
        "HOMOF": ["homof"],   # Home Office
                }
                )

    export_columns: list[str] = field(

        # Exportované sloupce:
        #
        # oscis    - osobní číslo
        # den      - datum
        # prijm    - příjmení
        # dt       - den v týdnu
        # duvod    - důvod nepřítomnosti (kód)
        # duvodt   - důvod nepřítomnosti (text)
        # duvod2   - druhý důvod nepřítomnosti (kód)
        # duvod2t  - druhý důvod nepřítomnosti (text)
        # pracvmes - pracoviště
        # priche   - příchod
        # odche    - odchod

        default_factory=lambda:
          [
            "oscis",   
            "den",     
            "prijm",   
            "dt",      
            "duvod",   
            "duvodt",  
            "duvod2", 
            "duvod2t", 
            "pracvmes",
            "priche",  
            "odche",   
        ]
        )

config = Config()





# ==========================================
# Logging
# ==========================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

logger = logging.getLogger(__name__)


# ==========================================
# Funkce
# ==========================================

def over_konfiguraci(config: Config) -> None:
    """
    Ověří správnost konfiguračních hodnot.
    """

    if not config.categories:
        raise ValueError("KATEGORIE nesmí být prázdné.")

    if not config.export_columns:
        raise ValueError("ZACHOVANE_SLOUPCE nesmí být prázdné.")

    if not config.filter_column:
        raise ValueError("SLOUPEC_FILTRU není nastaven.")
    """
    hodnoty = list(config.categories.values())

    duplicitni_hodnoty = {
        hodnota
        for hodnota in hodnoty
        if hodnoty.count(hodnota) > 1
    }
    """

    vsechny_hodnoty = [
        hodnota
        for hodnoty in config.categories.values()
        for hodnota in hodnoty
    ]

    duplicitni_hodnoty = {
        hodnota
        for hodnota in vsechny_hodnoty
        if vsechny_hodnoty.count(hodnota) > 1
        }

    if duplicitni_hodnoty:
        raise ValueError(
            "Duplicitní hodnoty filtrů: "
            f"{sorted(duplicitni_hodnoty)}"
        )

    logger.info(
        "Konfigurace byla úspěšně ověřena."
    )

def over_vstupni_data(df: pd.DataFrame, config: Config) -> None:    
    """
    Ověří, že vstupní DataFrame obsahuje všechny
    sloupce potřebné pro filtrování a export.

    Args:
        df:
            Načtená data ze vstupního Excel souboru.

    Raises:
        ValueError:
            Pokud chybí některý z požadovaných sloupců.

    """

    if df.empty:
        raise ValueError("Vstupní soubor neobsahuje žádná data.")

    pozadovane = set(config.export_columns + [config.filter_column])

    
    chybejici = pozadovane - set(df.columns)

    if chybejici:
        raise ValueError(
            f"Ve vstupním souboru chybí sloupce: "
            f"{', '.join(sorted(chybejici))}"
        )
    
    nalezene = set(df[config.filter_column].dropna().unique())

    logger.info(
        "Nalezené hodnoty ve sloupci %s: %s", 
        config.filter_column, 
        sorted(nalezene),
    )

def normalizuj_data(df: pd.DataFrame, config: Config) -> None:
    """
    Normalizuje hodnoty používané pro filtrování.

    Odstraňuje okolní mezery a převádí text
    na malá písmena.
    """

    df[config.filter_column] = ( 
        df[config.filter_column]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower())


def over_kategorie(
    df: pd.DataFrame, 
    config: Config
) -> None:    
    """
    Ověří, zda se ve vstupních datech
    nevyskytují neznámé hodnoty ve sloupci
    SLOUPEC_FILTRU.

    Neznámé hodnoty jsou zalogovány jako warning.
    """

    #zname_kategorie = set(config.categories.values())

    
    zname_kategorie = {
        hodnota
        for hodnoty in config.categories.values()
        for hodnota in hodnoty
        }


    nezname_kategorie = (
        set(df[config.filter_column].dropna().unique())
        - zname_kategorie
    )

    if nezname_kategorie:
        logger.warning(
            "Nalezeny neznámé hodnoty ve sloupci %s: %s",
            config.filter_column,
            sorted(nezname_kategorie),
        )

def uloz_kategorii(
    df: pd.DataFrame, 
    vystupni_adresar: Path, 
    suffix: str, 
    hodnota_filtru: list[str],
    config: Config
) -> int:    
    """
    Vyfiltruje záznamy podle hodnoty ve sloupci
    SLOUPEC_FILTRU a uloží výsledek do samostatného
    Excel souboru.

    Název souboru je vytvořen ve tvaru:

        <BASE_FILENAME>_FILTER_<suffix>.xlsx

    Args:
        df:
            Zdrojový DataFrame.

        vystupni_adresar:
            Adresář pro uložení výsledného souboru.

        suffix:
            Označení kategorie použité v názvu souboru.

        hodnota_filtru:
            Hodnota vyhledávaná ve sloupci SLOUPEC_FILTRU.
   

    """


    #df_filtered = df[df[config.filter_column] == hodnota_filtru]

    df_filtered = df[
                    df[config.filter_column].isin(hodnota_filtru)
                    ]

    pocet_radku = len(df_filtered)

    """
    logger.info(
        "Kategorie %s (%s): %s řádků",
        suffix,
        hodnota_filtru,
        pocet_radku,
    )
    """
    logger.info(
        "Kategorie %s (%s): %s řádků",
        suffix,
        ", ".join(hodnota_filtru),
        pocet_radku,
        )
    

    if pocet_radku == 0:
        logger.warning(
            "Kategorie %s neobsahuje žádné záznamy.",
            suffix,
        )

    vystupni_soubor = ( 
        vystupni_adresar 
        / f"{config.base_filename}_FILTER_{suffix}.xlsx")

    chybejici = set(config.export_columns) - set(df_filtered.columns)

    if chybejici:
        raise ValueError(f"Chybí exportované sloupce: {chybejici}")


    try:
        with pd.ExcelWriter(
            vystupni_soubor, 
            engine="xlsxwriter", 
            datetime_format="hh:mm",
            ) as writer:

            df_filtered[config.export_columns].to_excel(
                writer,
                index=False
            )
    
    
    except PermissionError:
        logger.error("Soubor %s je pravděpodobně otevřen v Excelu.", vystupni_soubor,)
        return 0


    logger.info(
        "Soubor uložen: %s",
        vystupni_soubor,
    )
    
    return pocet_radku

def main(config: Config) -> None:    
    """
    Řídí celý proces zpracování.

    Postup:
        1. Ověří konfiguraci.
        2. Ověří vstupní soubor.
        3. Načte data.
        4. Ověří vstupní data.
        5. Normalizuje data.
        6. Ověří kategorie.
        7. Exportuje soubory.
    """

    celkem_exportovano = 0
    over_konfiguraci(config)
    vstupni_soubor = config.input_file

    
    print(config.input_file)
    print(type(config.input_file))


    if not vstupni_soubor.exists():
        raise FileNotFoundError(
            f"Vstupní soubor neexistuje: {vstupni_soubor}"
        )


    if vstupni_soubor.suffix.lower() != ".xlsx":
        raise ValueError("Očekáván XLSX soubor.")

    config.output_dir.mkdir(parents=True, exist_ok=True)

    logger.info("Načítám soubor %s", vstupni_soubor)

    try:
        df = pd.read_excel(
            vstupni_soubor,
            engine="openpyxl",
            )

    except Exception:
        logger.exception(
            "Nepodařilo se načíst Excel soubor."
        )
        raise

    logger.info(
        "Načteno %s řádků a %s sloupců",
        len(df),
        len(df.columns),
    )

    over_vstupni_data(df, config)
    normalizuj_data(df, config)

    over_kategorie(df, config)

    for suffix, hodnota in config.categories.items():
        exportovano = uloz_kategorii(
                                df=df,
                                vystupni_adresar=config.output_dir,
                                suffix=suffix,
                                hodnota_filtru=hodnota,
                                config=config
                            )

    
        celkem_exportovano += exportovano

    logger.info("Exportováno %s z %s řádků.", celkem_exportovano, len(df),)

    if celkem_exportovano != len(df):
        logger.warning(
            "Nevyexportováno %s řádků.",
            len(df) - celkem_exportovano,
        )


    logger.info("Zpracování dokončeno.")


if __name__ == "__main__":
    try:
        main(config)
        
    except Exception:
        logger.exception("Neočekávaná chyba při zpracování.")
        raise
