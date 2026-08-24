import logging
import pandas as pd
from pathlib import Path
from typing import TypedDict

from dataclasses import dataclass


# ================================

# Architektura               9,6/10
# Čitelnost                  9,7/10
# Dokumentace                9,6/10
# Logging                    9,3/10
# Python styl                9,5/10
# Robustnost                 9,5/10
# Udržovatelnost             9,8/10

logger = logging.getLogger(__name__)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(name)s | %(levelname)s | %(message)s",
    datefmt="%d.%m.%Y %H:%M:%S"
    )



"""
Porovnání výstupů PDNYV a F60.

Skript načte dva zdrojové soubory (PDNYV a F60), porovná jejich
záznamy podle zaměstnance, dne, příjmu a typu pracovního vztahu
a vytvoří Excel s výsledky porovnání.

Výstup obsahuje listy:
    SHODNE
        Záznamy nalezené v obou souborech.

    POUZE_pdnyv
        Záznamy nalezené pouze v PDNYV, které mají vazbu
        na interní zaměstnance.

    POUZE_pdnyv_POUZE_EXT
        Záznamy nalezené pouze v PDNYV s vazbou na externí osoby.

    POUZE_f60
        Záznamy nalezené pouze ve F60 s vazbou
        na interní zaměstnance.

    POUZE_f60_POUZE_EXT
        Záznamy nalezené pouze ve F60 s vazbou na externí osoby.

    Verze:


    1.1

    Datum vytvoření:
        2026-08-14

    Požadavky:
        pandas
        openpyxl

"""

SOUBOR0_DOCHZARV2a="vstup_dochzarv2/07_DATA/dochzarv_07A_FILTER_vztahorg_NOT3.xlsx"
SOUBOR0_DOCHZARV2b="vstup_dochzarv2/07_DATA/dochzarv_07A_FILTER_vztahorg_3.xlsx"

@dataclass
class VysledkyPorovnani:
    shodne: pd.DataFrame

    pouze_pdnyv_interni: pd.DataFrame
    pouze_pdnyv_externi: pd.DataFrame

    pouze_f60_interni: pd.DataFrame
    pouze_f60_externi: pd.DataFrame

    pouze_pdnyv: pd.DataFrame
    pouze_f60: pd.DataFrame

class KonfiguraceFiltru(TypedDict):
    soubor1: str
    soubor2: str
    vystup: str
    razitkovnice: bool

KONFIGURACE: dict[str, KonfiguraceFiltru] = {
    "PLACV": {
        "soubor1": "porovnavač/porovnavač_DROBY_PLACV/07_DATA/pdnyv_07A_26_droby_FILTER_PLACV.xlsx",
        "soubor2": "porovnavač/porovnavač_DROBY_PLACV/07_DATA/f60_FILTER_PLACV.xlsx",
        "vystup": "porovnavač/porovnavač_DROBY_PLACV/vysledek_porovnani_",
        "razitkovnice": False,
    },

    "SVZOZ": {
        "soubor1": "porovnavač/porovnavač_DROBY_SVZOZ/07_DATA/pdnyv_07A_26_droby_FILTER_SVZOZ.xlsx",
        "soubor2": "porovnavač/porovnavač_DROBY_SVZOZ/07_DATA/f60_FILTER_SVZOZ.xlsx",
        "vystup": "porovnavač/porovnavač_DROBY_SVZOZ/vysledek_porovnani_",
        "razitkovnice": False,
    },

    "LEKAR": {
        "soubor1": "porovnavač/porovnavač_LEKAR/07_DATA/pdnyv_07A_26_FILTER_lékař_11_46.xlsx",
        "soubor2": "porovnavač/porovnavač_LEKAR/07_DATA/f60_FILTER_LEKAR.xlsx",
        "vystup": "porovnavač/porovnavač_LEKAR/vysledek_porovnani_",
        "razitkovnice": False,
    },

    "DOVOL": {
        "soubor1": "porovnavač/porovnavač_DOVOL/07_DATA/pdnyv_07A_26_FILTER_dovol_11_46.xlsx",
        "soubor2": "porovnavač/porovnavač_DOVOL/07_DATA/vystupf60_DOVOL_revize.xlsx",
        "vystup": "porovnavač/porovnavač_DOVOL/vysledek_porovnani_",
        "razitkovnice": True,
    },
   
    "PULDOVOL": {
        "soubor1": "porovnavač/porovnavač_PULDOVOL/07_DATA/pdnyv_07A_26_FILTER_PULDOVOL_11_12.xlsx",
        "soubor2": "porovnavač/porovnavač_PULDOVOL/07_DATA/f60_PULDOVOL.xlsx",
        "vystup": "porovnavač/porovnavač_PULDOVOL/vysledek_porovnani_",
        "razitkovnice": False,
     },

    "INDIV": {
        "soubor1": "porovnavač/porovnavač_INDIV/07_DATA/pdnyv_07A_26_FILTER_indiv_11_46.xlsx",
        "soubor2": "porovnavač/porovnavač_INDIV/07_DATA/vystupf60_INDIV.xlsx",
        "vystup": "porovnavač/porovnavač_INDIV/vysledek_porovnani_",
        "razitkovnice": True,
    },

    "STUDV": {
        "soubor1": "porovnavač/porovnavač_STUDV/07_DATA/pdnyv_07A_26_droby_FILTER_STUDV.xlsx",
        "soubor2": "porovnavač/porovnavač_STUDV/07_DATA/vystupf60_STUDV.xlsx",
        "vystup": "porovnavač/porovnavač_STUDV/vysledek_porovnani_",
        "razitkovnice": True,
    },

    "ABSEN": {
        "soubor1": "porovnavač/porovnavač_ABSEN/07_DATA/pdnyv_07A_26_FILTER_absen_11_46.xlsx",
        "soubor2": "porovnavač/porovnavač_ABSEN/07_DATA/f60_FILTER_ABSEN.xlsx",
        "vystup": "porovnavač/porovnavač_ABSEN/vysledek_porovnani_",
        "razitkovnice": True,
    },

    "DROBY": {
            "soubor1": "porovnavač/porovnavač_f60_DROBY/07_DATA/25_dni_BEZ_ZAZNAMU.xlsx",
            "soubor2": "porovnavač/porovnavač_f60_DROBY/07_DATA/vystupf60_DROBY.xlsx",
            "vystup": "porovnavač/porovnavač_f60_DROBY/vysledek_porovnani_",
            "razitkovnice": True,
        },
}

TAB0_SLOUPEC1 = "oscis"
TAB0_SLOUPEC2 = "ppvdr"

#FILTRY_KE_ZPRACOVANI = ["DOVOL","INDIV","PLACV", "ABSEN","STUDV", "PULDOVOL", "LEKAR", "SVZOZ"]
#FILTRY_KE_ZPRACOVANI = ["DOVOL","INDIV","PLACV", "ABSEN","STUDV",  "LEKAR", "SVZOZ"]
#FILTRY_KE_ZPRACOVANI = ["DOVOL", "PULDOVOL"]
FILTRY_KE_ZPRACOVANI = ["DROBY"]

def nacti_data(soubor0a: str, soubor0b: str, soubor1: str, soubor2: str) -> tuple[pd.DataFrame,
                                                                                  pd.DataFrame,
                                                                                  pd.DataFrame,
                                                                                  pd.DataFrame] | None:
    
    """
    Načte vstupní Excel soubory potřebné pro porovnání.

    Parametry
    ---------
    soubor0a : str
        Cesta k souboru DOCHZARV2 obsahujícímu interní
        pracovní vztahy zaměstnanců.

    soubor0b : str
        Cesta k souboru DOCHZARV2 obsahujícímu externí
        pracovní vztahy zaměstnanců.

    soubor1 : str
        Cesta ke zdrojovému souboru PDNYV.

    soubor2 : str
        Cesta ke zdrojovému souboru F60.

    Návratová hodnota
    -----------------
    tuple[
        pd.DataFrame,
        pd.DataFrame,
        pd.DataFrame,
        pd.DataFrame
    ]
        Načtené tabulky ve pořadí:

        (
            dochzarv2_interni,
            dochzarv2_externi,
            pdnyv,
            f60
        )

    None
        Pokud během načítání dojde k chybě.
    """


    logger.info("Načítám vstupní soubory")

    try:

        tab0a = pd.read_excel(soubor0a)
        tab0b = pd.read_excel(soubor0b)

        tab1 = pd.read_excel(soubor1)
        tab2 = pd.read_excel(soubor2)

    except Exception as e:

        logger.error(
            "Chyba při načítání vstupních souborů: %s",e
        )

        return None

    return tab0a, tab0b, tab1, tab2


def validuj_data(tab0a: pd.DataFrame, tab0b: pd.DataFrame, tab1: pd.DataFrame, tab2: pd.DataFrame, klice_pdnyv: list[str], klice_f60: list[str]) -> bool:
    
    """
    Provede kontrolu vstupních dat před porovnáním.

    Ověřuje:
        - neprázdnost tabulek,
        - existenci povinných sloupců,
        - správnost dat ve sloupcích s datem,
        - duplicity podle porovnávacích klíčů.

    Parametry
    ---------
    tab0a : pd.DataFrame
        Tabulka DOCHZARV2 interní.

    tab0b : pd.DataFrame
        Tabulka DOCHZARV2 externí.

    tab1 : pd.DataFrame
        Tabulka PDNYV.

    tab2 : pd.DataFrame
        Tabulka F60.

    klice_pdnyv : list[str]
        Seznam porovnávacích klíčů pro PDNYV.

    klice_f60 : list[str]
        Seznam porovnávacích klíčů pro F60.

    Návratová hodnota
    -----------------
    bool
        True pokud jsou data v pořádku,
        jinak False.
    """
 
    sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1 = klice_pdnyv

    sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2 = klice_f60


    if not over_prazdnou_tabulku(tab0a, "DOCHZARV2 interní"):
                        
        return False

    if not over_prazdnou_tabulku(tab0b, "DOCHZARV2 externí"):
        return False

    if not over_prazdnou_tabulku(tab1, "PDNYV"):
        return False

    if not over_prazdnou_tabulku(tab2, "F60"):
        return False

    logger.info(
        "Načteno PDNYV=%s řádků, F60=%s řádků",
        len(tab1),
        len(tab2)
    )

    if not over_sloupce(tab0a, ["oscis", "ppvdr"], "DOCHZARV2 interní"):
        return False

    if not over_sloupce(tab0b, ["oscis", "ppvdr"], "DOCHZARV2 externí"):
        return False

    if not over_sloupce( tab1, ["oscis", "den", "prijm", "pracvmes"], "PDNYV"):
        return False

    if not over_sloupce(tab2, [sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],"F60"):
        return False

    tab1[sloupec2_tab1] = over_datumy(
        tab1,
        sloupec2_tab1,
        "PDNYV"
    )

    tab2[sloupec2_tab2] = over_datumy(
        tab2,
        sloupec2_tab2,
        "F60"
    )

    tab1[sloupec2_tab1] = (
        tab1[sloupec2_tab1].dt.date
    )

    tab2[sloupec2_tab2] = (
        tab2[sloupec2_tab2].dt.date
    )

    if over_duplicity(
        tab1,
        [
            sloupec1_tab1,
            sloupec2_tab1,
            sloupec3_tab1,
            sloupec4_tab1
        ],
        "PDNYV"
    ) > 0:

        #logger.error(
        #    "Zpracování ukončeno kvůli duplicitám v PDNYV"
        #)

        return False

    if over_duplicity(
        tab2,
        [
            sloupec1_tab2,
            sloupec2_tab2,
            sloupec3_tab2,
            sloupec4_tab2
        ],
        "F60"
    ) > 0:

        logger.error(
            "Zpracování ukončeno kvůli duplicitám v F60"
        )

        return False
    
    logger.info("Validace dat proběhla úspěšně")

    return True

def over_duplicity(df: pd.DataFrame, sloupce: list[str], nazev: str) -> int:
    
    """
    Zkontroluje duplicity podle zadaných sloupců.

    Parametry
    ---------
    df : pd.DataFrame
        Kontrolovaná tabulka.

    sloupce : list[str]
        Sloupce tvořící unikátní klíč.

    nazev : str
        Název tabulky pro logging.

    Návratová hodnota
    -----------------
    int
        Počet duplicitních řádků.
    """

    duplicity = df.duplicated(
        subset=sloupce,
        keep=False
    )

    pocet = duplicity.sum()

    if pocet > 0:

        logger.debug(
            "Ukázka duplicit v %s:\n%s",
            nazev,
            df.loc[duplicity, sloupce].head(10)
        )

        logger.warning(
            "%s obsahuje %s duplicitních řádků podle %s",
            nazev,
            pocet,
            ", ".join(sloupce)
        )

        

    return pocet

def over_datumy(df: pd.DataFrame, sloupec: str, nazev: str) -> pd.Series:

    """
    Převede hodnoty ve sloupci na datumový typ.

    Současně identifikuje neplatné nebo chybně zadané
    hodnoty data a zapíše jejich výskyt do logu.

    Parametry
    ---------
    df : pd.DataFrame
        Kontrolovaná tabulka.

    sloupec : str
        Název sloupce obsahujícího datum.

    nazev : str
        Název tabulky pro logging.

    Návratová hodnota
    -----------------
    pd.Series
        Sloupec převedený na datetime.

        Neplatné hodnoty jsou nahrazeny hodnotou NaT.
    """

    puvodni = df[sloupec].copy()

    prevedeno = pd.to_datetime(
        puvodni,
        errors="coerce",
        format="%d.%m.%Y"
    )

    chybne = puvodni[
                 prevedeno.isna()
                & puvodni.notna()
                & (puvodni.astype(str).str.strip() != "")
                    ]

    if len(chybne) > 0:
        logger.warning(
            "%s obsahuje %s neplatných datumů ve sloupci %s. "
            "Příklady: %s",
            nazev,
            len(chybne),
            sloupec,
            chybne.head(5).tolist()
        )

    return prevedeno

def over_prazdnou_tabulku(df: pd.DataFrame, nazev: str) -> bool:
    
    """
    Ověří, zda tabulka obsahuje alespoň jeden řádek.

    Parametry
    ---------
    df : pd.DataFrame
        Kontrolovaná tabulka.

    nazev : str
        Název tabulky pro logging.

    Návratová hodnota
    -----------------
    bool
        True pokud tabulka není prázdná,
        jinak False.
    """

    if df.empty:
        logger.warning(
            "%s neobsahuje žádná data",
            nazev
        )
        return False
    
    return True

def over_sloupce(df: pd.DataFrame, povinne_sloupce: list[str], nazev: str) -> bool:
    
    """
    Ověří existenci povinných sloupců v tabulce.

    Parametry
    ---------
    df : pd.DataFrame
        Kontrolovaná tabulka.

    povinne_sloupce : list[str]
        Seznam požadovaných sloupců.

    nazev : str
        Název tabulky pro logging.

    Návratová hodnota
    -----------------
    bool
        True pokud jsou všechny sloupce přítomny,
        jinak False.
    """

    missing = set(povinne_sloupce) - set(df.columns)

    if missing:
        logger.error(
            "%s - chybí sloupce: %s",
            nazev,
            ", ".join(sorted(missing))
        )
        return False

    return True

def porovnej_data(tab0a: pd.DataFrame, tab0b: pd.DataFrame, 
                  tab1: pd.DataFrame, tab2: pd.DataFrame,
                  klice_pdnyv: list[str],klice_f60: list[str]) -> VysledkyPorovnani:
    
    
    """
    Provede porovnání záznamů PDNYV a F60.


    Parametry
    ---------
    tab0a : pd.DataFrame
        Tabulka DOCHZARV2 obsahující interní pracovní vztahy.
        Používá se k identifikaci záznamů, které mají vazbu
        na interní zaměstnance.

    tab0b : pd.DataFrame
        Tabulka DOCHZARV2 obsahující externí pracovní vztahy.
        Používá se k identifikaci záznamů, které mají vazbu
        na externí osoby.

    tab1 : pd.DataFrame
        Zdrojová tabulka PDNYV obsahující porovnávané
        záznamy zaměstnanců.

    tab2 : pd.DataFrame
        Zdrojová tabulka F60 obsahující porovnávané
        záznamy zaměstnanců.

    klice_pdnyv : list[str]
        Seznam názvů sloupců tvořících porovnávací klíč
        v tabulce PDNYV.

        Očekávané pořadí:
            [
                osobní číslo,
                datum,
                příjmení,
                pracoviště, 
                druh pracovní smlouvu (ppvdr)
            ]

    klice_f60 : list[str]
        Seznam názvů sloupců tvořících porovnávací klíč
        v tabulce F60.

        Očekávané pořadí:
            [
                osobní číslo,
                datum,
                příjmení,
                pracoviště, 
                druh pracovní smlouvu (ppvdr)
            ]
        
    Návratová hodnota
    -----------------

    VysledkyPorovnani

    Objekt obsahující všechny výsledky porovnání.


        shodne
            Záznamy nalezené v obou zdrojích.

        pouze_pdnyv_interni
            Záznamy nalezené pouze v PDNYV
            s vazbou na interní pracovní vztahy.

        pouze_pdnyv_externi
            Záznamy nalezené pouze v PDNYV
            s vazbou na externí pracovní vztahy.

        pouze_f60_interni
            Záznamy nalezené pouze ve F60
            s vazbou na interní pracovní vztahy.

        pouze_f60_externi
            Záznamy nalezené pouze ve F60
            s vazbou na externí pracovní vztahy.

    """

    sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1 = klice_pdnyv

    sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2 = klice_f60


    # Nalezení záznamů, které existují v obou zdrojích
    shodne = pd.merge(
            tab1,
            tab2,
            left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1],
            right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],
            how="inner",
            #validate="one_to_one"
        )

    if len(shodne) == 0:
        logger.warning("Nenalezen žádný shodný záznam")


    # Záznamy existující pouze v PDNYV
    pouze_tab1 = pd.merge(
            tab1,
            tab2,
            left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1],
            right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],
            how="left",
            #validate="one_to_one",
            indicator=True
        )
   
    
    pouze_tab1 = pouze_tab1[
        pouze_tab1["_merge"] == "left_only"
    ][[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1]]


    pouze_tab1a = pd.merge(
                        pouze_tab1,
                        tab0a[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                        left_on=sloupec1_tab1,
                        right_on=TAB0_SLOUPEC1,
                        #how="left"
                        how="inner",
                        #validate="one_to_one"
                                )

    
    pouze_tab1b = pd.merge(
                    pouze_tab1,
                    tab0b[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab1,
                    right_on=TAB0_SLOUPEC1,
                    how="right",
                    #validate="one_to_one"
                            )
    logger.debug("pouze_tab1:\n%s", pouze_tab1.head(20))
    # Záznamy existující pouze ve F60
    pouze_tab2 = pd.merge(
            tab1,
            tab2,
            left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1],
            right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],
            how="right",
            #validate="one_to_one",
            indicator=True
        )
   

    pouze_tab2 = pouze_tab2[
        pouze_tab2["_merge"] == "right_only"
    ][[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2]]
     
    logger.debug("Obsah pouze_tab2:\n%s", pouze_tab2)

    pouze_tab2a = pd.merge(
                    pouze_tab2,
                    tab0a[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab2,
                    right_on=TAB0_SLOUPEC1,
                    #how="left"
                    how="inner",
                    #validate="one_to_one"
                            )
    
    pouze_tab2b = pd.merge(
                    pouze_tab2,
                    tab0b[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab2,
                    right_on=TAB0_SLOUPEC1,
                    how="right",
                    #validate="one_to_one"
                            )

    logger.debug( "Ukázka pouze_tab2a:\n%s", pouze_tab2a.head(10) )
    
    shodne["den"] = pd.to_datetime(shodne["den"]).dt.strftime("%d.%m.%Y")

    pouze_tab1a["den"] = pd.to_datetime(pouze_tab1a["den"]).dt.strftime("%d.%m.%Y")
    pouze_tab1b["den"] = pd.to_datetime(pouze_tab1b["den"]).dt.strftime("%d.%m.%Y")
   
    pouze_tab2a[sloupec2_tab2 ] = pd.to_datetime(pouze_tab2a[sloupec2_tab2 ]).dt.strftime("%d.%m.%Y")
    pouze_tab2b[sloupec2_tab2 ] = pd.to_datetime(pouze_tab2b[sloupec2_tab2 ]).dt.strftime("%d.%m.%Y")

    logger.info("Výsledek porovnání | SHODNE=%s | "
                 "PDNYV_INT=%s | PDNYV_EXT=%s | "
                 "F60_INT=%s | F60_EXT=%s",
                 len(shodne),
                len(pouze_tab1a), len(pouze_tab1b),
                len(pouze_tab2a), len(pouze_tab2b),)
    return VysledkyPorovnani(
        shodne=shodne,
        pouze_pdnyv_interni=pouze_tab1a,
        pouze_pdnyv_externi=pouze_tab1b,
        pouze_f60_interni=pouze_tab2a,
        pouze_f60_externi=pouze_tab2b,
        pouze_pdnyv=pouze_tab1,
        pouze_f60=pouze_tab2,
        )


def uloz_vysledky(soubor: str, vysledky: VysledkyPorovnani) -> bool:

    """
    Uloží výsledky porovnání do Excel souboru.

    Parametry
    ---------
    soubor : str
        Cesta k výstupnímu Excel souboru.

    vysledky : VysledkyPorovnani
        Objekt obsahující všechny výsledky porovnání.

    Návratová hodnota
    -----------------
    bool
        True při úspěšném uložení,
        jinak False.
    """

    logger.info("Vytvářím výstupní soubor: %s", soubor)

    try:
        with pd.ExcelWriter(
            soubor,
            engine="openpyxl"
        ) as writer:

            vysledky.shodne.to_excel(
                writer,
                sheet_name="SHODNE",
                index=False
            )

            vysledky.pouze_pdnyv_interni.to_excel(
                writer,
                sheet_name="POUZE_pdnyv",
                index=False
            )

            vysledky.pouze_pdnyv_externi.to_excel(
                writer,
                sheet_name="POUZE_pdnyv_POUZE_EXT",
                index=False
            )

            vysledky.pouze_f60_interni.to_excel(
                writer,
                sheet_name="POUZE_f60",
                index=False
            )

            vysledky.pouze_f60_externi.to_excel(
                writer,
                sheet_name="POUZE_f60_POUZE_EXT",
                index=False
            )

    except Exception as e:
        logger.error(
            "Nepodařilo se uložit soubor: %s",
            e
        )
        return False
    
    logger.info( "Soubor úspěšně vytvořen: %s", soubor)

    logger.info("Hotovo: %s", soubor)

    return True

def porovnavac_f60(druh_filtru: str, soubor0a: str, soubor0b: str, cfg: KonfiguraceFiltru) -> None:
    
    
    """
    Řídicí funkce zajišťující kompletní porovnání PDNYV a F60.

    Postup:
        1. Kontrola existence souborů.
        2. Načtení dat.
        3. Validace dat.
        4. Porovnání záznamů.
        5. Uložení výsledků.

    Parametry
    ---------
    druh_filtru : str
        Název zpracovávaného filtru
        (např. DOVOL, INDIV, PLACV).

    soubor0a : str
        Cesta k souboru DOCHZARV2 interní.

    soubor0b : str
        Cesta k souboru DOCHZARV2 externí.

    cfg : KonfiguraceFiltru
        Konfigurace zpracovávaného filtru.

    Návratová hodnota
    -----------------
    None
    """

    logger.info("=" * 60)

    logger.info("Zahajuji zpracování filtru: %s", druh_filtru)
    
    soubor1 = cfg["soubor1"]
    soubor2 = cfg["soubor2"]
    soubor3 = cfg["vystup"]

    pouzita_razitkovnice = cfg["razitkovnice"]

    logger.info("Konfigurace | filtr=%s | razitkovnice=%s", druh_filtru, pouzita_razitkovnice)

    sloupec1_tab1 = "oscis"
    sloupec2_tab1 = "den"
    sloupec3_tab1 = "prijm"
    sloupec4_tab1= "pracvmes"

    sloupec1_tab2 = "oscis"
    sloupec2_tab2 = "neo2za"
    sloupec3_tab2 = "prijm"
    sloupec4_tab2= "pracv"

    if pouzita_razitkovnice:
        sloupec2_tab2 = "aktual_den"

    # ================================
    # NAČTENÍ DAT
    # ================================

    soubory = {
        "DOCHZARV2 interní": soubor0a,
        "DOCHZARV2 externí": soubor0b,
        "PDNYV": soubor1,
        "F60": soubor2
            }

    for nazev, cesta in soubory.items():
        if not Path(cesta).is_file():
            logger.error("%s - soubor neexistuje: %s", nazev, cesta)
            return
      
    nactena_data = nacti_data(soubor0a, soubor0b, soubor1, soubor2)

    if nactena_data is None:
        return

    tab0a, tab0b, tab1, tab2 = nactena_data
    
    logger.info("Načteny tabulky | DOCH_INT=%s | DOCH_EXT=%s | PDNYV=%s | F60=%s",
                 len(tab0a), len(tab0b), len(tab1), len(tab2)
                 )
    KLICE_PDNYV = [
        sloupec1_tab1,
        sloupec2_tab1,
        sloupec3_tab1,
        sloupec4_tab1
    ]

    KLICE_F60 = [
        sloupec1_tab2,
        sloupec2_tab2,
        sloupec3_tab2,
        sloupec4_tab2
    ]
    
    
    if not validuj_data(tab0a, tab0b, tab1, tab2, KLICE_PDNYV, KLICE_F60):
        return


    vysledky = porovnej_data( tab0a, tab0b, tab1, tab2, KLICE_PDNYV, KLICE_F60,)

    
    logger.info(
        "Filtr %s | Shodné=%s | Pouze PDNYV=%s | Pouze F60=%s",
        druh_filtru,
        len(vysledky.shodne),
        len(vysledky.pouze_pdnyv),
        len(vysledky.pouze_f60)
        )
    

    # ================================
    # ULOŽENÍ
    # ================================
    

    CELE_JMENO_SOUBORU=soubor3+druh_filtru+".xlsx"
      
    if not uloz_vysledky(CELE_JMENO_SOUBORU, vysledky):
        return   
    

    logger.info("Dokončeno zpracování filtru: %s", druh_filtru)

    logger.info("=" * 60)
       

if __name__ == "__main__":

    for filtr in FILTRY_KE_ZPRACOVANI:
        


        cfg = KONFIGURACE[filtr]

        porovnavac_f60(
            filtr,
            SOUBOR0_DOCHZARV2a,
            SOUBOR0_DOCHZARV2b,
            cfg
        )
