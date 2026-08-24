import logging
import pandas as pd


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
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
"""

# Architektura               6,0/10
# Čitelnost                  7,0/10
# Dokumentace                3,0/10
# Logging                    2,0/10
# Python styl                5,0/10
# Robustnost                 4,0/10
# Udržovatelnost             5,0/10

# ================================
# NASTAVENÍ (SEM DOSADÍTE NÁZVY)
# ================================

#soubor1 = "porovnavač/porovnavač_DROBY_PLACV/06_DATA/pdnyv_meziclen_05_droby_FILTER_PLACV.xlsx"
#soubor2 = "porovnavač/porovnavač_DROBY_PLACV/06_DATA/f60_FILTER_PLACV.xlsx"

SOUBOR0_DOCHZARV2a="vstup_dochzarv2/2026-07-09_14-18/dochzarv2_06_26_FILTER_vztahorg_NOT3_14_18.xlsx"
SOUBOR0_DOCHZARV2b="vstup_dochzarv2/2026-07-09_14-18/dochzarv2_06_26_FILTER_vztahorg_14_18.xlsx"

FILTER_PLACV="PLACV"
SOUBOR1_PLACV_IN = "porovnavač/porovnavač_DROBY_PLACV/06_DATA/pdnyv_meziclen_05_droby_FILTER_PLACV.xlsx"
SOUBOR2_PLACV_IN = "porovnavač/porovnavač_DROBY_PLACV/06_DATA/f60_FILTER_PLACV.xlsx"
SOUBOR1_PLACV_OUT = "porovnavač/porovnavač_DROBY_PLACV/vysledek_porovnani_"

FILTER_SVZOZ="SVZOZ"
SOUBOR1_SVZOZ_IN = "porovnavač/porovnavač_DROBY_SVZOZ/06_DATA/pdnyv_meziclen_05_droby_FILTER_SVZOZ.xlsx"
SOUBOR2_SVZOZ_IN = "porovnavač/porovnavač_DROBY_SVZOZ/06_DATA/f60_FILTER_SVZOZ.xlsx"
SOUBOR1_SVZOZ_OUT = "porovnavač/porovnavač_DROBY_SVZOZ/vysledek_porovnani_"

FILTER_LEKAR="LEKAR"
SOUBOR1_LEKAR_IN = "porovnavač/porovnavač_LEKAR/06_DATA/pdnyv_meziclen_05_lékař_09_46.xlsx"
SOUBOR2_LEKAR_IN = "porovnavač/porovnavač_LEKAR/06_DATA/f60_FILTER_LEKAR.xlsx"
SOUBOR1_LEKAR_OUT = "porovnavač/porovnavač_LEKAR/vysledek_porovnani_"

FILTER_DOVOL="DOVOL"
SOUBOR1_DOVOL_IN = "porovnavač/porovnavač_DOVOL/06_DATA/pdnyv_06_26_FILTER_dovol_11_12.xlsx"
SOUBOR2_DOVOL_IN = "porovnavač/porovnavač_DOVOL/06_DATA/vystupf60_DOVOL.xlsx"
SOUBOR1_DOVOL_OUT = "porovnavač/porovnavač_DOVOL/vysledek_porovnani_"

FILTER_PULDOVOL="PULDOVOL"
SOUBOR1_PULDOVOL_IN = "porovnavač/porovnavač_PULDOVOL/06_DATA/pdnyv_06_26_FILTER_PULDOVOL_11_12.xlsx"
SOUBOR2_PULDOVOL_IN = "porovnavač/porovnavač_PULDOVOL/06_DATA/vystupf60_PULDOVOL.xlsx"
SOUBOR1_PULDOVOL_OUT = "porovnavač/porovnavač_PULDOVOL/vysledek_porovnani_"

FILTER_INDIV="INDIV"
SOUBOR1_INDIV_IN = "porovnavač/porovnavač_INDIV/06_DATA/pdnyv_meziclen_05_indiv_09_46.xlsx"
SOUBOR2_INDIV_IN = "porovnavač/porovnavač_INDIV/06_DATA/vystupf60_INDIV.xlsx"
SOUBOR1_INDIV_OUT = "porovnavač/porovnavač_INDIV/vysledek_porovnani_"

FILTER_STUDV="STUDV"
SOUBOR1_STUDV_IN = "porovnavač/porovnavač_STUDV/06_DATA/pdnyv_meziclen_05_droby_FILTER_STUDV.xlsx"
SOUBOR2_STUDV_IN = "porovnavač/porovnavač_STUDV/06_DATA/vystupf60_DROBY_STUDV.xlsx"
SOUBOR1_STUDV_OUT = "porovnavač/porovnavač_STUDV/vysledek_porovnani_"

FILTER_ABSEN="ABSEN"
SOUBOR1_ABSEN_IN = "porovnavač/porovnavač_ABSEN/06_DATA/pdnyv_meziclen_absen_09_46.xlsx"
SOUBOR2_ABSEN_IN = "porovnavač/porovnavač_ABSEN/06_DATA/vystupf60_ABSEN.xlsx"
SOUBOR1_ABSEN_OUT = "porovnavač/porovnavač_ABSEN/vysledek_porovnani_"

def porovnavac_f60( druh_filtru, soubor0a, soubor0b, soubor1, soubor2, pouzita_razitkovnice, soubor3):
    
    """
    Provede porovnání výstupů PDNYV a F60.

    Parametry
    ---------
    druh_filtru : str
        Název porovnávané kategorie (např. DOVOL, PLACV, LEKAR).

    soubor0a : str
        Excel s vazbou zaměstnance na interní pracovní vztahy.

    soubor0b : str
        Excel s vazbou zaměstnance na externí pracovní vztahy.

    soubor1 : str
        Zdrojový soubor PDNYV.

    soubor2 : str
        Zdrojový soubor F60.

    pouzita_razitkovnice : bool
        Určuje, který datumový sloupec se použije ve F60.

        False:
            používá se sloupec 'neo2za'

        True:
            používá se sloupec 'aktual_den'

    soubor3 : str
        Cesta a prefix výsledného souboru.

    Výstup
    ------
    Vytvoří Excel soubor:
        <soubor3><druh_filtru>.xlsx

    se seznamem shodných a rozdílových záznamů.
    """


    TAB0_SLOUPEC1="oscis"
    TAB0_SLOUPEC2="ppvdr"
    TAB0_SLOUPEC3="vztahorg"
    
    sloupec1_tab1 = "oscis"
    sloupec2_tab1 = "den"
    sloupec3_tab1 = "prijm"
    sloupec4_tab1= "pracvmes"

    sloupec1_tab2 = "oscis"
    sloupec2_tab2 = "neo2za"
    sloupec3_tab2 = "prijm"
    sloupec4_tab2= "pracv"

    if pouzita_razitkovnice:
        #sloupec1_tab2 = "oscis"
        sloupec2_tab2 = "aktual_den"

    # ================================
    # NAČTENÍ DAT
    # ================================

    tab0a = pd.read_excel(soubor0a)
    tab0b = pd.read_excel(soubor0b)
    tab1 = pd.read_excel(soubor1)
    tab2 = pd.read_excel(soubor2)



    # Normalizace datumů před porovnáním
    tab1[sloupec2_tab1] = pd.to_datetime(tab1[sloupec2_tab1], errors="coerce", format="%d.%m.%Y")
    tab1[sloupec2_tab1] = tab1[sloupec2_tab1].dt.date

    tab2[sloupec2_tab2 ] = pd.to_datetime(tab2[sloupec2_tab2 ], errors="coerce", format="%d.%m.%Y")
    tab2[sloupec2_tab2] = tab2[sloupec2_tab2].dt.date

    # ================================
    # POROVNÁNÍ
    # ================================

    # Nalezení záznamů, které existují v obou zdrojích
    shodne = pd.merge(
        tab1,
        tab2,
        left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1],
        right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],
        how="inner"
    )

    # Záznamy existující pouze v PDNYV
    pouze_tab1 = pd.merge(
        tab1,
        tab2,
        left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1],
        right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],
        how="left",
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
                    how="inner"
                            )
    
    pouze_tab1b = pd.merge(
                    pouze_tab1,
                    tab0b[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab1,
                    right_on=TAB0_SLOUPEC1,
                    how="right"
                            )

    # Záznamy existující pouze ve F60
    pouze_tab2 = pd.merge(
        tab1,
        tab2,
        left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1, sloupec4_tab1],
        right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2],
        how="right",
        indicator=True
    )

    pouze_tab2 = pouze_tab2[
        pouze_tab2["_merge"] == "right_only"
    ][[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2, sloupec4_tab2]]
     
    print(pouze_tab2)

    pouze_tab2a = pd.merge(
                    pouze_tab2,
                    tab0a[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab2,
                    right_on=TAB0_SLOUPEC1,
                    #how="left"
                    how="inner"
                            )
    
    pouze_tab2b = pd.merge(
                    pouze_tab2,
                    tab0b[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab2,
                    right_on=TAB0_SLOUPEC1,
                    how="right"
                            )

    print(pouze_tab2a)
    shodne["den"] = pd.to_datetime(shodne["den"]).dt.strftime("%d.%m.%Y")

    pouze_tab1a["den"] = pd.to_datetime(pouze_tab1a["den"]).dt.strftime("%d.%m.%Y")
    pouze_tab1b["den"] = pd.to_datetime(pouze_tab1b["den"]).dt.strftime("%d.%m.%Y")
   
    pouze_tab2a[sloupec2_tab2 ] = pd.to_datetime(pouze_tab2a[sloupec2_tab2 ]).dt.strftime("%d.%m.%Y")
    pouze_tab2b[sloupec2_tab2 ] = pd.to_datetime(pouze_tab2b[sloupec2_tab2 ]).dt.strftime("%d.%m.%Y")

    # ================================
    # ULOŽENÍ
    # ================================

    CELE_JMENO_SOUBORU=soubor3+druh_filtru+".xlsx"

    with pd.ExcelWriter(CELE_JMENO_SOUBORU, engine="openpyxl") as writer:
        shodne.to_excel(writer, sheet_name="SHODNE", index=False)
        pouze_tab1a.to_excel(writer, sheet_name="POUZE_pdnyv", index=False)
        pouze_tab1b.to_excel(writer, sheet_name="POUZE_pdnyv_POUZE_EXT", index=False)
        pouze_tab2a.to_excel(writer, sheet_name="POUZE_f60", index=False)
        pouze_tab2b.to_excel(writer, sheet_name="POUZE_f60_POUZE_EXT", index=False)

    print("Hotovo: "+CELE_JMENO_SOUBORU)


if __name__ == "__main__":
    porovnavac_f60( FILTER_DOVOL, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_DOVOL_IN, SOUBOR2_DOVOL_IN, True, SOUBOR1_DOVOL_OUT)
    #porovnavac_f60( FILTER_INDIV, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_INDIV_IN, SOUBOR2_INDIV_IN, True, SOUBOR1_INDIV_OUT)
    #porovnavac_f60( FILTER_STUDV, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_STUDV_IN, SOUBOR2_STUDV_IN, True, SOUBOR1_STUDV_OUT)
      
    
    #porovnavac_f60( FILTER_ABSEN, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_ABSEN_IN, SOUBOR2_ABSEN_IN, True, SOUBOR1_ABSEN_OUT)    
    #porovnavac_f60( FILTER_LEKAR, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_LEKAR_IN, SOUBOR2_LEKAR_IN, False, SOUBOR1_LEKAR_OUT)
    #porovnavac_f60( FILTER_PLACV, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_PLACV_IN, SOUBOR2_PLACV_IN, False, SOUBOR1_PLACV_OUT)
    #porovnavac_f60( FILTER_PULDOVOL, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_PULDOVOL_IN, SOUBOR2_PULDOVOL_IN, True, SOUBOR1_PULDOVOL_OUT)
    #porovnavac_f60( FILTER_SVZOZ, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_SVZOZ_IN, SOUBOR2_SVZOZ_IN, False, SOUBOR1_SVZOZ_OUT)
