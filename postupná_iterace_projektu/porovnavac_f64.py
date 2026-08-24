import pandas as pd

# ================================

# Architektura               6,0/10
# Čitelnost                  8,0/10
# Dokumentace                3,0/10
# Logging                    2,0/10
# Python styl                6,0/10
# Robustnost                 4,0/10
# Udržovatelnost             5,0/10


SOUBOR0_DOCHZARV2a="vstup_dochzarv2/07_DATA/dochzarv_07A_FILTER_vztahorg_NOT3.xlsx"
SOUBOR0_DOCHZARV2b="vstup_dochzarv2/07_DATA/dochzarv_07A_FILTER_vztahorg_3.xlsx"

JMENO_VYSLEDNEHO_SOUBORU="vysledek_porovnani_"
FILTER_OTECD="OTECD"
SOUBOR1_OTECD_IN = "porovnavač/porovnavač_DROBY_OTECD/07_DATA/pdnyv_07A_26_droby_FILTER_OTECD.xlsx"
SOUBOR2_OTECD_IN = "porovnavač/porovnavač_DROBY_OTECD/07_DATA/vystupf64_07A_OTECD.xlsx"
SOUBOR1_OTECD_OUT = "porovnavač/porovnavač_DROBY_OTECD/"

FILTER_OCR="OCR"
SOUBOR1_OCR_IN = "porovnavač/porovnavač_DROBY_OCR/07_DATA/pdnyv_07A_26_droby_FILTER_OCR.xlsx"
SOUBOR2_OCR_IN = "porovnavač/porovnavač_DROBY_OCR/07_DATA/vystupf64_07A_OCR.xlsx"
SOUBOR1_OCR_OUT = "porovnavač/porovnavač_DROBY_OCR/"

FILTER_NEMOC="NEMOC"
SOUBOR1_NEMOC_IN = "porovnavač/porovnavač_NEMOC/07_DATA/pdnyv_07A_26_FILTER_nemoc_10_07.xlsx"
SOUBOR2_NEMOC_IN = "porovnavač/porovnavač_NEMOC/07_DATA/vystupf64_07A_NEMOC.xlsx"
SOUBOR1_NEMOC_OUT = "porovnavač/porovnavač_NEMOC/"

def porovnavac_f64(jmeno_vysledneho_souboru, druh_filtru, soubor0a, soubor0b, soubor1, soubor2, soubor3):
    
    
        #soubor1 = "pdnyv_meziclanek_NEMOC_17_05.xlsx"
        #soubor2 = "vystupf64_NEMOC.xlsx"

        TAB0_SLOUPEC1="oscis"
        TAB0_SLOUPEC2="ppvdr"
        TAB0_SLOUPEC3="vztahorg"

        sloupec1_tab1 = "oscis"
        sloupec2_tab1 = "den"
        sloupec3_tab1 = "prijm"
        sloupec4_tab1 = "pracvmes"

        sloupec1_tab2 = "oscis"
        sloupec2_tab2 = "aktual_den"
        sloupec3_tab2 = "prijm"
        sloupec4_tab2 = "pracv"

        # ================================
        # NAČTENÍ DAT
        # ================================
        tab0a = pd.read_excel(soubor0a)
        tab0b = pd.read_excel(soubor0b)
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
            left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1,sloupec4_tab1],
            right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2,sloupec4_tab2],
            how="left",
            indicator=True
        )

        pouze_tab1 = pouze_tab1[
            pouze_tab1["_merge"] == "left_only"
        ][[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1,sloupec4_tab1]]

        
        pouze_tab1a = pd.merge(
                    pouze_tab1,
                    tab0a[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab1,
                    right_on=TAB0_SLOUPEC1,
                    how="left"
                            )
        
        print(tab0b[TAB0_SLOUPEC3])

        pouze_tab1b = pd.merge(
                    pouze_tab1,
                    tab0b[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab1,
                    right_on=TAB0_SLOUPEC1,
                    how="right"
                            )


        # 3) pouze v TAB2
        pouze_tab2 = pd.merge(
            tab1,
            tab2,
            left_on=[sloupec1_tab1, sloupec2_tab1, sloupec3_tab1,sloupec4_tab1],
            right_on=[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2,sloupec4_tab2],
            how="right",
            indicator=True
        )

        pouze_tab2 = pouze_tab2[
            pouze_tab2["_merge"] == "right_only"
        ][[sloupec1_tab2, sloupec2_tab2, sloupec3_tab2,sloupec4_tab2]]

        pouze_tab2a = pd.merge(
                    pouze_tab2,
                    tab0a[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab2,
                    right_on=TAB0_SLOUPEC1,
                    how="left"
                            )
        print(tab0b[TAB0_SLOUPEC3])

        pouze_tab2b = pd.merge(
                    pouze_tab2,
                    tab0b[[TAB0_SLOUPEC1, TAB0_SLOUPEC2]],
                    left_on=sloupec1_tab2,
                    right_on=TAB0_SLOUPEC1,
                    how="right"
                            )

        # ================================
        # ULOŽENÍ
        # ================================
        CELE_JMENO_SOUBORU=soubor3+jmeno_vysledneho_souboru+druh_filtru+".xlsx"

        with pd.ExcelWriter(CELE_JMENO_SOUBORU, engine="openpyxl") as writer:
            shodne.to_excel(writer, sheet_name="SHODNE", index=False)
            pouze_tab1a.to_excel(writer, sheet_name="POUZE_pdnyv", index=False)
            pouze_tab1b.to_excel(writer, sheet_name="POUZE_pdnyv_POUZE_EXT", index=False)
            pouze_tab2a.to_excel(writer, sheet_name="POUZE_f64", index=False)
            pouze_tab2b.to_excel(writer, sheet_name="POUZE_f64_POUZE_EXT", index=False)

        print("Hotovo: "+CELE_JMENO_SOUBORU)



if __name__ == "__main__":
    porovnavac_f64(JMENO_VYSLEDNEHO_SOUBORU, FILTER_OTECD, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_OTECD_IN, SOUBOR2_OTECD_IN, SOUBOR1_OTECD_OUT)
    porovnavac_f64(JMENO_VYSLEDNEHO_SOUBORU, FILTER_OCR, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_OCR_IN, SOUBOR2_OCR_IN, SOUBOR1_OCR_OUT)
    porovnavac_f64(JMENO_VYSLEDNEHO_SOUBORU, FILTER_NEMOC, SOUBOR0_DOCHZARV2a, SOUBOR0_DOCHZARV2b, SOUBOR1_NEMOC_IN, SOUBOR2_NEMOC_IN, SOUBOR1_NEMOC_OUT)
    