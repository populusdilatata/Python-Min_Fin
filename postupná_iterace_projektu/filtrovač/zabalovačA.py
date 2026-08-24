from pathlib import Path
import shutil

CATEGORIES = [
    "ABSEN",
    "INDIV",
    "LEKAR",
    "DOVOL", 
    "NEMOC", 

    "DROBY_OCR",
    "DROBY_OTECD",
    "DROBY_PLACV",
    "DROBY_STUDV",
    "DROBY_SVZOZ"
]

SOURCE_ROOT = Path("porovnavač")
TARGET_ROOT = Path("PRO_Sko")

for category in CATEGORIES:

    category_folder = SOURCE_ROOT / f"porovnavač_{category}/"

    # Najdi nejnovější složku XX_DATA
    data_folders = sorted(
        [d for d in category_folder.iterdir()
         if d.is_dir() and d.name.endswith("_DATA")]
    )

    if not data_folders:
        print(f"Nebyla nalezena DATA složka pro {category}.")
        continue

    source_folder = data_folders[-1]

    # Cílová složka
    
    # Pro DROBY vytvoř složku "_DROBY", ostatní beze změny
    target_category = f"_{category}" if category.startswith("DROBY") else category

    target_folder = TARGET_ROOT/target_category
    target_folder.mkdir(parents=True, exist_ok=True)

    pattern = f"vysledek_porovnani_{category}*"
    files = list(source_folder.glob(pattern))

    if not files:
        print(f"Soubor {pattern} nebyl nalezen.")
        continue

    for file in files:
        shutil.copy2(file, target_folder)
        print(f"{file.name} -> {target_folder}")