from pathlib import Path
import argparse
from excel_to_xml import excel_to_xml
from xml_to_excel import xml_to_excel

def main() -> None:
    p = argparse.ArgumentParser(description="Конвертация Excel ↔ XML (UTD 5.03)")
    p.add_argument("src", type=Path, help="Исходный .xlsx | .xml")
    p.add_argument("-o", "--out", type=Path,
                   help="Имя выходного файла (по умолчанию рядом с исходником)")
    args = p.parse_args()

    src: Path = args.src
    if not src.exists():
        p.error(f"Файл не найден: {src}")

    out = args.out or src.with_suffix(".xml" if src.suffix.lower() in (".xls", ".xlsx") else ".xlsx")

    if src.suffix.lower() in (".xls", ".xlsx"):
        excel_to_xml(src, out)
    elif src.suffix.lower() == ".xml":
        xml_to_excel(src, out)
    else:
        p.error("Допустимы только .xlsx/.xls и .xml")

if __name__ == "__main__":
    main()
