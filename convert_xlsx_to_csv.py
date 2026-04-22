from pathlib import Path
import pandas as pd

"""
Данный скрипт преобразовывает файлы .xlsx в .csv
для упрощения дальнейшей работы с данными.

Ожидаемая структура проекта:
project_root/
│
├── data/
│   ├── data_1.xlsx
│   ├── data_2.xlsx
│   ├── ...
│   └── data_6.xlsx
│
└── convert_xlsx_to_csv.py
"""

# Корень проекта = папка, где лежит этот скрипт
PROJECT_ROOT = Path(__file__).resolve().parent

# Папка с входными Excel-файлами
BASE = PROJECT_ROOT / "data"

# Входные файлы
INPUT_FILES = [BASE / f"data_{i}.xlsx" for i in range(1, 7)]

# Папка для выходных CSV
OUTPUT_DIR = BASE / "csv_from_xlsx"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def parse_params_sheet(path: Path) -> pd.DataFrame:
    """
    Преобразует лист 'Технологические параметры' в long-format:
    d | parameter | value
    """
    raw = pd.read_excel(path, sheet_name="Технологические параметры")
    cols = list(raw.columns)
    parts = []

    i = 0
    while i < len(cols):
        if i + 1 >= len(cols):
            break

        date_col = cols[i]
        value_col = cols[i + 1]

        chunk = raw[[date_col, value_col]].copy()
        chunk.columns = ["d", "value"]
        chunk["parameter"] = value_col
        chunk = chunk.dropna(subset=["d"]).sort_values("d")

        parts.append(chunk[["d", "parameter", "value"]])
        i += 2

    if not parts:
        return pd.DataFrame(columns=["d", "parameter", "value"])

    return pd.concat(parts, ignore_index=True)


def convert_one_file(path: Path) -> None:
    """
    Конвертирует один Excel-файл в набор CSV-файлов.
    """
    stem = path.stem

    wabt = pd.read_excel(path, sheet_name="WABT").sort_values("d")
    params_long = parse_params_sheet(path)
    limit_df = pd.read_excel(path, sheet_name="Ограничение", header=None)

    wabt.to_csv(OUTPUT_DIR / f"{stem}_WABT.csv", index=False, encoding="utf-8-sig")
    params_long.to_csv(OUTPUT_DIR / f"{stem}_tech_params_long.csv", index=False, encoding="utf-8-sig")
    limit_df.to_csv(
        OUTPUT_DIR / f"{stem}_limit_raw.csv",
        index=False,
        header=False,
        encoding="utf-8-sig"
    )

    try:
        limit_value = float(limit_df.iloc[0, 1])
        pd.DataFrame([{"limit_wabt": limit_value}]).to_csv(
            OUTPUT_DIR / f"{stem}_limit.csv",
            index=False,
            encoding="utf-8-sig"
        )
    except (IndexError, ValueError, TypeError):
        print(f"[WARNING] Не удалось извлечь limit_wabt из файла: {path.name}")


def main():
    existing_files = [p for p in INPUT_FILES if p.exists()]

    if not BASE.exists():
        raise FileNotFoundError(
            f"Папка с данными не найдена: {BASE}\n"
            f"Создайте папку 'data' в корне проекта и положите туда Excel-файлы."
        )

    if not existing_files:
        raise FileNotFoundError(
            f"Файлы data_1.xlsx ... data_6.xlsx не найдены в папке:\n{BASE}"
        )

    for path in existing_files:
        print(f"[INFO] Обработка файла: {path.name}")
        convert_one_file(path)

    print(f"\n[OK] CSV-файлы сохранены в папке:\n{OUTPUT_DIR}")
    print("\nСписок созданных файлов:")
    for file in sorted(OUTPUT_DIR.glob("*.csv")):
        print(f" - {file.name}")


if __name__ == "__main__":
    main()