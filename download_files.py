import pandas as pd
from typing import Literal

KEY_COLUMNS = {"Магазин", "Номенклатура", "характеристика"}

# Переименовываем оригинальные названия на удобные
RENAME_COLUMNS = {
    "остатки": {
        "Остаток на складе": "Остаток",
        "Себестоимость": "Себестоимость сумма",
        "Стоимость ( в розничных ценах)": "Сумма остатков в РЦ"
    },
    "продажи": {
        "Количество товаров": "Продажи",
        "Сумма продаж со скидкой": "Сумма продаж в РЦ"
    }
}

# Какие столбцы должны быть после переименования
VALUE_COLUMNS = {
    "остатки": {"Остаток", "Себестоимость сумма", "Сумма остатков в РЦ"},
    "продажи": {"Продажи", "Сумма продаж в РЦ"}
}


def load_1c_file(path: str, тип_файла: Literal["остатки", "продажи"]) -> pd.DataFrame:
    if тип_файла not in VALUE_COLUMNS:
        raise ValueError(f"Неизвестный тип файла: {тип_файла}")
    
    preview = pd.read_excel(path, header=None, nrows=20)
    header_raw = None
    for i, row in preview.iterrows():
        values = set(str(cell).strip() for cell in row if pd.notna(cell)) # преобразуем каждую ячейку строки в строку, убираем пробелы и сключаем пустые

        if KEY_COLUMNS.issubset(values): # проверяем ключевые заголовки в текущей строке
            header_raw = i
            break                        # если строка с заголовками найдена - прерываем цикл- дальше искать не нужно

        if header_raw is None:           # если ни одна строка не содержала все ключевые названия - выбрасываем исключение - защита от некорректных файлов
            raise ValueError("Не найдена строка с заголовками")
    
    df_raw = pd.read_excel(path, skiprows=0) #пропускаем служебные строки
    df_clean = (df_raw
                .dropna(how='all') #удаляем полностью пустые строки
                .dropna(axis=1, how='all')
                 )             # удаляем пустые столбцы
    df.columns = df.columns.str.strip()
    df_clean.columns = df_clean.columns.str.strip() # чистим заголовки от пробелов
    
    # переименования колонок
    rename_map = RENAME_COLUMNS.get(тип_файла, {})
    df = df.rename(columns = rename_map)

    # проверка ключевых колонок
    missing_keys = KEY_COLUMNS - set(df.columns)
    if missing_keys:
        raise ValueError(f"Отсутусвуют ключевые колонки: {missing_keys}")
    
    # проверка значимых колонок
    expected_values = VALUE_COLUMNS[тип_файла]
    missing_values = expected_values - set(df.columns)
    if missing_values:
        raise ValueError(f"Отсутствуют столбцы для типа '{тип_файла}")

    return df

df_ostatki = load_1c_file("остатки.xlsx", тип_файла="остатки")
df_prodazhi = load_1c_file("продажи.xlsx", тип_файла="продажи")

# Шаг 1: объединение
df_all = pd.contact([df_ostatki, df_prodazhi], ignore_index=True)

#Шаг 2: Расчет себестоимости и РЦ за единицу
df_all["Себестоимость за ед."] = df_all.apply(
    lambda row: row["Себестоимость сумма"] / row["Остаток"] if row["Остаток"] else None, axis=1
)
df_all["РЦ за ед."] = df_all.apply(
    lambda row: row["Сумма продаж в РЦ"] / row["Продажи"] if row["Продажи"] else None, axis=1
)

# Шаг 3: Удаление ненужных столбцов
df_all = df_all.drop(columns=["Сумма остатков в РЦ", "Сумма продаж в РЦ", "Себестоимость сумма"])

# Шаг 4: сводная таблица
KEY_COLUMNS = ["Магазин", "Номенклатура", "Характеристика"]
VALUE_COLUMNS = ["Остаток", "Продажи", "Себестоимость за ед.", "РЦ за ед."]

df_pivot = df_all.groupby(KEY_COLUMNS, dropna=False)[VALUE_COLUMNS].sum().reset_index()

df_pivot["Расчет заказа на 4 недели"] = df_pivot.apply(
    lambda row: (row["Продажи"] / 7 * 28 - row["Остаток"]) if pd.notna(row["Остаток"]) else None,
    axis=1
)

def flag(row):
    order = row["Расчет заказа на 4 недели"]
    stock = row["Остаток"]
    sales = row["Продажи"]

    if pd.notna(order) and order <= -1:
        return "отдает"
    elif pd.notna(order) and order > 0:
        return "принимает"
    elif order == 0 and stock == 0 and sales == 0:
        return "принимает по минималке"
    else:
        return None

df_pivot["Флаг заказа"] = df_pivot.apply(flag, axis=1)

def enrich_with_price(df_pivot, path_to_price):
    # 1. загрузка прас-листа с заголовками на 5й строке
    df_price = pd.read_excel(path_to_price, header=4)

    # 2. переименование колонок
    df_price = df_price.rename(columns = {
       "Номенклатура.Марка (Бренд)": "Бренд",
        "Номенклатура.Категория": "Категория",
        "Номенклатура.Сезон": "Сезон",
        "Цена (тг.)": "Прайс за ед."
    })

    # 3. удаление ненужных колонок
    df_price = df_price.drop(columns=["Единица измерения"], errors="ignore")

    # 4. Оставляем тлько нужные столбцы
    df_price_subset = df_price[["Номенклатура", "Категория", "Бренд", "Сезон"]]

    df_final = pd.marge(
        df_pivot,
        df_price_subset,
        on = "Номенклатура",
        how = "left"
    )

