import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import numpy as np

RENAME_COLUMNS = {
   "остатки": {
        "Остаток на складе": "Остаток",
        "Себестоимость": "Себестоимость сумма",
        "Стоимость ( в розничных ценах)": "Сумма остатков в РЦ"
    },
    "продажи": {
        "Количество товаров": "Продажи",
        "Сумма продаж со скидкой": "Сумма продаж в РЦ"
    },
    "прайс": {
        "Номенклатура.Марка (Бренд)": "Бренд",
        "Номенклатура.Категория": "Категория",
        "Номенклатура.Сезон": "Сезон",
        "Цена (тг.)": "Прайс за ед.",
        "Марка (Бренд)": "Бренд"
    }
}

def find_header_row(path, keywords=["Магазин", "Номенклатура", "Характеристика"], max_rows=20):
    raw_df = pd.read_excel(path, header=None, nrows=max_rows)
    for i, row in raw_df.iterrows():
        values = row.astype(str).str.strip().tolist()
        if all(keyword in values for keyword in keywords):
            return i
    raise ValueError(f"Не удалось найти строку заголовков по ключевым словам: {keywords}")

def clean_file(path, тип_файла):
    try:
        header_row = find_header_row(path)
        if not isinstance(header_row, int):
            raise TypeError(f"Ожидался тип int для header_row, но получено: {type(header_row)}")
        print(f"Заголовки найдены на строке: {header_row}")
        df = pd.read_excel(path, header=header_row)
        print(f"Загружено: {df.shape[0]} строк, {df.shape[1]} столбцов")
        df.dropna(axis=0, how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        print(f"После очистки: {df.shape[0]} строк, {df.shape[1]} столбцов")
        rename_map = RENAME_COLUMNS.get(тип_файла, {})
        df = df.rename(columns=rename_map)
        print(f"Переименованные колонки: {list(df.columns)}")
        required = ["Магазин", "Номенклатура", "Характеристика"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError(f"Отсутствуют обязательные колонки: {missing}")
        return df
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        raise

class DataProcessor:
    def __init__(self):
        self.stock_df = None
        self.sales_df = None
        self.price_df = None

    def load_stock(self, path):
        self.stock_df = clean_file(path, тип_файла="остатки")

    def load_sales(self, path):
        self.sales_df = clean_file(path, тип_файла="продажи")

    def load_price(self, path):
        self.price_df = clean_file(path, тип_файла="прайс")

    def generate_summary(self):
        # Проверка загрузки всех таблиц
        if self.stock_df is None:
            raise ValueError("Файл с остатками не загружен")
        if self.sales_df is None:
            raise ValueError("Файл с продажами не загружен")
        if self.price_df is None:
            raise ValueError("Файл с прайсом не загружен")

        print("📦 Остатки:", self.stock_df.shape)
        print("📈 Продажи:", self.sales_df.shape)
        print("💰 Прайс:", self.price_df.shape)

        # Объединение таблиц
        df = self.stock_df.merge(
            self.sales_df,
            on=["Магазин", "Номенклатура", "Характеристика"],
            how="outer"
        )

        df_all = self.price_df.merge(
            df,
            on=["Магазин", "Номенклатура", "Характеристика"],
            how="left"
        )

        df_all = df_all.drop_duplicates(subset=["Магазин", "Номенклатура", "Характеристика"])

        columns_to_keep = [
            "Магазин", "Бренд", "Номенклатура", "Характеристика",
            "Категория", "Сезон",
            "Остаток", "Себестоимость сумма", "Продажи", "Прайс за ед.",
            "Сумма продаж в РЦ", "Сумма остатков в РЦ"
        ]

        missing = [col for col in columns_to_keep if col not in df_all.columns]
        if missing:
            raise ValueError(f"Отсутствуют ожидаемые колонки: {missing}")

        df_all = df_all[columns_to_keep]

        # Приводим к числовому типу
        df_all["Остаток"] = pd.to_numeric(df_all["Остаток"], errors="coerce")
        df_all["Себестоимость сумма"] = pd.to_numeric(df_all["Себестоимость сумма"], errors="coerce")

        # Расчет себестоимости за ед.
        df_all["Себестоимостьза ед."] = np.where(
            df_all["Остаток"] > 0,
            (df_all["Себестоимость сумма"] / df_all["Остаток"])
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0)
                .round(0)
                .astype(int),
            0
        )

        desired_order = [
            "Магазин", "Бренд", "Категория", "Сезон",
            "Номенклатура", "Характеристика",
            "Остаток", "Себестоимость сумма", "Себестоимостьза ед.", "Продажи", "Прайс за ед.",
            "Сумма продаж в РЦ", "Сумма остатков в РЦ"
        ]

        df_all = df_all[desired_order]

        sort_columns = ["Магазин", "Бренд", "Категория", "Номенклатура", "Характеристика"]
        for col in sort_columns:
            if col not in df_all.columns:
                raise ValueError(f"Не хватает колонки для сортировки: {col}")

        df_all = df_all.sort_values(by=sort_columns)

        return df_all


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Распределение заказов")

        # DataFrames
        self.stock_df = None
        self.sales_df = None
        self.price_df = None
        self.min_stock_df = None
        self.df_all = None  # объединённые данные
        self.root = root
        self.root.title("Обработка остатков и продаж")
        self.root.geometry("400x400")
        self.processor = DataProcessor()
        

        # === Кнопки для загрузки файлов ===
        tk.Button(root, text="📦 Загрузить остатки", command=self.load_stock, width=30).pack(pady=5)
        tk.Button(root, text="📈 Загрузить продажи", command=self.load_sales, width=30).pack(pady=5)
        tk.Button(root, text="📋 Загрузить прайс-лист", command=self.load_price, width=30).pack(pady=5)
        tk.Button(root, text="⚙ Загрузить минимальные остатки", command=self.load_min_stock, width=30).pack(pady=5)

        tk.Label(root, text="").pack()  # Пустой отступ

        # Поле для ввода количества дней
        self.days_label = tk.Label(root, text="Период прогноза (в днях):")
        self.days_label.pack()

        self.days_entry = tk.Entry(root)
        self.days_entry.insert(0, "14")
        self.days_entry.pack()

        # === Кнопки действий ===
        self.calc_button = tk.Button(root, text="Рассчитать заказ и сохранить", command=self.calculate_order)
        self.calc_button.pack(pady=10)

        self.dist_button = tk.Button(root, text="📦 Подсорт с Центрального склада", command=self.save_distribution)
        self.dist_button.pack(pady=10)

    # === Загрузка файлов ===
    def load_stock(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            self.processor.load_stock(path) # здесь clean_file вызовется автоматически
            self.stock_df = self.processor.stock_df
            messagebox.showinfo("Файл загружен", f"Остатки загружены: {len(self.stock_df)} строк")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить остатки:\n{e}")

    def load_sales(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            self.processor.load_sales(path) # здесь clean_file вызовется автоматически
            self.sales_df = self.processor.sales_df
            messagebox.showinfo("Файл загружен", f"Продажи загружены: {len(self.sales_df)} строк")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить продажи:\n{e}")

    def load_price(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            self.processor.load_price(path) # здесь clean_file вызовется автоматически
            self.price_df = self.processor.price_df
            messagebox.showinfo("Файл загружен", f"Прайс-лист загружен: {len(self.price_df)} строк")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить прайс-лист:\n{e}")

    def load_min_stock(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            self.min_stock_df = pd.read_excel(path)
            if not {"Категория", "min stock", "max прием"}.issubset(self.min_stock_df.columns):
                raise ValueError("В файле должны быть колонки: Категория, min stock, max прием")
            messagebox.showinfo("Файл загружен", f"Минимальные остатки загружены: {len(self.min_stock_df)} строк")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить минимальные остатки:\n{e}")

    # === Расчёт заказа ===
    def calculate_order(self):
        if self.stock_df is None or self.sales_df is None or self.price_df is None:
            messagebox.showerror("Ошибка", "Сначала загрузите остатки, продажи и прайс-лист")
            return

        try:
            days = int(self.days_entry.get())
            if days <= 0:
                raise ValueError("Введите положительное число дней")
            
            # 🔹 передаём данные в процессор
            self.processor.stock_df = self.stock_df
            self.processor.sales_df = self.sales_df
            self.processor.price_df = self.price_df

            df = self.processor.generate_summary()  # Берём объединённый датафрейм через DataProcessor

            # Расчет заказа
            df["Заказ на период"] = df.apply(
                lambda row: (row.get("Продажи", 0) / 7.0) * days - row.get("Остаток", 0),
                axis=1
            )

            # Комментарии
            def comment(row):
                if row["Остаток"] == 0 and row.get("Продажи", 0) == 0 and row["Заказ на период"] == 0:
                    return "Отправить минимальное количество"
                elif row["Заказ на период"] < 0:
                    return "Излишек"
                elif row["Заказ на период"] > 0:
                    return "Дозаказ"
                else:
                    return ""

            df["Комментарий"] = df.apply(comment, axis=1)

            # Минимальные остатки
            if self.min_stock_df is not None:
                df = df.merge(
                    self.min_stock_df[["Категория", "max прием"]],
                    on="Категория",
                    how="left"
                )
                df.loc[df["Комментарий"] == "Отправить минимальное количество", "Заказ на период"] = \
                    df.loc[df["Комментарий"] == "Отправить минимальное количество", "max прием"].fillna(0)
                df.drop(columns=["max прием"], inplace=True, errors="ignore")

            self.df_all = df

            # Сохраняем файл
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Сохранить файл заказа"
            )
            if path:
                df.to_excel(path, index=False)
                messagebox.showinfo("Сохранено", f"Файл сохранен:\n{path}")
                try:
                    os.startfile(path)
                except Exception:
                    pass

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось рассчитать заказ:\n{e}")

    # === Распределение с Центрального склада ===
    def safe_int(self, value):
        """Преобразует значение в целое число, заменяя NaN и ошибки на 0"""
        num = pd.to_numeric(value, errors="coerce")
        return int(0 if pd.isna(num) else num)


    def build_distribution(self, df):
        priority_stores = [
            "Гранд парк",
            "Азия парк Астана",
            "Шымкент «Love is mama»",
            "Aport East",
            "Aport West",
            "ГЦРЧ"
        ]

        central_df = df[df["Магазин"] == "Центральный склад"]
        result_rows = []

        for _, central_row in central_df.iterrows():
            central_stock = self.safe_int(central_row.get("Остаток", 0))

            row_data = {
                "Категория": central_row.get("Категория", ""),
                "Сезон": central_row.get("Сезон", ""),
                "Бренд": central_row.get("Бренд", ""),
                "Номенклатура": central_row.get("Номенклатура", ""),
                "Характеристика": central_row.get("Характеристика", ""),
                "Откуда": "Центральный склад",
                "Начальное кол-во у отправителя": central_stock,
            }

            for store in priority_stores:
                store_row = df[
                    (df["Магазин"] == store) &
                    (df["Номенклатура"] == central_row["Номенклатура"]) &
                    (df["Характеристика"] == central_row["Характеристика"])
                ]

                if store_row.empty:
                    row_data["{} Начальный остаток".format(store)] = 0
                    row_data["{} Количество заказа".format(store)] = 0
                    row_data["{} Конечный остаток".format(store)] = 0
                else:
                    store_row = store_row.iloc[0]
                    start_stock = self.safe_int(store_row.get("Остаток", 0))

                    if store_row.get("Комментарий") == "Дозаказ":
                        need = self.safe_int(store_row.get("Заказ на период", 0))
                    else:
                        need = 0

                    give = min(central_stock, need)
                    central_stock -= give

                    row_data["{} Начальный остаток".format(store)] = start_stock
                    row_data["{} Количество заказа".format(store)] = give
                    row_data["{} Конечный остаток".format(store)] = start_stock + give

            result_rows.append(row_data)

        return pd.DataFrame(result_rows)

    def save_distribution(self):
        if self.df_all is None:
            messagebox.showerror("Ошибка", "Сначала рассчитайте заказ")
            return

        try:
            dist_df = self.build_distribution(self.df_all)
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Сохранить файл Подсорт"
            )
            if path:
                dist_df.to_excel(path, index=False)
                messagebox.showinfo("Сохранено", f"Файл сохранен:\n{path}")
                try:
                    os.startfile(path)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось построить распределение:\n{e}")


if __name__ == "__main__":
    print("Старт программы")  # Проверка запуска
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()
