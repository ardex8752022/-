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

def find_header_row(path, keywords = ["Магазин", "Номенклатура", "Характеристика"], max_rows=20):
    raw_df = pd.read_excel(path, header=None, nrows=max_rows)

    for i, row in raw_df.iterrows():
        values = row.astype(str).str.strip().tolist()
        if all(keyword in values for keyword in keywords):
            return i

    raise ValueError(f"Не удалось найти строку заголовков по ключевым словам: {keywords}")
 


def clean_file(path, тип_файла):
    try:
        header_row = find_header_row(path)

        # Проверка типа

        if not isinstance(header_row, int):
            raise TypeError(f"Ожидался тип int для header_row, но получено: {type(header_row)}")
        
        print(f"Заголовки найдены на строке:{header_row}")

        df = pd.read_excel(path, header=header_row)

        print(f"Загружено: {df.shape[0]} строк, {df.shape[1]} столбцов")

        # Удаляем полностью пустые строки и столбцы
        df.dropna(axis=0, how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        print(f"После очистки: {df.shape[0]} строк, {df.shape[1]} столбцов")

        # Переименовываем колонки
        rename_map = RENAME_COLUMNS.get(тип_файла, {})
        df = df.rename(columns = rename_map)
        print(f"Переименованные колонки: {list(df.columns)}")

        # Проверка на обязаьельные поля

        required = ["Магазин", "Номенклатура", "Характеристика"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError(f"Отстутсвуют обязательные колонки: {missing}")

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
        self.stock_df = clean_file(path, тип_файла = "остатки")

    def load_sales(self, path):
        self.sales_df = clean_file(path, тип_файла = "продажи")

    def load_price(self, path):
        self.price_df = clean_file(path, тип_файла = "прайс")

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
            on = ["Магазин", "Номенклатура", "Характеристика"],
            how="outer"
        )

        df_all = self.price_df.merge(
            df,
            on = ["Магазин", "Номенклатура", "Характеристика"],
            how = "left"
        )

        # Удаление дубликатов после объединения
        df_all = df_all.drop_duplicates(subset=["Магазин", "Номенклатура", "Характеристика"])

        # Сохранение нужных столбцов
        columns_to_keep = [
            "Магазин", "Бренд", "Номенклатура", "Характеристика",
            "Категория", "Сезон",
            "Остаток","Себестоимость сумма", "Продажи", "Прайс за ед.",
            "Сумма продаж в РЦ", "Сумма остатков в РЦ"
]      

        missing = [col for col in columns_to_keep if col not in df_all.columns]
        if missing:
            raise ValueError(f"Отсутствуют ожидаемые колонки: {missing}")
        
        df_all = df_all[columns_to_keep]
 
        # 1) Приводим колонки к числовому типу (строки → NaN)
        df_all["Остаток"] = pd.to_numeric(df_all["Остаток"], errors="coerce")
        df_all["Себестоимость сумма"] = pd.to_numeric(df_all["Себестоимость сумма"], errors="coerce")

       
        
        #  Расчет себестоимостиза единицу
        df_all["Себестоимостьза ед."] = np.where(
            df_all["Остаток"] > 0,
            (df_all["Себестоимость сумма"] / df_all["Остаток"])
                    .replace([np.inf, -np.inf], np.nan)  # убираем inf
                    .fillna(0)                           # заменяем NaN
                    .round(0)
                    .astype(int),
                0
        )

          # Перестановка столбцов
        desired_order = [
            "Магазин", "Бренд", "Категория", "Сезон",
            "Номенклатура", "Характеристика",
            "Остаток","Себестоимость сумма", "Себестоимостьза ед.", "Продажи", "Прайс за ед.",
            "Сумма продаж в РЦ", "Сумма остатков в РЦ"
        ]

        df_all = df_all[desired_order]
    
         # Сортировка
        sort_columns = ["Магазин", "Бренд", "Категория", "Номенклатура", "Характеристика"]
        for col in sort_columns:
            if col not in df_all.columns:
                raise ValueError(f"Не хватает колонки для сортировки: {col}")

        df_all = df_all.sort_values(by=sort_columns)


        return df_all

        # Добавить расчет заказа

    
        


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка остатков и продаж")
        self.root.geometry("400x400")
        self.processor = DataProcessor()
        self.df_all = None

        tk.Label(root, text="Загрузка данных", font=("Arial", 14)).pack(pady=10)

        tk.Button(root, text="📦 Загрузить остатки", command=self.load_stock, width=30).pack(pady=5)
        tk.Button(root, text="📈 Загрузить продажи", command=self.load_sales, width=30).pack(pady=5)
        tk.Button(root, text="📋 Загрузить прайс-лист", command=self.load_price, width=30).pack(pady=5)

        tk.Label(root, text="").pack() # Пустой отступ

         # Поле для ввода количества дней
        self.days_label = tk.Label(root, text="Период прогноза(в днях):")
        self.days_label.pack()

        self.days_entry = tk.Entry(root)
        self.days_entry.insert(0, "14") # значение по умолчанию
        self.days_entry.pack()

        # Кнопка для расчёта заказа
        self.calc_button = tk.Button(root, text="Расчитать заказ и сохранить", command=self.calculate_order)
        self.calc_button.pack(pady=10)

        
    def load_stock(self):
        path = filedialog.askopenfilename(title="Выберите файл с остатками")
        if path:
            self.processor.load_stock(path)
            messagebox.showinfo("Успех", "Файл с остатками загружен")

    def load_sales(self):
        path = filedialog.askopenfilename(title="Выберите файл с продажами")
        if path:
            self.processor.load_sales(path)
            messagebox.showinfo("Успех", "Файл с продажами загружен")

    def load_price(self):
        path = filedialog.askopenfilename(title="Выберите прайс-лист")
        if path:
            self.processor.load_price(path)
            messagebox.showinfo("Успех", "Прайс-лист загружен")

    def calculate_order(self):
        try:
            # формируем сводную таблицу
            df = self.processor.generate_summary()

            days = int(self.days_entry.get())
            if days <= 0:
                raise ValueError("Введите положительное число дней")
            
            # расчет заказа
            df["Заказ на период"] = df.apply(
                lambda row: max(0, row["Продажи"] / 7 * days - row["Остаток"])
                if pd.notnull(row.get("Продажи")) and pd.notnull(row.get("Остаток")) else 0,
                axis=1
            )

            self.df_all = df # сохраняем в атрибут

            # сохраняем сразу с колонкой "Заказ на период"
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
            )

            if path:
                df.to_excel(path, index=False)
                messagebox.showinfo("Сохранено", f"Файл сохранен: \n{path}")
                try:
                    os.startfile(path)
                except Exception:
                    pass

        except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось рассчитать заказ:\n{eval}")
        

    def save_summary(self):
        try:
            # Генерация сводной таблицы
            df = self.processor.generate_summary()

            if df is None:
                raise ValueError("Сводная таблица не была создана — проверь загрузку файлов")

            # --- Важно: сохраняем сводную таблицу в атрибут класса ---
            self.df_all = df
             # Разрешаем кнопку расчёта
            self.calc_button.config(state=tk.NORMAL)

            # Сохранение файла
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Сохранить сводную таблицу"
            )

            if path:
                df.to_excel(path, index=False)
                messagebox.showinfo("Сохранено", f"Файл сохранен: \n{path}")
                try:
                    os.startfile(path)
                except Exception:
                    # на случай, если os.startfile недоступен — просто игнорируем
                    pass

            # Сохраняем таблицу в атрибут класса
            
            return df

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
            return None

    
        

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()
