import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

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
        "Цена (тг.)": "Прайс за ед."
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
            raise TypeError(f"Оидался тип int для header_row, но получено:")
        
        print(f"Заголовки найдены на строке:{header_row}")

        df = pd.read_excel(path, header=header_row)

        print(f"Загружено: {df.shape[0]} строк, {df.shape[1]} столбцов")

        # Удаляем полностью пустые строки и столбцы
        df.dropna(axis=0, how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        print(f"После очистки: {df.shape[0]} cnhjr, {df.shape[1]} столбцов")

        # Переименовываем колонки
        rename_map = RENAME_COLUMNS.get(тип_файла, {})
        df = df.rename(columns = rename_map)
        print(f"Переименованные колонки: {list(df.columns)}")

        # Проверка на обязаьельные поля

        required = ["Магазин", "Номенклатура", "Характеристика"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError(f"Отстутсвуют обязательныеколонки: {missing}")

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
            "Магазин", "Номенклатура", "Характеристика",
            "Бренд", "Категория", "Сезон",
            "Остаток","Себестоимость сумма", "Продажи", "Прайс за ед.",
            "Сумма продаж в РЦ", "Сумма остатков в РЦ"
]

        df_all = df_all[columns_to_keep]

        missing = [col for col in columns_to_keep if col not in df_all.columns]
        if missing:
            raise ValueError(f"Отсутсвуют ожидаемые колонки: {missing}")
        
        # Перестановка столбцов
        desired_order = [
            "Магазин", "Бренд", "Категория", "Сезон",
            "Номенклатура", "Характеристика",
            "Остаток","Себестоимость сумма", "Продажи", "Прайс за ед.",
            "Сумма продаж в РЦ", "Сумма остатков в РЦ"
        ]

        df_all = df[desired_order]

        # Сортировка
        sort_columns = ["Магазин", "Бренд", "Категория", "Номенклатура", "Характеристика"]
        df_all = df_all.sort_values(by=sort_columns)


        return df_all
    # Добавить расчет заказа



class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка остатков и продаж")
        self.root.geometry("400x250")
        self.processor = DataProcessor()
        self.df_all = None

        tk.Label(root, text="Загрузка данных", font=("Arial", 14)).pack(pady=10)

        tk.Button(root, text="📦 Загрузить остатки", command=self.load_stock, width=30).pack(pady=5)
        tk.Button(root, text="📈 Загрузить продажи", command=self.load_sales, width=30).pack(pady=5)
        tk.Button(root, text="📋 Загрузить прайс-лист", command=self.load_price, width=30).pack(pady=5)

        tk.Label(root, text="").pack() # Пустой отступ

        tk.Button(root, text="✅ Выгрузить сводную таблицу", command=self.save_summary, width=30).pack(pady=10)


    def load_stock(self):
        path = filedialog.askopenfilename(title="Выберите файл с остатками")
        if path:
            self.processor.load_stock(path)
            messagebox.showinfo("Успех", "Файл с остатками загружен")

    def load_sales(self):
        path = filedialog.askopenfilename(title="Выберите файлс продажами")
        if path:
            self.processor.load_sales(path)
            messagebox.showinfo("Успех", "Файл с продажами загружен")

    def load_price(self):
        path = filedialog.askopenfilename(title="Выберите прайс-лист")
        if path:
            self.processor.load_price(path)
            messagebox.showinfo("Успех", "Прайс-лист загружен")

    def save_summary(self):
        try:
            # Генерация сводной таблицы
            df = self.processor.generate_summary()

            if df is None:
                raise ValueError("Сводная таблица не была создана — проверь загрузку файлов")

            # Сохранение файла
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Сохранить сводную таблицу"
            )

            if path:
                df.to_excel(path, index=False)
                messagebox.showinfo("Сохранено", f"Файл сохранен: \n{path}")
                os.startfile(path)

            # Сохраняем таблицу в атрибут класса
            self.df_all = df
            return df

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
            return None


        

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()
