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
        if all(keywords in values for keyword in keywords):
            return i

    raise ValueError(f"Не удалось найти строку заголовков по ключевым словам: {keywords}")
 


def clean_file(path, тип_файла):
    header_row = find_header_row(path)
    df = pd.read_excel(path, header=header_row)

    # Удаляем полностью пустые строки и столбцы
    df.dropna(axis=0, how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    # Переименовываем колонки
    rename_map = RENAME_COLUMNS.get(тип_файла, {})
    df = df.rename(columns = rename_map)

    return df


class DataProcessor:
    def __init__(self):
        self.stock_df = None
        self.sales_df = None
        self.price_df = None

    def load_stock(self, path):
        self.stock_df = pd.read_excel(path, тип_файла = "остатки")

    def load_sales(self, path):
        self.sales_df = pd.read_excel(path, тип_файла = "продажи")

    def load_price(self, path):
        self.price_df = pd.read_excel(path, тип_файла = "прайс")

    def generate_summary(self):
        if self.stock_df is None or self.sales_df is None or self.price_df is None:
            raise ValueError("Не все файлы загружены")
        
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
        df_all = df_all.drop_duplicate(subset=["Магазин", "Номенклатура", "Характеристика"])
    
    # Добавить расчет заказа



class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка остатков и продаж")
        self.root.geometry("400x250")
        self.processor = DataProcessor()
        self.df_all = None

        tk.Lable(root, text="Загрузка данных", font=("Arial", 14)).pack(pady=10)

        tk.Button(root, text="📦 Загрузить остатки", command=self.load_stock, width=30).pack(pady=5)
        tk.Button(root, text="📈 Загрузить продажи", command=self.load_sales, width=30).pack(pady=5)
        tk.Button(root, text="📋 Загрузить прайс-лист", command=self.load_price, width=30).pack(pady=5)

        tk.Label(root, text="").pack() # Пустой отступ

        tk.Button(root, text="✅ Выгрузить сводную таблицу", )

    def save_summary(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить сводную таблицу"
        )
        if file_path and self.df_all is not None:
            self.df_all.to_excel(file_path, index=False)

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
            df = self.processor.generate_summary()
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if path:
                df.to_excel(path, index=False)
                messagebox.showinfo("Сохранено", f"Файл сохранен: \n{path}")
                os.startfile(path)
        except Exception as e:
            messagebox.showinfo("Ошибка", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()
