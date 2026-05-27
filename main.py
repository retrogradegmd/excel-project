import pandas as pd
import os
import re
from pathlib import Path
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment

BASE_DIR = Path(__file__).resolve().parent
FOLDER_PATH = BASE_DIR / "Gmail"  
OUTPUT_FILE = BASE_DIR / "report.xlsx"

COL_INDEX = 'Наименование'
COL_VALUES_PURCHASE = 'Стоимость покупок с НДС\n\n(стр. 170)'
COL_VALUES_SALE = ['в руб. и коп.\n\n(стр. 160)', '20%\n\n(стр. 170)', '20%\n\n(стр. 200)']

def extract_quarter(filename: str) -> str:
    quarter = re.search(r'\d квартал \d{4}', filename)
    if quarter:
        return quarter.group(0)

def excel(file_path: Path) -> pd.DataFrame:
    '''Функция возвращает dataframe со столбцом квартала'''
    file_df = pd.read_excel(file_path, header=14)
    quarter = extract_quarter(file_path.name)
    file_df['Квартал'] = quarter
    return file_df

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    '''Функция чистит лишние столбцы и строки'''
    cols = [col for col in df.columns if df[col].nunique(dropna=False) <= 1]
    df = df.drop(columns=cols)
    for name in df.keys():
        if 'Код вида операции' in name:
            df = df[df[name] != 21]
    return df

def format_ws(ws: Worksheet) -> None:
    '''Процедура форматирует таблицу (расширяет колонки для читаемости)'''
    for col in ws.columns:
        top_cell = col[0]
        col_letter = top_cell.column_letter
        max_len = max([len(str(cell.value or '')) for cell in col[1:]])
        top_cell.value = top_cell.value.replace('\n\n', ' ')
        if len(top_cell.value) - max_len > 20:
            max_len = len(top_cell.value)
        ws.column_dimensions[col_letter].width = max_len + 3
        top_cell.alignment = Alignment(horizontal='center', wrap_text=True)

def main():
    purchase_files = []
    sale_files = []

    if not FOLDER_PATH.exists():
        print(f"Ошибка: Папка {FOLDER_PATH} не найдена. Пожалуйста, создайте её.")
        return

    for file in FOLDER_PATH.glob('*.xlsx'):
        if 'Раздел-8' in str(file):
            purchase_files.append(excel(file))
        elif 'Раздел-9' in str(file):
            sale_files.append(excel(file))

    purchase_df = pd.concat(purchase_files, ignore_index=True).dropna(axis='columns', how='all') if purchase_files else pd.DataFrame()
    sale_df = pd.concat(sale_files, ignore_index=True).dropna(axis='columns', how='all') if sale_files else pd.DataFrame()

    pivot_purchases = pd.DataFrame()
    if not purchase_df.empty and COL_VALUES_PURCHASE in purchase_df.columns:
        pivot_purchases = purchase_df.pivot_table(
            values=COL_VALUES_PURCHASE, index=COL_INDEX, aggfunc='sum'
        ).reset_index()

    pivot_sales = pd.DataFrame()
    if not sale_df.empty:
        # Проверяем, все ли колонки для свода продаж присутствуют в файле
        available_sale_vals = [col for col in COL_VALUES_SALE if col in sale_df.columns]
        if available_sale_vals:
            pivot_sales = sale_df.pivot_table(
                values=available_sale_vals, index=COL_INDEX, aggfunc='sum'
            ).reset_index()
    relations = {'Покупки': clean_df(purchase_df) if not purchase_df.empty else purchase_df,
                'Продажи': clean_df(sale_df) if not sale_df.empty else sale_df,
                'Свод покупок': pivot_purchases,
                'Свод продаж': pivot_sales}

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for sheet_name, table in relations.items():
            if table.empty:
                continue
                
            table.to_excel(writer, sheet_name=sheet_name, index=False)
            format_ws(writer.sheets[sheet_name])

if __name__ == '__main__':
    main()