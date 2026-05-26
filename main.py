import pandas as pd
import os
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment

folder_path = r"C:/Users/apostol/Desktop/excel-project/Gmail" 
files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

def excel(name: str) -> pd.DataFrame:
    '''Функция возвращает dataframe со столбцом квартала'''
    index = name.find('квартал')
    quarter_data = name[index - 2: index + 12]
    file_df = pd.read_excel(f'{folder_path}/{name}', header = 14)
    quarter_df = pd.DataFrame({
        'Квартал': [quarter_data] * len(file_df)})
    result = pd.concat([file_df, quarter_df], axis=1)
    return result

def dropping_columns(df: pd.DataFrame) -> pd.DataFrame:
    '''Процедура чистит лишние столбцы и строки'''
    for i, name in enumerate(df.keys()):
        if len(set(df[name])) == 1:
            df = df.drop(columns=name)
        if i == 3:
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

purchase_files_df = pd.concat([excel(file) for file in files if 'Раздел-8' in file]).dropna(axis='columns', how='all')
sale_files_df = pd.concat([excel(file) for file in files if 'Раздел-9' in file]).dropna(axis='columns', how='all')

pivot_purchases_df = pd.pivot_table(purchase_files_df,
                                    values='Стоимость покупок с НДС\n\n(стр. 170)',
                                    index='Наименование',
                                    aggfunc='sum').reset_index()
pivot_sales_df = pd.pivot_table(sale_files_df,
                                     values=['в руб. и коп.\n\n(стр. 160)', '20%\n\n(стр. 170)', '20%\n\n(стр. 200)'],
                                     index='Наименование',
                                     aggfunc='sum').reset_index()

relations = {'Покупки': purchase_files_df,
             'Продажи': sale_files_df,
             'Свод покупок': pivot_purchases_df,
             'Свод продаж': pivot_sales_df}

file_name = 'report.xlsx'

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    for name, table in relations.items():
        if 'Свод' in name:
            table.to_excel(writer, sheet_name=name, index=False)
        else:
            dropping_columns(table).to_excel(writer, sheet_name=name, index=False)
        format_ws(writer.sheets[name])
    