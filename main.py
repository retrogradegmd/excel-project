import pandas as pd
import re
import time
from pathlib import Path
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment
from openpyxl import load_workbook
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

BASE_DIR = Path(__file__).resolve().parent
FOLDER_PATH = BASE_DIR / "Gmail"  
OUTPUT_FILE = BASE_DIR / "report.xlsx"
TEMPLATE_FILE = BASE_DIR / "template.xlsx"

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

def main():
    template_book = load_workbook(TEMPLATE_FILE)
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

    with pd.ExcelWriter(TEMPLATE_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        writer.workbook = template_book
        
        if not purchase_df.empty:
            purchase_df.to_excel(writer, sheet_name='Покупки', index=False)
        if not sale_df.empty:
            sale_df.to_excel(writer, sheet_name='Продажи', index=False)

    template_book.save(OUTPUT_FILE)

class ExcelFolderHandler(FileSystemEventHandler):
    """Класс, который ждет изменений в папке"""
    def __init__(self):
        super().__init__()
        self.last_run = 0

    def on_created(self, event):    
        '''Метод, срабатывающий при появлении файла excel'''
        if not event.is_directory and event.src_path.endswith('.xlsx') and not '~$' in event.src_path:
            print('Заметил файл')
            current_time = time.time()
            if current_time - self.last_run > 2:

                time.sleep(0.5)

                try:
                    main()
                    self.last_run = time.time()
                except Exception as e:
                    print(f'{e}')

if __name__ == '__main__':
    main()

    event_handler = ExcelFolderHandler()
    observer = Observer()
    observer.schedule(event_handler, path = str(FOLDER_PATH), recursive=False)
    observer.start()
    print('Слежу за папкой')

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
