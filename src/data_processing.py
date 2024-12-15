import os
from openpyxl import load_workbook
import openpyxl

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
folder_root = os.path.join(project_root, 'data')


def get_file_path(file_name, base_folder):
    '''Возвращает полный путь к файлу, если он существует в указанной папке'''
    file_path = os.path.join(folder_root, file_name)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f'Файл "{file_name}" не найден в папке "{base_folder}"')
    return file_path

def read_excel(file_path):
    '''Читает Excel-файл и возвращает его содержимое как `openpyxl.Workbook`'''
    try:
        data = load_workbook(file_path)
        return data
    except Exception as e:
        raise ValueError(f"Ошибка при чтении файла: {e}")
    
def extract_info(data, start_row=1, step=2, max_rows=7):
    '''Выделяем основную информацию об ученике и возвращаем список'''
    ws = data.active
    info = {}
    for row in range(start_row, max_rows + start_row, step):
        key_cell = ws[f"A{row}"]
        value_cell = ws[f"A{row+1}"]

        # Извлекаем данные из ячеек
        key = key_cell.value[:-1] if key_cell.value and isinstance(key_cell.value, str) else None
        value = value_cell.value if value_cell.value is not None else None

        if key:
            info[key] = value
    return info

def main():
    # Проверяем существование папки
    if not os.path.exists(folder_root):
        print(f"Папка 'data' не найдена по пути: {folder_root}")
        return

    # Запрос имени файла у пользователя
    file_name = input("Введите имя файла (например, example.xlsx): ").strip()

    try:
        file_path = get_file_path(file_name, folder_root)
        data = read_excel(file_path)
        print("Данные успешно загружены!")
        return data
    except (FileNotFoundError, ValueError) as e:
        print(e)

if __name__ == "__main__":
    data = main()

