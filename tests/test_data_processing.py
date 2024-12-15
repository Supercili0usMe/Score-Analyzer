import unittest
import os
import openpyxl as xl
from src.data_processing import *
from unittest.mock import patch, MagicMock

class TestDataProcessing(unittest.TestCase):
    def setUp(self):
        self.base_folder = "data"
        self.file_name = "Отметки_1.xlsx"
        self.project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        self.folder_root = os.path.join(self.project_root, self.base_folder)
        self.valid_file_folder = os.path.join(self.folder_root, self.file_name)
        self.invalid_file_folder = os.path.join(self.folder_root, "нету.xlsx")

    @patch("os.path.exists")
    def test_get_file_path_valid(self, mock_exists):
        '''Тестируем успешное получение пути к файлу.'''
        mock_exists.return_value = True
        result = get_file_path(self.file_name, self.base_folder)
        self.assertEqual(result, self.valid_file_folder)
    
    @patch("os.path.exists")
    def test_get_file_path_file_not_found(self, mock_exists):
        '''Тестируем поведение при отсутствии файла.'''
        mock_exists.return_value = False
        with self.assertRaises(FileNotFoundError):
            get_file_path(self.file_name, self.base_folder)

    @patch("openpyxl.load_workbook")
    def test_read_excel_valid(self, mock_read_excel):
        '''Тестируем успешное чтение файла.'''
        mock_read_excel.return_value = xl.Workbook()
        result = read_excel(self.valid_file_folder)
        self.assertIsInstance(result, xl.workbook.workbook.Workbook)

    @patch("openpyxl.load_workbook")
    def test_read_excel_error(self, mock_read_excel):
        '''Тестируем обработку ошибок при чтении файла.'''
        mock_read_excel.side_effect = Exception("Файл поврежден")
        with self.assertRaises(ValueError):
            read_excel(self.invalid_file_folder)
    
class TestExtractInfo(unittest.TestCase):
    def setUp(self):
        self.workbook = xl.Workbook()
        self.sheet = self.workbook.active

    def test_basic_case(self):
        '''Тестируем корректное извлечение информации.'''
        self.sheet["A1"] = "Организация:"
        self.sheet["A2"] = "Хогвардс"
        self.sheet["A3"] = "Обучающийся:"
        self.sheet["A4"] = "Гарри Владимирович Поттер"
        self.sheet["A5"] = "Класс:"
        self.sheet["A6"] = "Выпускной"
        self.sheet["A7"] = "Период:"
        self.sheet["A8"] = "Расцвет римской империи"

        expected = {"Организация": "Хогвардс", "Обучающийся": "Гарри Владимирович Поттер",
                    "Класс": "Выпускной", "Период": "Расцвет римской империи"}
        
        result = extract_info(self.workbook)
        self.assertEqual(result, expected)
    
    def test_empty_cells(self):
        '''Тестируем случай, когда ячейки пусты.'''
        self.sheet["A1"] = "Организация:"
        self.sheet["A2"] = "Хогвардс"
        self.sheet["A3"] = "Обучающийся:"
        self.sheet["A4"] = "Гарри Владимирович Поттер"
        self.sheet["A5"] = "Класс:"
        self.sheet["A6"] = None
        self.sheet["A7"] = None
        self.sheet["A8"] = None

        expected = {"Организация": "Хогвардс", "Обучающийся": "Гарри Владимирович Поттер",
                    "Класс": None}
        result = extract_info(self.workbook)
        self.assertEqual(result, expected)

    def test_invalid_cells(self):
        '''Тестируем случай, когда ключи или значения отсутствуют'''
        self.sheet["A1"] = None
        self.sheet["A2"] = "Хогвардс"
        self.sheet["A3"] = "Обучающийся:"
        self.sheet["A4"] = "Гарри Владимирович Поттер"
        self.sheet["A5"] = None
        self.sheet["A6"] = None
        self.sheet["A7"] = None
        self.sheet["A8"] = None

        expected = {"Обучающийся": "Гарри Владимирович Поттер"}
        result = extract_info(self.workbook)
        self.assertEqual(result, expected)


if __name__ == "__main__":
    unittest.main()