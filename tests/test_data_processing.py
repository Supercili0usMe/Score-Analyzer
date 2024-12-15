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

if __name__ == "__main__":
    unittest.main()