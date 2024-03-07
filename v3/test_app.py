import unittest
from unittest.mock import patch
from app import app, save_excel_on_server, delete_file

class TestApp(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()

    @patch('app.sts.split_workbook')
    def test_split_workbook_route(self, mock_split_workbook):
        filename = 'example.xlsx'
        response = self.app.get(f'/split_workbook/{filename}')
        
        mock_split_workbook.assert_called_once_with(filename)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers['Content-Disposition'], f'attachment; filename={filename}')

    def test_save_excel_on_server(self):
        workbook = 'example.xlsx'
        base_dir = '/path/to/save'
        regnum = '12345'
        
        # Call the function
        save_excel_on_server(workbook, base_dir, regnum)
        
        # Add your assertions here to verify the expected behavior

    def test_delete_file(self):
        directory = '/path/to/files'
        age_days = 14
        
        # Call the function
        delete_file(directory, age_days)
        
        # Add your assertions here to verify the expected behavior

if __name__ == '__main__':
    unittest.main()