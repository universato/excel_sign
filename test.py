import unittest
from pywinauto.keyboard import send_keys
import sign_excel_winauto
from unittest import TestCase


class TestSignExcelWinauto(unittest.TestCase):
    # def setUp(self):
        # print("setup")
        # self.eq = QuadraticEquation.QuadraticEquation()

    def test_find_excel_exe(self):
        """test method of find_excel_exe"""
        unexpected = r"C:\Program Files\WRONG_PATH\EXCEL.EXE"
        sign_excel_winauto.EXCEL_APP_PATH = unexpected
        sign_excel_winauto.find_excel_exe()

        actual = sign_excel_winauto.EXCEL_APP_PATH
        expected_regex = r"C:\\Program Files( \(x86\))?\\Microsoft Office\\(root\\)?Office\d\d\\EXCEL.EXE"

        self.assertNotEqual(unexpected, actual)
        self.assertRegex(actual, expected_regex)

        expected = actual
        sign_excel_winauto.find_excel_exe()
        actual = sign_excel_winauto.EXCEL_APP_PATH
        self.assertEqual(expected, actual)
        self.assertRegex(actual, expected_regex)

    def test_open_excel_and_open_signature_dialog(self):
        """ test method of open_excel_and_open_signature_dialog() """

        excel_file_path = r'C:\Users\unive\dev\vba_labo\a.xlsm'
        excel_app = sign_excel_winauto.open_excel_and_open_signature_dialog(excel_file_path)
        # expected = (4.0, 9.0)
        # actual = (self.eq.calc_value(1.0), self.eq.calc_value(2.0))

        self.assertTrue(excel_app[u"デジタル署名"].exists())

        # excel_app.kill()



    # def tearDown(self):
        # print("tearDown")
        # del self.eq


if __name__ == '__main__':
    unittest.main()
