import unittest
from MetadataWriter import MetadataWriter
import openpyxl

class MetadataWriterTest(unittest.TestCase):
    def setUp(self):
        self.writer = MetadataWriter()
        return 0

    def test_spreadsheet_exists(self):
        writer = self.writer.f
        self.assertIsInstance(writer, openpyxl.Workbook)

    def test_spreadsheet_parse(self):
        self.writer.parse_workbook()
        self.assertEqual([], [1])


if __name__ == '__main__':
    unittest.main()
