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

    def test_gather_xmp_stats(self):
        fields = self.writer.gather_xmp_stats()
        self.assertEqual(fields,[])

    def test_handle_data(self):
        results = self.writer._handle_data('C:/Users/evharley/source/repos/MetadataWriter/MetadataWriter/2012/2005a0726001.TIF')
        self.assertEqual(results, [])

    def test_write_data(self):
        self.writer._write_data()
        self.fail('Not Implemented yet')

    def test_delete_duplicates(self):
        self.writer.delete_duplicates()
        self.fail('Not Implemented yet')



if __name__ == '__main__':
    unittest.main()
