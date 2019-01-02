import unittest

import DocumentCatalog as DC
import os
import hashlib

test_dir = os.path.join(os.getcwd(), 'test')
CP = DC.CatalogProperties()
CP.search_dirs = [test_dir]

class TestDC(unittest.TestCase):

    def test_search_in_new_directory(self):
        FC = DC.FileCatalog(CP)
        self.assertEqual(len(FC), 9)

    def test_exclude_directories(self):
        CP2 = DC.CatalogProperties()
        CP2.search_dirs = [test_dir]
        CP2.exclude_dirs = ['sub_dir']
        FC2 = DC.FileCatalog(CP2)
        self.assertEqual(len(FC2), 6)

    def test_search_in_directory_with_existing_catalog(self):
        input_file = os.path.join(test_dir, 'some_files.xlsx')
        CP3 = DC.CatalogProperties()
        CP3.search_dirs = [test_dir]
        CP3.existing_catalog = input_file
        FC3 = DC.FileCatalog(CP3)
        self.assertEqual(len(FC3), 9)


    def test_duplicate_detection(self):
        FC = DC.FileCatalog(CP)
        df = FC.as_df()
        self.assertEqual(len(df.loc[df['Duplicate']==False]), 5)


    def test_checksum(self):
        h = hashlib.sha1()
        buffer_size = 4096
        fp1 = os.path.join(test_dir, 'email02.msg')
        chksum1 = DC.compute_checksum_for_file(fp1, h, buffer_size)
        fp2 = os.path.join(test_dir, 'sub_dir', 'email02-renamed.msg')
        chksum2 = DC.compute_checksum_for_file(fp2, h, buffer_size)
        self.assertEqual(chksum1, chksum2)

if __name__ == '__main__':
    unittest.main()
