import unittest

import DocumentCatalog as DC
import os
import hashlib

test_dir = os.path.join(os.getcwd(), 'test')
                        

class TestDC(unittest.TestCase):

    def test_find_files(self):
        all_files =  DC.find_files(test_dir)
        self.assertEqual(len(all_files), 9)
        
        some_files =  DC.find_files(test_dir, exclusion_dirs=['sub_dir'])
        self.assertEqual(len(some_files), 6)

        existing_files = [sf['File Path'] for sf in some_files]
        diff_files = DC.find_files(test_dir, existing_files=existing_files)
        self.assertEqual(len(diff_files), 3)


    def test_search_in_new_directory(self):
        all_df = DC.search_in_new_directory(test_dir)
        self.assertEqual(len(all_df), 9)


    def test_search_in_directory_with_existing_catalog(self):
        input_file = os.path.join(test_dir, 'some_files.xlsx')
        some_df = DC.search_in_directory_with_existing_catalog(test_dir,
                                                               input_file)
        self.assertEqual(len(some_df), 9)


    def test_duplicate_detection(self):
        all_df = DC.search_in_new_directory(test_dir)
        self.assertEqual(len(all_df.loc[all_df['Duplicate']==False]), 5)


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
