import unittest

import DocumentCatalog as DC
import os

test_dir = os.path.join(os.getcwd(), 'test')
                        

class TestDC(unittest.TestCase):

    def test_find_files(self):
        all_files =  DC.find_files(test_dir)
        self.assertTrue(len(all_files) == 7)
        
        some_files =  DC.find_files(test_dir, exclusion_dirs=['sub_dir'])
        self.assertTrue(len(some_files) == 6)

        existing_files = [sf['File Path'] for sf in some_files]
        diff_files = DC.find_files(test_dir, existing_files=existing_files)
        self.assertTrue(len(diff_files) == 1)


    def test_search_in_new_directory(self):
        all_df = DC.search_in_new_directory(test_dir)
        self.assertTrue(len(all_df) == 7)


    def test_search_in_directory_with_existing_catalog(self):

        input_file = os.path.join(test_dir, 'some_files.xlsx')
        some_df = DC.search_in_directory_with_existing_catalog(test_dir,
                                                               input_file)
        self.assertTrue(len(some_df) == 7)
        

if __name__ == '__main__':
    unittest.main()
