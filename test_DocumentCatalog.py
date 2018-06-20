import unittest

import DocumentCatalog as DC
import os

test_dir = os.path.join(os.getcwd(), 'test')
                        

class TestDC(unittest.TestCase):

    def test_find_files(self):
        all_files =  DC.find_files(test_dir)
        self.assertTrue(len(all_files) == 6)
        
        some_files =  DC.find_files(test_dir, exclusion_dirs=['sub_dir'])
        self.assertTrue(len(some_files) == 5)

        existing_files = [sf['File Path'] for sf in some_files]
        diff_files = DC.find_files(test_dir, existing_files=existing_files)
        self.assertTrue(len(diff_files) == 1)


if __name__ == '__main__':
    unittest.main()
