""" 
DocumentCatalog

Module used to catalog documents and their metadata in an Excel
spreadsheet. Includes a shortened hard link to facilitate linking to
documents.

Charles DeVore
June 2018
"""

import os
import sys
import argparse
import pandas as pd
import platform
import hashlib
import datetime
import xlsxwriter
import win32com.client
import sqlite3

class CatalogProperties(object):
    """CatalogProperties provides an interface for FileCatalog.

This class manages the properties for constructing a catalog and
provides an interface between argparse command line arguments and the
required catalog properties. Manages default behavior and provides
input error checking.

Args:
    args (:obj:): A parser object that contains the argument values 
        from the command line using argparse.

Attributes:
    search_dirs (list of strings): List of search directory 
        paths of where to search for files.  
    existing_catalog (str): Filename or path to filename containing a
        spreadsheet with an existing document catalog.
    database (str): Filename of the SQLite3 database to store
        intermediate results.
    output_file (str): Filename or path where the output spreadsheet
        should be saved.
    base_dir (str): The base directory that should be used as the 
        pivot to compute the subdirectory columns.
    link (bool): Whether hard links should be constructed.
    link_dir (str): Location where hard links should be saved.
    exclude_dirs (list of strings): List of directories to exclude 
        from file search. Stored as relative paths and used to 
        exclude through entire search path.
    hash_function (:obj: hashlib.function): The hash function used
        to compute the checksum to differentiate unique files.
    buffer_size (int): The number of bytes to use as a buffer when
        reading the file for computation of the checksum.
    verbose (bool): Flag for verbose output.
    """

    def __init__(self, args=None):

        self.search_dirs = [os.getcwd()]
        self.existing_catalog = None

        self.database = 'document_catalog.db'

        # The number of rows to buffer before inserting into the database
        self.database_row_buffer = 100

        # output_file is the name of the output filename or path of
        # the output Excel spreadsheet. The file extension must be
        # xlsx. If output_file is None, then no output is generated.
        self.output_file = None

        self.base_dir = None

        self.link = False
        self.link_dir = os.path.join(os.getcwd(), '_Links')
        self.exclude_dirs = ['_Links']

        self.hash_function = hashlib.sha1()
        self.buffer_size = 65536

        self.verbose = False

        # Use input args to set Catalog parameters
        if args:
            self.set_input_args(args)

    def set_input_args(self, args):
        
        # Set directories to exclude
        if args.exclude_directories:
            self.exclude_dirs = args.exclude_directories

        if args.search_dir:
            self.search_dirs = [args.search_dir]

        if args.input_file:
            if os.path.isfile(args.input_file):
                self.existing_catalog = os.path.realpath(args.input_file)

            else:
                print('Error with input file.')
                raise InputError
            
        if args.copy:
            self.copy = args.copy
            self.copy_dir = args.copy_dir
            self.copy_key = args.copy_key

        if args.create_links:
            self.link = True
            if args.link_dir:
                self.link_dir = os.path.realpath(args.link_dir)
            else:
                self.link_dir = os.path.realpath(os.path.join(os.curdir, '_Links'))
            
        if args.create_OSX_links:
            pass

        if args.database:
            self.database = args.database

        if args.output:
            file_out = args.output_file
            if os.path.splitext(file_out)[1] == '.xlsx':
                self.output_file = os.path.realpath(file_out)

            else:
                print('Error with output file, extension not .xlsx')
                raise InputError

        if args.verbose:
            self.verbose = True
            

    def as_dict(self):

        return {'Search Directories': self.search_dirs,
                'Base Directory': self.base_dir,
                'Link Directory': self.link_dir,
                'Database': self.database}
    


class FileCatalog(object):
    """FileCatalog organizes File objects.

    FileCatalog provides a collection of File objects corresponding to a
    particular search operation. The parameters of the search are
    specified by the CatalogProperties object that is taken as input.

    Args:
        catalog_properties (:CatalogProperties:): Parameters for 
            constructing FileCatalog.

    Attributes:
        catalog_properties (:CatalogProperties:): Same as input.
        files (list[:File:]): A list of file objects corresponding
            to each file contained within the catalog. 

    """
    def __init__(self, catalog_properties):

        self.catalog_properties = catalog_properties

        self.files = []
        self.load_files()
        self.export()

    def __len__(self):
        return len(self.files)

    def load_files(self):

        self.create_database()

        self.load_existing_catalog()
        N_existing_files = len(self.files)
        if self.catalog_properties.verbose:
            print('Existing Files Loaded: {}'.format(N_existing_files))
            
        self.search_for_new_files()
        N_new_files = len(self.files) - N_existing_files
        if self.catalog_properties.verbose:
            print('New Files Loaded: {}'.format(N_new_files))

        self.insert_to_database()
        
        # Add computed properties to files
        self.add_links()
        # self.add_checksum()
        self.check_duplicates()

    def add_file(self, file_obj):
        
        if file_obj not in self.files:
            self.files.append(file_obj)

        if len(self.files) % self.catalog_properties.database_row_buffer == 0:
            self.insert_to_database(row_buffer = self.catalog_properties.database_row_buffer)

    def insert_to_database(self, row_buffer=None):

        if not row_buffer:
            row_buffer = len(self.files) % self.catalog_properties.database_row_buffer

        self.cursor.executemany('INSERT INTO files VALUES (?,?,?,?,?,?)', [f.as_tuple() for f in self.files[-row_buffer:]])

        self.connection.commit()
            
    def load_existing_catalog(self):

        df = self.import_existing_catalog()
        if df.empty:
            return
        
        CP = self.import_existing_properties()

        for index,row in df.iterrows():
            info = dict(row)
            path = row['File Path']
            file_obj = ExistingFile(path, info=info, CP=CP)
            self.add_file(file_obj)


    def import_existing_catalog(self):

        existing_filename = self.catalog_properties.existing_catalog

        if existing_filename:
            df = pd.read_excel(existing_filename, sheet_name='Catalog')

        else:
            df = pd.DataFrame()

        return df

    def import_existing_properties(self):

        return CatalogProperties()

    def search_for_new_files(self):

        if self.catalog_properties.verbose:
            print('Searching...')
        for search_dir in self.catalog_properties.search_dirs:
            for root, dirs, files in os.walk(search_dir):
                for f in files:
                    file_path = os.path.join(root, f)
                    file_obj = File(file_path, search_dir, self.catalog_properties)
                    self.add_file(file_obj)

                for exclude_dir in self.catalog_properties.exclude_dirs:
                    if exclude_dir in dirs:
                        dirs.remove(exclude_dir)

                if self.catalog_properties.verbose:
                    print(root)

    def create_database(self):

        if os.path.isfile(self.catalog_properties.database):
            usr_response = input('Warning: {} already exists, continue writing to database? [y/N]'.format(self.catalog_properties.database))
            if not usr_response.lower() == 'y':
                self.catalog_properties.database = input('Please enter new database name: ')
                self.create_database()

            self.connection = sqlite3.connect(self.catalog_properties.database)
            self.cursor = self.connection.cursor()

        else:
            self.connection = sqlite3.connect(self.catalog_properties.database)
            self.cursor = self.connection.cursor()
            self.cursor.execute('''CREATE TABLE files
            (rel_path text,
            filename text,
            extension text,
            size integer,
            human_readable text,
            checksum text)''')
            self.connection.commit()


    def add_links(self):

        if self.catalog_properties.link:

            link_dir = self.catalog_properties.link_dir

            if not os.path.isdir(link_dir):
                os.mkdir(link_dir)

            # Check length of link_dir to ensure that links will be under
            # Excel limit of 256 characters. Assume max link value of 45.
            if len(link_dir) > (256-45):
                print("""The link directory is {} characters long and
                may result in hyperlinks not working. Please find a
                new link directory with a shorter
                path.""".format(len(link_dir)))
                user_continue = input('Continue? [Y/n]')
                if not (lower(user_continue) == 'y' or not user_continue):
                    self.link = False
                    return

            for file_obj in self.files:
                file_obj.add_link(link_dir)

                
    def add_checksum(self):

        hash_function = self.catalog_properties.hash_function
        buffer_size = self.catalog_properties.buffer_size

        for file_obj in self.files:
            file_obj.set_checksum(hash_function, buffer_size)

    def check_duplicates(self):
        hash_map = {}

        for file_obj in self.files:
            if file_obj.checksum in hash_map:
                file_obj.duplicate = True

            else:
                file_obj.duplicate = False
                hash_map[file_obj.checksum] = True
                

    def get_base_dir(self):

        if not self.catalog_properties.base_dir:
            paths = [f.path for f in self.files]
            self.catalog_properties.base_dir = os.path.commonpath(paths)

        return self.catalog_properties.base_dir

    def export(self):
        if self.catalog_properties.output_file:

            if os.path.isfile(self.catalog_properties.output_file):
                allow_overwrite = input('Output file exists. Allow overwrite? [Y/n]\n')

                if allow_overwrite.lower() == 'y' or not allow_overwrite:
                    self.to_excel()

                else:
                    output_file = input('Please enter the output filename.\n')
                    self.catalog_properties.output_file = output_file
                    self.export()

            else:
                self.to_excel()
                

    def to_excel(self):

        writer = pd.ExcelWriter(self.catalog_properties.output_file,
                                engine='xlsxwriter')
        
        # Export files information to Worksheet named "Catalog"
        df = self.as_df()        
        df.to_excel(writer, sheet_name='Catalog')

        # Export catalog_properties to Worksheet named "Properties" by
        # using the xlsxwriter workbook object
        workbook = writer.book
        self.properties_to_excel(workbook)

        writer.save()


    def properties_to_excel(self, workbook):
        
        worksheet = workbook.add_worksheet('Properties')

        row, col = 0,0

        # Header
        header_str = 'Document Catalog Properties'
        worksheet.write(row, col, header_str)
        row += 2

        # Search Directories
        worksheet.write(row, col, 'Search Directories:')
        for sd in self.catalog_properties.search_dirs:
            worksheet.write(row, col+1, sd)
            row += 1

        # Exclude Directories
        worksheet.write(row, col, 'Exclude Directories:')
        for ed in self.catalog_properties.exclude_dirs:
            worksheet.write(row, col+1, ed)
            row += 1

        # Existing Catalog
        if self.catalog_properties.existing_catalog:
            worksheet.write(row, col, 'Existing Catalog:')
            worksheet.write(row, col+1, self.catalog_properties.existing_catalog)
            row += 1

        # Buffer Size
        worksheet.write(row, col, 'Buffer Size:')
        worksheet.write(row, col+1, self.catalog_properties.buffer_size)
        row +=1

        # Hash Function
        worksheet.write(row, col, 'Hash Function:')
        worksheet.write(row, col+1, self.catalog_properties.hash_function.name)
            
                        

    def as_df(self):

        base_dir = self.get_base_dir()

        files = [f.as_dict(base_dir) for f in self.files]

        df = pd.DataFrame(files)
        
        ordered_cols = self.ordered_columns(df.columns)
        
        return df[ordered_cols]

    def ordered_columns(self, columns):

        ordered_cols = []

        if 'File Path' in columns:
            ordered_cols.append('File Path')
        else:
            raise InputError

        if 'Base Directory' in columns:
            ordered_cols.append('Base Directory')

        if 'Relative Path' in columns:
            ordered_cols.append('Relative Path')

        sub_dir_cols = [c for c in columns if 'Subdirectory' in c]
        sub_dir_cols.sort()
        ordered_cols += sub_dir_cols

        goal_cols = ['Filename', 'Extension', 'File Size', 'Readable Size',
                     'Checksum', 'Duplicate', 'File Link', 'Directory']

        for gc in goal_cols:
            if gc in columns:
                ordered_cols.append(gc)

        remaining_cols = [c for c in columns if c not in ordered_cols]

        ordered_cols += remaining_cols

        return ordered_cols
        

class File(object):
    """File finds and stores file metadata.

    File is an object that finds and stores the metadata found during
    the search operation. File is instantiated by FileCatalog during
    its directory walk.

    Args:
        path (str): A path to the file. If the file does not exist,
            an InputError is thrown.

    Attributes:
        path (str): A file path to the file in question.
        name (str): The filename with extension.
        extension (str): The file extension.
        size (int): The file size in bytes.
        hash_function (:hashlib:): The hashlib function used to compute 
            the checksum.
        buffer_size (int): The buffer size used when reading the file
            during checksum computation.
        duplicate (bool): Whether a file is a duplicate based on the 
            checksum.
        dir_link (str): An Excel function hyperlink to the directory.
        link_dir (str): The directory that the link is saved in.
        link (str): An Excel function hyperlink to the file link.

    """

    def __init__(self, path, base_dir, catalog_properties):

        # Check path exists
        if not os.path.isfile(path):
            raise InputError

        # Assign constructor input parameters
        self.path = path
        self.base_dir = base_dir
        self.catalog_properties = catalog_properties

        # Set basic properties
        self.name = self.find_file_name()
        self.extension = self.find_extension()

        self._size = None
        self._checksum = None
        self.duplicate = False
        self.dir_link = None
        self.link_dir = None
        self.link = None
        self.link_name = None

    def __str__(self):
        
        return str(self.__dict__)

    def __eq__(self, other):
        # Test to see if the two paths the same. First check using the
        # faster string comparison of the lower case path and if true
        # then check using the slower os based method.
        if self.path.lower() == other.path.lower():
            return os.path.samefile(self.path, other.path)
        else:
            return False

    def as_dict(self, base_dir=None):

        file_dict = {'File Path': self.path,
                     'Filename': self.name,
                     'Extension': self.extension,
                     'File Size': self.size,
                     'Readable Size': self.human_readable,
                     'Checksum': self.checksum,
                     'Duplicate': self.duplicate,
                     'Directory': self.dir_link,
                     'File Link': self.link,
                     'Link Directory': self.link_dir,
                     'Link Name': self.link_name}
        
        if base_dir:
            sub_dirs = self.find_sub_dirs(base_dir)

        else:
            return file_dict

        file_dict['Base Directory'] = base_dir
        file_dict['Relative Path'] = self.relative_path
        for ii, sd in enumerate(sub_dirs):
            file_dict['Subdirectory {}'.format(ii+1)] = sd

        return file_dict

    def as_tuple(self):

        return (self.relative_path, self.name, self.extension, self.size, self.human_readable, self.checksum)

    def find_sub_dirs(self, base_dir):
        """For a given base directory, find the relative path and return as
        list of individual directories.
        """        
        return os.path.normpath(self.relative_path).split(os.path.sep)[:-1]

    @property
    def human_readable(self):
        return get_human_readable(self.size)

    @property
    def relative_path(self):
        return os.path.relpath(self.path, self.base_dir)

    @property
    def checksum(self):
        if not self._checksum:
            self._checksum = self.find_checksum()

        return self._checksum

    @property
    def size(self):
        if not self._size:
            self._size = self.find_file_size()

        return self._size

    
    def find_file_name(self):

        return os.path.split(self.path)[1]

    def find_extension(self):

        return os.path.splitext(self.name)[1]

    def find_file_size(self):

        return os.path.getsize(self.path)

    def find_checksum(self):

        return compute_checksum_for_file(self.path,
                                         self.catalog_properties.hash_function,
                                         self.catalog_properties.buffer_size)

    def add_link(self, link_dir):

        if not self.link or self.link_dir is link_dir:

            self.link_dir = link_dir
            self.link_name = self.find_link_name() + self.extension

            long_name = long_file_name(self.path)

            os.link(long_name, self.link_path())

            self.link = '=hyperlink("{}","File")'.format(self.link_path())

            self.dir_link = '=hyperlink("{}","Link")'.format(self.directory_path())

    def directory_path(self):
        return os.path.split(self.path)[0]
            
    def link_path(self):
        return os.path.join(self.link_dir, self.link_name)

    def find_link_name(self):
        h = hashlib.new(hashlib.sha1().name)
        h.update(self.path.encode())
        return h.hexdigest()

class ExistingFile(File):

    def __init__(self, path, info=None, CP=None):

        super().__init__(path)

        self.input_info = info
        self.input_CP = CP

        self.process_input()

    def __eq__(self, other):
        
        return super().__eq__(other)

    def __str__(self):

        return super().__str__()

    def process_input(self):
        if self.input_info:
            self.process_input_info()

        if self.input_CP:
            self.process_input_CP()

    def process_input_info(self):

        key_attr = {'File Size': 'size',
                    'Checksum':  'checksum',
                    'Duplicate': 'duplicate'}

        for key in self.input_info:
            value = self.input_info[key]
            if key in key_attr:
                attr  = key_attr[key]
                setattr(self, attr, value)
                # print(key, value, getattr(self, attr))

    def process_input_CP(self):

        key_attr = {'hash_function': 'hash_function',
                    'buffer_size':   'buffer_size'}                    
                
        for key in self.input_CP.__dict__:
            value = getattr(self.input_CP, key)
            if key in key_attr:
                attr = key_attr[key]
                setattr(self, attr, value)
                # print(key, value, getattr(self, attr))

    
def copy_files(source_dir, dest_dir, batch_file = 'run_DC_copy.bat', allow_dest_exist=False):

    """
    copy_files(source_dir, dest_dir)

    Use the following windows commands to copy the files and change
    the attributes. Create a batch file and run using Windows
    command.

    robocopy source_dir dest_dir *.* /E /COPY:DT /DCOPY:DAT
    attrib +R dest_dir\* /S
    """

    if not platform.system() == 'Windows':
        raise OSError

    if not allow_dest_exist:
        if os.path.isdir(dest_dir):
            # Destination directory already exists
            print('''Destination directory exists. Rerun 
                     with --allow-overwrite flag to enable 
                     copying. Warning, this may cause overwriting 
                     of existing files.''')
            
            return -1

        else:
            os.mkdir(dest_dir)


    with open(batch_file, 'w') as bfile:

        bfile.write('ECHO OFF\n')
        bfile.write('ROBOCOPY "{}" "{}" *.* /E /COPY:DT /DCOPY:DAT\n'.format(source_dir, dest_dir))
        bfile.write('ATTRIB +R "{}"\\* /S'.format(dest_dir))

    try:
        os.system(batch_file)

    except:
        print('Batch file did not run correctly.')
        return -2

    finally:
        os.remove(batch_file)

    return 1


def compute_checksum_for_file(file_path, hash_function, buffer_size):

    h = hashlib.new(hash_function.name)

    try:
        with open(file_path, 'rb') as f:
            data = f.read(buffer_size)
            while data:
                h.update(data)
                data = f.read(buffer_size)

        return h.hexdigest()

    except PermissionError:
        return None

    return None
    
def long_file_name(fname):

    # Create the Windows long file name representation for local and
    # network locations.

    if(fname.lower().startswith('c:')):

        long_name = r'\\?\{}'.format(os.path.normpath(fname))

    elif(fname.startswith(r'\\')):

        long_name = r'\\?\UNC{}'.format(fname[1:])

    return long_name


def get_human_readable(size, precision=0):

    # Take bytes as input and return human readable string to
    # specified precision.

    suffixes = ['B', 'KB', 'MB', 'GB', 'TB']
    suffixIndex = 0

    while size > 1024 and suffixIndex < 4:
        suffixIndex += 1  # increment the index of the suffix
        size = size/1024.0  # apply the division

    return "%.*f%s"%(int(precision), size, suffixes[suffixIndex])


def email_catalog(emails):


    # Create Outlook object using MAPI
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Initialize data structure as an empty list. Entries will
    # be dictionaries including the appropriate data for each
    # email.
    catalog = []

    for e in emails:

        d = {'Filename': e['Filename'],
             'Link Path': e['Link Path'],
             'Directory': e['Directory'],
             'File Size': e['File Size'],
             'File Path': e['File Path'],
             'File Link': e['File Link'],
             'Directory Link': e['Directory Link']}

        d['error'] = 0

        try:

            msg = outlook.OpenSharedItem(d['Link Path'])

            try:
                d['Subject'] = msg.Subject

                try:
                    d['From']    = msg.SenderName
                    d['To']      = msg.To
                    d['CC']      = msg.CC

                except:
                    d['error'] += 4

                d['Number of Attachments']  = int(msg.Attachments.Count)

                t = msg.SentOn
                d['Sent Date'] = datetime.datetime(t.year,t.month,t.day,t.hour,t.minute,t.second)

            except:
                d['error'] += 2

        except:
            d['error'] += 1

        else:
            del(msg)

        # Append dictionary to catalog list
        catalog.append(d)

    EC = pd.DataFrame(catalog)

    # Order columns
    cols = ['File Path', 'Filename', 'To', 'From', 'CC', 'Sent Date',
            'Subject', 'Number of Attachments', 'File Link', 'Directory Link',
            'Link Path', 'Directory', 'error']

    EC = EC[cols]

    return EC


def copy_specific_files(df, dest_dir, allow_existing_dir=False):

    """
    copy_specific_files(df, dest_dir)

    Use to copy files containted in a dataframe to a destination
    directory. The dataframe must contain a column 'File Path' that
    corresponds to the absolute path location of the file. A batch file is
    created and robocopy is used to copy the files.  
    """

    
    if not platform.system() == 'Windows':
        raise OSError

    if not allow_existing_dir:
        if os.path.isdir(dest_dir):
            # Destination directory already exists
            print('''Destination directory exists. Rerun 
                     with --allow-overwrite flag to enable 
                     copying. Warning, this may cause overwriting 
                     of existing files.''')
            return -1

        else:
            os.mkdir(dest_dir)

    batch_file = 'run.bat'

    with open(batch_file, 'w') as bfile:

        bfile.write('ECHO OFF\n')

        for ii, row in df.iterrows():
            fp = row['File Path']
            lp = row['Link Path']
            
            if os.path.isfile(fp):
                path, fname = os.path.split(fp)

                # Add a unique identifier to the filename to prevent name collisions
                prefix = os.path.splitext(os.path.basename(lp))[0]
                dest_fname = prefix + '--' + fname

                if os.path.isfile(os.path.join(dest_dir, dest_fname)):
                    #print('Destination name collision.\n{}\n'.format(fp))
                    pass

                else:
                    bfile.write('ROBOCOPY "{}" "{}" "{}"\n'.format(path, dest_dir, fname))
                    bfile.write('RENAME "{}" "{}"\n'.format(os.path.join(dest_dir, fname), dest_fname))

            else:
                print('Skipping file, does not exist.\n{}\n'.format(fp))


    try:
        os.system(batch_file)

    except:
        # Batch file did not run correctly.
        return -2

    finally:
        os.remove(batch_file)

    return 1


    
def OSX_links(files):

    out_files = []
    
    for file in files:

        file_cmd = 'link.sh -o "{}"'.format(file['Link Path'])
        file['OSX File Link'] = '=shell("{}")'.format(file_cmd)

        dir_cmd = 'link.sh -l "{}"'.format(file['Directory Path'])
        file['OSX Directory Link'] = '=shell("{}")'.format(dir_cmd)
        
        out_files.append(file)

    return out_files

def parse_arugments():

    parser = argparse.ArgumentParser(description='Process arguments for DocumentCatalog')
    parser.add_argument('-s', '--search-dir', type=str)
    parser.add_argument('-d', '--database', type=str)
    parser.add_argument('-o', '--output', action='store_true', default=False)
    parser.add_argument('--output-file', type=str, default='Document Catalog.xlsx')
    parser.add_argument('-i', '--input-file', type=str)
    parser.add_argument('-c', '--copy', action='store_true', default=False)
    parser.add_argument('--copy-dir', type=str)
    parser.add_argument('--copy-key', type=str)
    parser.add_argument('--output-copy-dir', type=str)
    parser.add_argument('--allow-overwrite', action='store_true', default=False)
    parser.add_argument('--exclude-directories', nargs='+')
    parser.add_argument('--link-dir', type=str)
    parser.add_argument('-l', '--create-links', action='store_true', default=False)
    parser.add_argument('--create-OSX-links', action='store_true', default=False)
    parser.add_argument('-v', '--verbose', action='store_true', default=False)
    parser.add_argument('--do-not-check-existing-file-paths', action='store_true', default=False)

    return parser.parse_args()


def main(args=None):

    CP = CatalogProperties(args)
    FC = FileCatalog(CP)

    
if __name__ == '__main__':

    args = parse_arugments()
    main(args)
            
    # print(files_df)

