""" 
DocumentCatalog

Module used to catalog documents and their metadata in SQLite3
database with an option to export to an Excel spreadsheet. 

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
import random
import string

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

        self.search_dir = os.getcwd()
        self.existing_catalog = None
        self.existing_database = None

        # Session ID is the primary key for the search session that is
        # saved in the database
        self.session_id = None

        self.database = 'document_catalog.db'

        # The number of rows to buffer before inserting into the database
        self.database_row_buffer = 100

        self.exclude_dirs = []
        
        # output_file is the name of the output filename or path of
        # the output Excel spreadsheet. The file extension must be
        # xlsx. If output_file is None, then no output is generated.
        self.output_file = None

        self.base_dir = None

        self.hash_function = hashlib.sha1()
        self.buffer_size = 65536

        self.check_file_contents = True
        
        self.verbose = False

        # Use input args to set Catalog parameters
        if args:
            self.set_input_args(args)

    def set_input_args(self, args):
        
        # Set directories to exclude
        if args.exclude_directories:
            self.exclude_dirs = args.exclude_directories

        if args.search_dir:
            self.search_dir = args.search_dir

        if args.base_dir:
            self.base_dir = args.base_dir
        else:
            self.base_dir = self.search_dir

        if args.session_id:
            self.session_id = args.session_id
        else:
            self.session_id = ''.join([random.choice(string.ascii_lowercase) for ii in range(4)])

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

        if args.existing_database:
            self.existing_database = args.existing_database
            self.database = args.existing_database

        if args.database:
            self.database = args.database

        if args.output:
            file_out = args.output_file
            if os.path.splitext(file_out)[1] == '.xlsx':
                self.output_file = os.path.realpath(file_out)

            else:
                print('Error with output file, extension not .xlsx')
                raise InputError

        if args.do_not_check_file_contents:
            self.check_file_contents = False
            
        if args.verbose:
            self.verbose = True


    def load_existing_catalog(self):
        
        pass

    def as_dict(self):

        return {'Search Directories': self.search_dirs,
                'Base Directory': self.base_dir,
                'Database': self.database,
                'Session ID': self.session_id,
                'Hash Function': self.hash_function.name,
                'Buffer Size': self.buffer_size}

    def insert_to_database(self, cursor):

        cursor.execute('INSERT INTO catalog_properties VALUES (?,?,?,?,?,?)', self.as_tuple())

    def as_tuple(self):

        return (self.session_id, self.search_dir, self.base_dir,
                self.hash_function.name, self.buffer_size,
                datetime.datetime.isoformat(datetime.datetime.utcnow()))


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
    _files_to_database = []
    
    def __init__(self, catalog_properties):

        self.catalog_properties = catalog_properties

        self.files = []
        self.load_files()
        self.export()

    def __len__(self):
        return len(self.files)

    def load_files(self):

        self.create_database()
        self.catalog_properties.insert_to_database(self.cursor)
        self.connection.commit()
        print('Session ID: {}'.format(self.catalog_properties.session_id))

        if self.catalog_properties.existing_catalog:
            self._load_existing_catalog()

        if self.catalog_properties.existing_database:
            self._load_existing_database()
            

        if self.catalog_properties.verbose:
            N_existing_files = len(self.files)
            print('Existing Files Loaded: {}'.format(N_existing_files))
            
        self.search_for_new_files()
        if self.catalog_properties.verbose:
            N_new_files = len(self.files) - N_existing_files
            print('New Files Loaded: {}'.format(N_new_files))

        # Include a final insert_to_database call to add any remaining
        # files in the buffer
        if len(self._files_to_database) > 0:
            self.insert_to_database()
        
        # Compute duplicates
        # self.check_duplicates()

    def add_file(self, file_obj, existing=False):
        
        if file_obj in self.files:
            return
        
        self.files.append(file_obj)

        if self.catalog_properties.verbose and not existing:
            print(file_obj)
            sys.stdout.flush()

        if not existing:
            self._files_to_database.append(file_obj)

        if len(self._files_to_database) == self.catalog_properties.database_row_buffer:
            self.insert_to_database()


    def insert_to_database(self):

        self.cursor.executemany(
            'INSERT INTO files VALUES (?,?,?,?,?,?,?,?)',
            [f.as_tuple() for f in self._files_to_database])

        self.connection.commit()

        # Clear files to database array
        self._files_to_database = []

    def _load_existing_catalog(self):

        df = self.import_existing_catalog()
        if df.empty:
            return
        
        CP = self.import_existing_properties()

        for index,row in df.iterrows():
            info = dict(row)
            path = row['File Path']
            file_obj = ExistingFile(path, info=info, CP=CP)
            self.add_file(file_obj, existing=True)

    def _load_existing_database(self):

        if os.path.isfile(self.catalog_properties.existing_database):
            existing_conn = sqlite3.connect(self.catalog_properties.existing_database)

        else:
            raise InputError('Error loading existing database, file does not exist.\n{}'.format(self.catalog_properties.existing_database))

        existing_cursor = existing_conn.cursor()

        existing_cursor.execute('''
        SELECT base_dir, rel_path, filename, extension, size, checksum, file_key
	FROM files f 
	INNER JOIN catalog_properties cp ON f.session_id = cp.session_id;
        ''')

        rows = existing_cursor.fetchall()

        for row in rows:
            file_obj = DatabaseFile(row, self.catalog_properties)
            self.add_file(file_obj, existing=True)
        
        
    def import_existing_catalog(self):

        existing_filename = self.catalog_properties.existing_catalog

        if existing_filename:
            df = pd.read_excel(existing_filename, sheet_name='Catalog')

        else:
            df = pd.DataFrame()

        return df

    def import_existing_properties(self):
        pass

    def search_for_new_files(self):

        if self.catalog_properties.verbose:
            print('Searching...')

        for root, dirs, files in os.walk(self.catalog_properties.search_dir):
            if self.catalog_properties.verbose:
                print(root)

            for f in files:
                file_path = os.path.join(root, f)
                try:
                    file_obj = File(file_path, self.catalog_properties)
                    self.add_file(file_obj)

                except:
                    print('Error loading {}'.format(file_path))
                    

            for exclude_dir in self.catalog_properties.exclude_dirs:
                if exclude_dir in dirs:
                    dirs.remove(exclude_dir)

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
            self.cursor.execute('''
            CREATE TABLE files
            (rel_path text,
            filename text,
            extension text,
            size integer,
            human_readable text,
            checksum text,
            session_id text,
            file_key text,
            PRIMARY KEY(file_key));
            ''')
            self.cursor.execute('''
            CREATE TABLE catalog_properties
            (session_id text,
            search_dir text,
            base_dir text,
            hash_function text,
            hash_buffer_size integer,
            date text,
            PRIMARY KEY(session_id ASC));
            ''')
            
            self.connection.commit()


    def check_duplicates(self):
        hash_map = {}

        for file_obj in self.files:
            if file_obj.checksum in hash_map:
                file_obj.duplicate = True

            else:
                file_obj.duplicate = False
                hash_map[file_obj.checksum] = True
                

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

        files = [f.as_dict() for f in self.files]

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
                     'Checksum', 'Duplicate']

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

    """

    def __init__(self, path, catalog_properties):

        # Check path exists
        if not os.path.isfile(path):
            raise InputError

        # Assign constructor input parameters
        self.path = path
        self.catalog_properties = catalog_properties

        # Set basic properties
        self.name = self.find_file_name()
        self.extension = self.find_extension()

        self._relative_path = None
        self._size = None
        self._checksum = None
        self._key = None
        self.duplicate = False

    def __str__(self):
        return self.name

    def __eq__(self, other):
        if self.catalog_properties.check_file_contents:
            return self.key == other.key

        return self.relative_path == other.relative_path

    def as_dict(self):
        base_dir = self.catalog_properties.base_dir

        file_dict = {'File Path': self.path,
                     'Filename': self.name,
                     'Extension': self.extension,
                     'File Size': self.size,
                     'Readable Size': self.human_readable,
                     'Checksum': self.checksum,
                     'Duplicate': self.duplicate}
        
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
        return (self.relative_path, self.name, self.extension,
                self.size, self.human_readable, self.checksum,
                self.catalog_properties.session_id, self.key)

    def find_sub_dirs(self):
        """For a given base directory, find the relative path and return as
        list of individual directories.
        """        
        return os.path.normpath(self.relative_path).split(os.path.sep)[:-1]

    @property
    def base_dir(self):
        if self._base_dir:
            return self._base_dir
        
        if self.catalog_properties.base_dir:
            return self.catalog_properties.base_dir

        return None
    
    @property
    def human_readable(self):
        return get_human_readable(self.size)

    @property
    def relative_path(self):
        if not self._relative_path:
            self._relative_path = os.path.relpath(self.path, self.catalog_properties.base_dir)
        return self._relative_path

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

    @property
    def key(self):
        if not self._key:
            self._key = self.find_key()

        return self._key

    
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

    def directory_path(self):
        return os.path.split(self.path)[0]
            

    def find_key(self):
        h = hashlib.new(hashlib.sha1().name)
        h.update(self.path.encode())
        h.update(self.checksum.encode())
        return h.hexdigest()



class DatabaseFile(File):
    def __init__(self, row, catalog_properties):

        self._base_dir = row[0]
        self._relative_path = row[1]
        self.path = os.path.join(row[0], row[1])
        self.name = row[2]
        self.extension = row[3]
        self._size = row[4]
        self._checksum = row[5]
        self._key = row[6]

        self.catalog_properties = catalog_properties


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
    parser.add_argument('-b', '--base-dir', type=str)
    parser.add_argument('-g', '--session-id', type=str)
    parser.add_argument('-d', '--database', type=str)
    parser.add_argument('-e', '--existing-database', type=str)
    parser.add_argument('-o', '--output', action='store_true', default=False)
    parser.add_argument('--output-file', type=str, default='Document Catalog.xlsx')
    parser.add_argument('-i', '--input-file', type=str)
    parser.add_argument('-c', '--copy', action='store_true', default=False)
    parser.add_argument('--copy-dir', type=str)
    parser.add_argument('--copy-key', type=str)
    parser.add_argument('--output-copy-dir', type=str)
    parser.add_argument('--allow-overwrite', action='store_true', default=False)
    parser.add_argument('--exclude-directories', nargs='+')
    parser.add_argument('-v', '--verbose', action='store_true', default=False)
    parser.add_argument('--do-not-check-existing-file-paths', action='store_true', default=False)
    parser.add_argument('--do-not-check-file-contents', action='store_true', default=False)

    return parser.parse_args()


def main(args=None):

    CP = CatalogProperties(args)
    FC = FileCatalog(CP)

    
if __name__ == '__main__':

    args = parse_arugments()
    main(args)
            
    # print(files_df)

