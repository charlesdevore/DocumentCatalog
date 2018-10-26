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
import win32com.client

class CatalogProperties(object):
    
    def __init__(self, args=None):

        # Default parameters
        # search_dir is the directory to search for new files in.
        self.search_dirs = [os.getcwd()]

        # base_dir is the base pivot directory to use for determining
        # subdirectories from.
        self.base_dir = os.getcwd()

        # link is a flag to indicate whether symlinks should be added
        # to the files.
        self.link = False
        
        # link_dir is the location for links used in the Excel
        # spreadsheet. These links are sym links that point to the
        # exiting file location.
        self.link_dir = os.path.join(os.getcwd(), '_Links')

        # exclude_dirs is a list of relative directory names that
        # should be excluded from each subdirectory.
        self.exclude_dirs = ['_Links']
        
        # Default parameters for the checksum calculation
        self.hash_function = hashlib.sha1()
        self.buffer_size = 65536

        # Use input args to set Catalog parameters
        if args:
            self.set_input_args(args)

    def set_input_args(self, args):
        pass


class FileCatalog(object):

    def __init__(self, catalog_properties):

        self.catalog_properties = catalog_properties

        self._files = []
        self.load_files()

    def load_files(self):

        self.load_existing_catalog()
        self.search_for_new_files()

        # Add computed properties to files
        self.add_links()
        self.add_checksum()
        self.check_duplicates()

    def add_file(self, file_path):
        
        file_obj = File(file_path)

        if file_obj not in self._files:
            self._files.append(file_obj)
            
    def load_existing_catalog(self):
        pass

    def search_for_new_files(self):

        for search_dir in self.catalog_properties.search_dirs:
            for root, dirs, files in os.walk(search_dir):
                for f in files:
                    file_path = os.path.join(root, f)
                    self.add_file(file_path)

                for exclude_dir in self.catalog_properties.exclude_dirs:
                    if exclude_dir in dirs:
                        dirs.remove(exclude_dir)

    def add_links(self):

        if self.catalog_properties.link:

            link_dir = self.catalog_properties.link_dir

            for file_obj in self._files:
                file_obj.add_link(link_dir)

    def add_checksum(self):

        hash_function = self.catalog_properties.hash_function
        buffer_size = self.catalog_properties.buffer_size

        for file_obj in self._files:
            file_obj.set_checksum(hash_function, buffer_size)

    def check_duplicates(self):
        hash_map = {}

        hash_function = self.catalog_properties.hash_function
        buffer_size = self.catalog_properties.buffer_size

        for file_obj in self._files:
            chksum = file_obj.get_checksum(hash_function, buffer_size)

            if chksum in hash_map:
                file_obj.duplicate = True

            else:
                file_obj.duplicate = False
                hash_map[chksum] = True
                

    def get_base_dir(self):
        if self.catalog_properties.base_dir:
            return self.catalog_properties.base_dir
        else:
            paths = [f.path for f in self._files]
            return os.path.commonpath(paths)

    def as_df(self):

        base_dir = self.get_base_dir()

        files = [f.as_dict(base_dir) for f in self._files]

        df = pd.DataFrame(files)
        
        ordered_cols = self.ordered_columns(df.columns)
        
        return df[ordered_cols]

    def ordered_columns(self, columns):

        ordered_cols = []

        if 'File Path' in columns:
            ordered_cols.append('File Path')
        else:
            raise InputError

        sub_dir_cols = [c for c in columns if 'Subdirectory' in c]
        sub_dir_cols.sort()
        ordered_cols += sub_dir_cols

        goal_cols = ['Filename', 'File Size', 'Readable Size',
                     'Checksum', 'Duplicate', 'File Link', 'Directory']

        for gc in goal_cols:
            if gc in columns:
                ordered_cols.append(gc)

        remaining_cols = [c for c in columns if c not in ordered_cols]

        ordered_cols += remaining_cols

        return ordered_cols
        

class File(object):
    def __init__(self, path):

        # Check path exists
        if not os.path.isfile(path):
            raise InputError

        # Assign constructor input parameters
        self.path = path

        # Set basic properties
        self.name = self.find_file_name()
        self.extension = self.find_extension()
        self.size = self.find_file_size()

        self.hash_function = None
        self.buffer_size = None
        self.checksum = None
        self.duplicate = False
        self.dir_link = None
        self.link_dir = None
        self.link_path = None
        self.link = None

    def __str__(self):
        
        return str(self.__dict__)

    def __eq__(self, other):
        # Test to see if the two paths the same. First check using the
        # faster string comparison of the lower case path and if true
        # then check using the slower os based method.
        if self.path.lower() is other.path.lower():
            return os.path.samefile(self.path, other.path)
        else:
            return False

    def as_dict(self, base_dir=None):

        file_dict = {'File Path': self.path,
                     'Filename': self.name,
                     'File Size': self.size,
                     'Readable Size': get_human_readable(self.size),
                     'Checksum': self.checksum,
                     'Duplicate': self.duplicate,
                     'Directory': self.dir_link,
                     'File Link': self.link}
        
        if base_dir:
            sub_dirs = self.find_sub_dirs(base_dir)

        else:
            return file_dict

        for ii, sd in enumerate(sub_dirs):
            file_dict['Subdirectory {}'.format(ii+1)] = sd

        return file_dict

    def find_sub_dirs(self, base_dir):
        """For a given base directory, find the relative path and return as
        list of individual directories.
        """        
        rel_path = os.path.relpath(self.path, base_dir)

        return os.path.normpath(rel_path).split(os.path.sep)[:-1]
        
    def find_file_name(self):

        return os.path.split(self.path)[1]

    def find_extension(self):

        return os.path.splitext(self.name)[1]

    def find_file_size(self):

        return os.path.getsize(self.path)

    def set_checksum(self, hash_function, buffer_size):

        if any([hash_function is not self.hash_function,
                not self.checksum]):

            self.hash_function = hash_function
            self.buffer_size = buffer_size
            self.checksum = compute_checksum_for_file(self.path,
                                                      self.hash_function,
                                                      self.buffer_size)

    def get_checksum(self, hash_function, buffer_size):

        self.set_checksum(hash_function, buffer_size)
        
        return self.checksum

    def add_link(self, link_dir):

        if not self.link or self.link_dir is link_dir:

            self.link_dir = link_dir
            self.link_name = self.find_link_name() + self.extension

            long_name = long_file_name(self.path)

            os.link(long_name, self.link_path())

            self.link = '=hyperlink("{}","File")'.format(self.link_path)

    def link_path(self):
        return os.path.join(self.link_dir, self.link_name)

    def find_link_name(self):
        h = hashlib.new(hashlib.sha1().name)
        h.update(self.path.encode())
        return h.hexdigest()

def search_in_new_directory(search_dir, exclusion_dirs=['_Links'],
                            verbose_flag=False, check_existing_file_paths=True):

    # Search in a new directory
    files_list = find_files(search_dir, exclusion_dirs=exclusion_dirs,
                            verbose_flag=verbose_flag,
                            check_existing_file_paths=check_existing_file_paths)
    files_list, max_depth = subdirectory(files_list, search_dir)
    files_list = find_duplicates(files_list)
    files_df = file_catalog(files_list, max_depth)

    return files_df
    

def search_in_directory_with_existing_catalog(search_dir, input_file,
                                              exclusion_dirs=['_Links'], verbose_flag=False,
                                              check_existing_file_paths=True):

    # Search in directory with an existing catalog
    existing_df = load_existing(input_file)
    existing_cols = list(existing_df)
    existing_list = [row['File Path']
                     for ii, row in existing_df.iterrows()]
    files_list = find_files(search_dir, exclusion_dirs=exclusion_dirs,
                            existing_files=existing_list,
                            verbose_flag=verbose_flag,
                            check_existing_file_paths=check_existing_file_paths)
    files_list, max_depth = subdirectory(files_list, search_dir)
    files_list = find_duplicates(files_list)
    new_df = file_catalog(files_list, max_depth)
    files_df = existing_df.append(new_df, ignore_index=True)
    ordered_cols = existing_cols + list(set(list(files_df)) - set(existing_cols))
    files_df = files_df[ordered_cols]
    
    return files_df
    
    
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

    
if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Process arguments for DocumentCatalog')
    parser.add_argument('-s', '--search-dir', type=str)
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

    args = parser.parse_args()

    # Set directories to exclude
    if args.exclude_directories is None:
        exclusion_dirs = ['_Links']
    else:
        exclusion_dirs = args.exclude_directories


    if args.copy:

        if args.copy_dir is not None and args.search_dir is not None:

            # Copy files from search directory to copy directory
            copy_files(args.search_dir,
                       args.copy_dir,
                       allow_dest_exist=args.allow_overwrite)


        elif args.copy_key is not None and args.output_copy_dir is not None:

            # Copy specific files to output copy directory
            pass


        else:

            print("""Error: Copy requested but cannot complete
            due to improper specifications.""")
        

    else:

        if args.search_dir is not None:

            if args.input_file is None:

                files_df = search_in_new_directory(args.search_dir,
                                                   exclusion_dirs=exclusion_dirs,
                                                   verbose_flag=args.verbose,
                                                   check_existing_file_paths=not args.do_not_check_existing_file_paths)
            else:

                files_df = search_in_directory_with_existing_catalog(args.search_dir,
                                                                     args.input_file,
                                                                     exclusion_dirs=exclusion_dirs,
                                                                     verbose_flag=args.verbose,
                                                                     check_existing_file_paths=not args.do_not_check_existing_file_paths)

        if args.create_links:

            if args.link_dir is None:
                if args.search_dir is not None:
                    link_dir = os.path.join(args.search_dir, '_Links')
                    files_df = link(files_df, link_dir, verbose_flag=args.verbose)

                else:
                    print('Error: Link directory and search directory not specified.')

            else:
                link_dir = args.link_dir
                files_df = link(files_df, link_dir, verbose_flag=args.verbose)


            if args.create_OSX_links:

                # Add OSX links
                pass


        if args.output:

            fname = args.output_file
            export(files_df, fname, sheet_name='Files', allow_overwrite=args.allow_overwrite)

            
    # print(files_df)

