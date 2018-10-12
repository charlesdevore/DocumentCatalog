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
    
    
def load_existing(fname):

    df = pd.read_excel(fname)

    return df


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


def find_files(search_dir, existing_files=[],
               exclusion_dirs=['_Links'], verbose_flag=False,
               check_existing_file_paths=True):

    """
    find_files(search_dir)
    
    This function walks through the files in search_dir and creates
    a catalog of the files including their directory, filename,
    extension, and file path. The information is stored in a
    list of dictionaries including the Directory, Filename, File Path,
    Extension, File Size, File Size bytes.

    This function takes optional input for excluding existing files,
    excluding directories, and to show verbose output.
    """

    # Remove any files that don't exist from existing files list
    if check_existing_file_paths:
        existing_files = [ef for ef in existing_files if os.path.isfile(ef)]
        
    if verbose_flag:
        print('Searching with {} existing files.'.format(len(existing_files)))
        sys.stdout.flush()
        
    files_list = []
    counter = 0

    for root, dirs, files in os.walk(search_dir):
        for f in files:
            file_path = os.path.join(root, f)

            # existing_files is a list of absolute paths. Use a string
            # comparison to quickly check the lower case for matches
            # and if a match is found, use the slower
            # os.path.samefile() function
            if (not file_path.lower() in [ef.lower() for ef in
                                          existing_files] or not
                any([os.path.samefile(file_path,
                                      ef) for ef in existing_files
                     if ef.lower() is
                     file_path.lower()])):
                
                # The file extension is the second entry of the split list
                ext = os.path.splitext(f)[1]

                file = {}

                file['Directory'] = root
                file['Filename']  = f
                file['Extension'] = ext
                file['File Path'] = file_path

                file['File Size'] = get_human_readable(
                    os.path.getsize(file_path), 0)

                file['File Size (bytes)'] = os.path.getsize(file_path)

                files_list.append(file)
                counter += 1

        # Exclude directories based on the exclusion_dirs list.
        for exclude_dir in exclusion_dirs:
            if exclude_dir in dirs:
                dirs.remove(exclude_dir)

    # Output file statistics
    if verbose_flag:
        print('--------------')
        print('\t==> Found {} new files in: {} \n'.format(counter, search_dir))
        
    return files_list


def link(files_df, link_dir, verbose_flag=False, allow_overwrite=False):

    """
    Create hard links in link directory corresponding to a unique
    hexadecimal representation of the file iterator number.
    """

    if not os.path.isdir(link_dir):
        os.mkdir(link_dir)

    # Check length of link_dir to ensure that links will be under
    # Excel limit of 256 characters. Assume max link value of 12.
    if len(link_dir) > (256-12):
        print('The link directory is {} characters long and may result in hyperlinks not working. Please find a  new link directory with a shorter path.'.format(len(link_dir)))
        user_continue = raw_input('Continue? [Y/n]')
        if not (lower(user_continue) == 'y' or lower(user_continue) == None):
            return files_df
        
    # Add columns for the link path, file link, and directory link if
    # not already existing in the dataframe
    new_cols = ['Link Path', 'File Link', 'Directory Link']
    for col in new_cols:
        if col not in files_df.columns:
            files_df[col] = ''

    link_counter = 0

    for ii,row in files_df.iterrows():

        # The link file name starts with the letter 'f' and then is the
        # DataFrame index integer expressed as hexadecimal with its
        # extension (Windows freaks out if extension is omitted).
        
        link_fname = 'f_{0:x}{1:s}'.format(row.name, row['Extension'])
        link_path = os.path.join(link_dir, link_fname)

        long_name = long_file_name(row['File Path'])


        if os.path.isfile(link_path) and not allow_overwrite:
            if verbose_flag:
                print('Warning: Link already exists.\n\t{}\n\t{}'.format(long_name, link_path))
            continue
        
        try:
            os.link(long_name, link_path)
            link_counter += 1

        except:
            print('Error: Cannot make link.\n\t{}\n\t{}'.format(long_name, link_path))
            continue

        # Save the link paths and add hyperlinks for Excel
        files_df.loc[ii, 'Link Path'] = link_path

        files_df.loc[ii, 'File Link'] = '=hyperlink("{}","File")'.format(link_path)

        dir_path = row['Directory']
        files_df.loc[ii, 'Directory Link'] = '=hyperlink("{}","Directory")'.format(dir_path)
        

    if verbose_flag:
        print('{} links out of {} added in:\n\t{}\n'.format(link_counter, len(files_df), link_dir))

    return files_df


def file_catalog(files_list, max_depth):

    # DC.file_catalog() builds a DataFrame catalog corresponding to the
    # file information.

    keys = []
    for i in range(1, max_depth+1):
        keys.append('Sub-Directory {}'.format(i))

    new_files_list = []

    for file in files_list:

        try:
            for i in range(0, len(file['Sub-Directories'])):

                file[keys[i]] = file['Sub-Directories'][i]

            del(file['Sub-Directories'])

            new_files_list.append(file)

        except:
            print(file['File Path'])

    files_df = pd.DataFrame(new_files_list)

    files_df = order_file_columns(files_df)
    
    return files_df


def find_duplicates(files_list, hash_function=hashlib.sha1(), buffer_size=65536):

    # Find duplicate files by reading each file anc computing a
    # hash. The default hash function is SHA1.

    file_hash_map = {}
    new_files_list = []
    
    for file in files_list:
        checksum = compute_checksum_for_file(file['File Path'],
                                             hash_function, buffer_size)
        
        file['Checksum'] = checksum

        file['Duplicate'] = True if checksum in file_hash_map else False

        file_hash_map[checksum] = file['File Path']

        new_files_list.append(file)

    return new_files_list


def compute_checksum_for_file(file_path, hash_function, buffer_size):

    h = hashlib.new(hash_function.name)
    
    with open(file_path, 'rb') as f:
        data = f.read(buffer_size)
        while data:
            h.update(data)
            data = f.read(buffer_size)

    return h.hexdigest()

def subdirectory(files_list, root_dir):

    # Compute the individual sub-directories based on the root
    # directory. Store the sub-directories in files_list as list in the
    # file's dictionary. Also output the maximum sub-directory depth.

    max_depth = 0
    new_files_list = []

    for file in files_list:

        try:
            rel_path = os.path.relpath(file['File Path'], root_dir)

            # Find the individual sub-directories by spliting the
            # relative path using os.path.split().
            head, fname = os.path.split(rel_path)

            # Test the extracted filename to make sure it matches the
            # filename in the dictionary.
            if not fname == file['Filename']:
                raise NameError('Filename does not match extracted value.')

            subdirs = []
            while len(head) > 0:

                head, tail = os.path.split(head)
                subdirs.append(tail)

            # Reverse the order of subdirs to achieve expected order
            # of decreasing sub-directory depth.
            subdirs = subdirs[::-1]

            file['Sub-Directories'] = subdirs

            new_files_list.append(file)

            max_depth = max([max_depth, len(file['Sub-Directories'])])

        except:
            print(file['File Path'])

    return new_files_list, max_depth



def export(catalog_df, fname, sheet_name='Files', allow_overwrite=False):

    """Export the catalog to an Excel workbook. Take as input either a
    file catalog or an email catalog and save to a specific sheet in the
    workbook."""

    try:
        if (allow_overwrite and os.path.isfile(fname)) or not os.path.isfile(fname):
            writer = pd.ExcelWriter(fname)
            catalog_df.to_excel(writer, sheet_name)
            writer.save()

        else:
            print('Error: Output file already exists. Enable overwrite or choose a new file name.')
            # new_fname = raw_input('New file name: ')
            # export(catalog_df, new_fname, sheet_name=sheet_name, allow_overwrite=allow_overwrite)

    except Exception as err:
        print('Error exporting catalog\nFile Name: {}\nSheet Name: {}'.format(fname, sheet_name))
        print(err)


def order_file_columns(files_df):

    cols = list(files_df)

    sub_dir_cols = [c for c in cols if c.startswith('Sub-Directory')]

    sub_dir_cols.sort()

    ordered_cols =  ['File Path'] \
                    + sub_dir_cols \
                    + ['Filename', 'Extension', 'File Size']

    # Add any remaing columns
    ordered_cols += set(cols) - set(ordered_cols)

    return files_df[ordered_cols]
    

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

