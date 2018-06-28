""" 
DocumentCatalog

Module used to catalog documents and their metadata in an Excel
spreadsheet. Includes a shortened hard link to facilitate linking to
documents.

Charles DeVore
June 2018
"""

import os
import argparse
import pandas as pd
import platform
import datetime
import win32com.client


def load_existing(fname):

    df = pd.read_excel(fname)

    return df


def add_files_to_existing(df, search_dir):

    files = find_new_files_from_existing(df, search_dir)

    files, max_depth = subdirectory(files, search_dir)

    N = df.iloc[-1].name + 1
    files = link(files, search_dir, link_start=N)

    files = add_hyperlinks(files)

    FC = file_catalog(files, max_depth)

    ndf = pd.concat([df, FC], ignore_index=True)

    # Organize columns
    cols = df.columns.to_list()
    ndf = ndf[cols]

    return ndf


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
                     with allow_existing_directory flag to enable 
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


def find_files(search_dir, existing_files=[''],
               exclusion_dirs=['_Links'], verbose_flag=False):

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

    file_list = []
    counter = 0

    for root, dirs, files in os.walk(search_dir):
        for f in files:
            file_path = os.path.join(root, f)

            # existing_files should be a list of absolute paths
            if not file_path in existing_files:

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

                file_list.append(file)
                counter += 1

        # Exclude directories based on the exclusion_dirs list.
        for exclude_dir in exclusion_dirs:
            if exclude_dir in dirs:
                dirs.remove(exclude_dir)

    # Output file statistics
    if verbose_flag:
        print('--------------')
        print('\t==> Found ' + str(counter) + ' files in: ' + search_dir + '\n')

    return file_list


def link(files_df, base_dir, verbose_flag=False):

    """
    Create hard links in link directory corresponding to a unique
    hexadecimal representation of the file iterator number.
    """

    link_dir = os.path.join(base_dir, '_Links')
    if not os.path.isdir(link_dir):
        os.mkdir(link_dir)


    link_counter = 0

    for file in files_df.iterrows():

        link_fname = 'file_{0:x}.{1:s}'.format(file.name, file['Extension'])
        link_path = os.path.join(link_dir, link_fname)

        long_name = long_file_name(file['File Path'])

        try:
            os.link(long_name, link_path)
            link_counter += 1

        except:
            print('Error making link:\n{}\n{}'.format(long_name, link_path))

        # Save the link paths and add hyperlinks for Excel
        file['Link Path'] = link_path

        file['File Link'] = '=hyperlink("{}","File")'.format(
            file['Link Path'])

        file['Directory Link'] = '=hyperlink("{}","Directory")'.format(
            file['Directory'])

    if verbose_flag:
        print('{} links out of {} added in:\n\t{}\n'.format(link_counter, len(files_df), link_dir))

    return


def file_catalog(file_list, max_depth):

    # DC.file_catalog() builds a DataFrame catalog corresponding to the
    # file information.

    keys = []
    for i in range(1, max_depth+1):
        keys.append('Sub-Directory {}'.format(i))

    new_file_list = []

    for file in file_list:

        try:
            for i in range(0, len(file['Sub-Directories'])):

                file[keys[i]] = file['Sub-Directories'][i]

            del(file['Sub-Directories'])

            new_file_list.append(file)

        except:
            print(file['File Path'])

    files_df = pd.DataFrame(new_file_list)

    # # Order columns
    # cols = ['File Path'] + keys + ['Filename', 'Extension',
    #                                'File Size', 'Link Path',
    #                                'Directory', 'File Link',
    #                                'Directory Link']

    # file_catalog = file_catalog[cols]

    return files_df


def subdirectory(file_list, root_dir):

    # Compute the individual sub-directories based on the root
    # directory. Store the sub-directories in file_list as list in the
    # file's dictionary. Also output the maximum sub-directory depth.

    max_depth = 0
    new_file_list = []

    for file in file_list:

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

            new_file_list.append(file)

            max_depth = max([max_depth, len(file['Sub-Directories'])])

        except:
            print(file['File Path'])

    return new_file_list, max_depth



def export(file_catalog, email_catalog, fname):

    # Export the file catalog to an Excel workbook

    try:
        writer = pd.ExcelWriter(fname)
        file_catalog.to_excel(writer, 'Files')
        email_catalog.to_excel(writer, 'Emails')
        writer.save()

        return 1

    except:
        return -1


def long_file_name(fname):

    # Create the Windows long file name representation for local and
    # network locations.

    if(fname.lower().startswith('c:')):

        long_name = r'\\?\{}'.format(fname)

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
    parser.add_argument('--link-dir', type=str)
    parser.add_argument('-l', '--create-links', action='store_true', default=False)
    parser.add_argument('--create-OSX-links', action='store_true', default=False)
    parser.add_argument('-v', '--verbose', action='store_true', default=False)

    args = parser.parse_args()

    if args.copy:

        if args.copy_dir is not None and args.search_dir is not None:

            # Copy files from search directory to copy directory

        elif args.copy_key is not None and args.output_copy_dir is not None:

            # Copy specific files to output copy directory


        else:

            print("""Error: Copy requested but cannot complete
            due to improper specifications.""")
            return
            
    
    if args.search_dir is not None:

        if args.input_dir is None:

            # Search in a new directory
            
            
        else:

            # Search in with an existing catalog

            
    if args.create_links:

        # Add links
        

        if args.create_OSX_links:

            # Add OSX links
            


    if args.output:

        if args.input_dir is None:

            # Build output with a new file

        else:

            # Build output with existing file

            
