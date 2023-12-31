'''
Takes IP addresses and device IDs from a spreadsheet and exports that info into
 a Windows-readable hosts file.
Hosts file info: https://en.wikipedia.org/wiki/Hosts_(file)
Looks for worksheets named "ARCH_LTG IP", "PROD_LTG_IP", and "ARCH_CTRL_IP"
Only rows with both DEVICE ID and IP ADDRESS will be written to hosts file. If 
either of those columns are mising data, that row will be ignored.
'''
# pandas is a python data analysis and manipulation library - https://pandas.pydata.org/

__author__ = "John Musarra"
__license__ = "MIT"
__email__ = "john@mightymu.net"
__maintainer__ = "John Musarra"
__version__ = "testing"
import sys
import os
import time

from datetime import datetime
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas

HOSTS_FILE_SOURCE = 'C:\\Windows\\System32\\drivers\\etc\\hosts'
hosts_file_backup_location = f'{Path.home()}\\hosts-backup'
# Open a filepicker dialog:
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
ip_doc_source = Path(askopenfilename(
	                                 title = "IP Document location",
	                                 filetypes=[("Excel files", ".xlsx .xls")])
                                     )
if ip_doc_source == Path('.'):
    sys.exit('No source document slected. Exiting')
#escape the quotes in the filename:
ip_doc_source = f'{ip_doc_source}"'
if os.path.isfile(ip_doc_source):
    print('yes, that seems like a valid file')
    print(f'Using source file: {str(ip_doc_source)}')

# this is becoming a horrible mess ☹️ and I'm making it worse and worse ☹️☹️☹️
HOSTS_FILE_BACKUP_LOCATION = 'C:\\Windows\\System32\\drivers\\etc\\'
HOSTS_FILE_BACKUP_PATH = f'{HOSTS_FILE_BACKUP_LOCATION}hosts_backup-{datetime.now().strftime("%Y%m%d")}'

#Try to backup to System32\etc\drivers
def copy_with_powershell(source, destination):
    '''
    copy source to destination using Powershell
    '''
    print(f'Copying {source} to {destination}...')
    exit_code = os.system(f'powershell.exe copy {source} {destination}')
    if exit_code == 1:
        print('the copy did not worked')
    elif exit_code == 0:
        print('the copy did worked')
    return exit_code

result = copy_with_powershell(HOSTS_FILE_SOURCE, HOSTS_FILE_BACKUP_PATH)
if result == 0:
    print('looks like it worked?')
elif result == 1:
    question = input(''' This script is running without administrator privileges.
 Therefore it cannot create a new file inside the Windows\\System32\\drivers\\etc 
directory. You can proceed by overwriting the existing hosts file, or create the 
backup somewhere we have write access to.
    To overwrite any existing hosts file, type 'o'
    To make a backup in C:\\Users\\<you>, type 'h'
    To make a backup C:\\Windows\\System32\\drivers\\etc, type 'q' to quit, and re-run as Administrator.
    h, o, q? \n''')
    if question == 'h':
    	# write backup to homedir
        HOSTS_FILE_BACKUP_LOCATION = f'{Path.home()}\\'
        HOSTS_FILE_BACKUP_PATH = f'{HOSTS_FILE_BACKUP_LOCATION}hosts_backup-{datetime.now().strftime("%Y%m%d")}'
        copy_with_powershell(HOSTS_FILE_SOURCE, HOSTS_FILE_BACKUP_PATH)
    elif question == 'o':
    	# do not backup, overwrite original
        print("You have chosen to overwrite! Very well, it's your data.")
    else:
        print('Exiting. To run as administrator, right-click and select "Run as Administrator"')
        time.sleep(3)
        sys.exit()

if os.path.isfile(HOSTS_FILE_BACKUP_PATH):
    print(f'Existing hosts file backed up at {HOSTS_FILE_BACKUP_PATH}')

# make a local copy of the document so we're not trying to work with an active document:
IP_DOC_TEMP = str(Path.home()) + '\\ip_doc_temp.xlsx'
# just in case there's an old temp file hanging out, remove it:
if os.path.isfile(IP_DOC_TEMP):
    print('old temp file found, removing')
    os.remove(IP_DOC_TEMP)
elif os.path.isdir(IP_DOC_TEMP):
    print("old temp file found, except it's a directory. Removing.")
    os.rmdir(IP_DOC_TEMP)

command_escaped = f'powershell.exe copy \\"{ip_doc_source}\\" "{IP_DOC_TEMP}"'
try:
    #powershell copy \"C:\Users\jmusarra\ip_doc.xlsx\" "C:\Users\jmusarra\ip_doc_temp.xlsx"
    os.system(command_escaped)
    if os.path.isfile(IP_DOC_TEMP):
        print(command_escaped)
        print('Great, copy sucesful. Moving on.')
    else:
        print('Copy failed somehow wtf')
except PermissionError:
    print("oh goddamnit")
if os.path.isfile(IP_DOC_TEMP):
    print(f'IP doc temp location: {IP_DOC_TEMP}')
else:
    print('hm the copy failed...')

# windows boilerplate:
FAFF = '''
# Copyright (c) 1993-2009 Microsoft Corp.
#
# This is a sample HOSTS file used by Microsoft TCP/IP for Windows.
#
# This file contains the mappings of IP addresses to host names. Each
# entry should be kept on an individual line. The IP address should
# be placed in the first column followed by the corresponding host name.
# The IP address and the host name should be separated by at least one
# space.
#
# Additionally, comments (such as these) may be inserted on individual
# lines or following the machine name denoted by a '#' symbol.
#
# For example:
#
#      102.54.94.97     rhino.acme.com          # source server
#       38.25.63.10     x.acme.com              # x client host

# localhost name resolution is handled within DNS itself.
#	127.0.0.1       localhost
#	::1             localhost

# Generated hosts content follows:

'''

# make a local copy of the document :
#destination_file = "C:\\Windows\\System32\\drivers\\etc\\hosts"
if os.path.isfile(IP_DOC_TEMP):
    with pandas.ExcelFile(IP_DOC_TEMP) as file:
        print(f'Document found: {IP_DOC_TEMP}.')
        # check if our target sheets (ARCH_LTG IP, PROD_LTG IP, ARCH_CTRL IP) exist:
        sheets = file.sheet_names
        print(sheets)
        # make Pandas dataframes from selected worksheets:
        print('Reticulating splines...')
        available_sheets = []
        if 'ARCH_LTG IP' in sheets:
            print('found ARCH_LTG IP')
            arch_ltg_ip = (
            	           pandas.read_excel(IP_DOC_TEMP,
            	           sheet_name = "ARCH_LTG IP",
            	           header = 4,
            	           index_col = None,
            	           usecols = ['DEVICE ID', 'IP ADDRESS'])
            	           )
            available_sheets.append(arch_ltg_ip)
        else:
            print('nope')
        if 'PROD_LTG IP' in sheets:
            print('found PROD_LTG IP')
            prod_ltg_ip = (
            	           pandas.read_excel(IP_DOC_TEMP,
            	           sheet_name = "PROD_LTG IP",
            	           header = 4,
            	           index_col = None,
            	           usecols = ['DEVICE ID', 'IP ADDRESS'])
            	           )
            available_sheets.append(prod_ltg_ip)
        else:
            print('nope')
        if 'ARCH_CTRL IP' in sheets:
            print('found ARCH_CTRL IP')
            arch_ctrl_ip = (
            	            pandas.read_excel(IP_DOC_TEMP,
            	            sheet_name = "ARCH_CTRL IP",
            	            header = 4,
            	            index_col = None,
            	            usecols = ['DEVICE ID', 'IP ADDRESS'])
            	            )
            available_sheets.append(arch_ctrl_ip)
        else:
            print('nope')
        # merge each of the individual dataframes into a complete dataframe:
        merged_frames = pandas.concat(available_sheets)
        print('Done.')
else:
    sys.exit("Could not find IP document. Exiting.")

# rearrange columns as needed for hosts file order:
print("Reordering columns...")
merged_frames = merged_frames[['IP ADDRESS', 'DEVICE ID']]
print('Done.')

# strip whitespace:
merged_frames['DEVICE ID'] = merged_frames['DEVICE ID'].str.replace(' ', '')
merged_frames['IP ADDRESS'] = merged_frames['IP ADDRESS'].str.replace(' ', '')

# drop all rows that are missing either hostname or IP:
print('Dropping incomplete rows...')
merged_frames.dropna(how = 'any', inplace = True)
print('Done')
num_devices = f'# Number of devices: {merged_frames.shape[0]}\n\n'

with pandas.option_context('display.max_rows', None):
	# formatters make columns be left-justified
    formatters = {'IP ADDRESS': (lambda x: '{:<14}'.format(x)),
                  'DEVICE ID': (lambda x: '{:<32}'.format(x))}
    hosts = merged_frames.to_string(
    	                            index = False,
    	                            header = False,
    	                            formatters = dict(formatters)
    	                            )

generated_date = f'# Date generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}\n'

#overwriting existing hosts file; this should only happen if user selected 'o'
print('Writing file....')
with open(HOSTS_FILE_SOURCE, 'w', encoding='cp1252') as f:
    f.write(FAFF + generated_date + num_devices + hosts)

print(f"Hosts file generation complete. Written to {HOSTS_FILE_SOURCE}.")
print("Removing temporary files\n")

# clean up temporary file 😬
os.remove(IP_DOC_TEMP)
time.sleep(6)
