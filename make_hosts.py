'''
Takes IP addresses and device IDs from a spreadsheet and exports that info into a Windows-readable hosts file
Hosts file info: https://en.wikipedia.org/wiki/Hosts_(file)
Only hosts with both DEVICE ID and IP ADDRESS will be written to hosts file. If either of those columns are 
 mising data, that row will be ignored.
'''
# pandas is a python data analysis and manipulation library - https://pandas.pydata.org/
import pandas, sys, os, time
from datetime import datetime
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Open a filepicker dialog:
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
ip_doc_source = Path(askopenfilename(title = "IP Document location", filetypes=[("Excel files", ".xlsx .xls")]))
print(str(ip_doc_source) + '\\')

# make a local copy of the document so we don't error 😬
ip_doc_temp = str(Path.home()) + '\\ip_doc_temp.xlsx'
print(f'IP doc temp location: {ip_doc_temp}')
try:
	command = f'powershell.exe copy \\"{ip_doc_source}\\" "{ip_doc_temp}"'
	print(command)
	os.system(command)
except PermissionError:
	print("oh goddamnit")

# windows boilerplate:
faff = '''
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
destination_file = "C:\\Windows\\System32\\drivers\\etc\\hosts"
exists = os.path.isfile(ip_doc_temp)
if exists:
	with pandas.ExcelFile(ip_doc_temp):
	    print(f'Document found: {ip_doc_temp}.')
	    # make Pandas dataframes from selected worksheets:
	    print('Reticulating splines...')
	    arch_ltg_ip = pandas.read_excel(ip_doc_temp, sheet_name = "ARCH_LTG IP", header = 4, index_col = None, usecols = ['DEVICE ID', 'IP ADDRESS'])
	    prod_ltg_ip = pandas.read_excel(ip_doc_temp, sheet_name = "PROD_LTG IP", header = 4, index_col = None, usecols = ['DEVICE ID', 'IP ADDRESS'])
	    prod_ltg_ip = pandas.read_excel(ip_doc_temp, sheet_name = "ARCH_CTRL IP", header = 4, index_col = None, usecols = ['DEVICE ID', 'IP ADDRESS'])
	    print('Done.')
else:
	sys.exit("Could not find IP document. Exiting.")

# rearrange columns as needed for hosts file order:
print("Reordering columns...")
arch_ltg_ip = arch_ltg_ip[['IP ADDRESS', 'DEVICE ID']]
print('Done.')

# strip whitespace:
arch_ltg_ip['DEVICE ID'] = arch_ltg_ip['DEVICE ID'].str.replace(' ', '')
arch_ltg_ip['IP ADDRESS'] = arch_ltg_ip['IP ADDRESS'].str.replace(' ', '')

# drop all rows without hostname:
print('Dropping incomplete rows...')
arch_ltg_ip.dropna(how = 'any', inplace = True)
print('Done')

with pandas.option_context('display.max_rows', None):
	hosts = arch_ltg_ip.to_string(index = False, header = False)
generated_date = f'# Date generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}\n'

print('Writing file....')
with open(destination_file, 'w') as f:
	f.write(faff + generated_date + hosts)

print(f"Hosts file generation complete. Written to {destination_file}.")
print("Removing temporary files")

# clean up temporary file 😬
os.remove(ip_doc_temp)
time.sleep(10)
