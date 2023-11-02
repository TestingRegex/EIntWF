import os
import shutil
from oletools.olevba3 import VBA_Parser
import subprocess

EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)
KEEP_NAME = False  # Set this to True if you would like to keep "Attribute VB_Name"

#A function to obtain each vba-module as its own file
#Input: the location of the workbooks containing macros
#Output: A file contained in a directory by the same name as the workbook that the underlying module comes from
def parse(workbook_path):
    vba_path = workbook_path + '_vba'
    if not os.path.exists(vba_path):
         os.makedirs(vba_path)

    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, filename, content in vba_modules:
        lines = []
        if '\r\n' in content:
            lines = content.split('\r\n')
        else:
            lines = content.split('\n')
        if lines:
            content = []
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    with open(os.path.join(vba_path, filename), 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content))

#A function to obtain a list containing the names of the modified files in the git repository
def get_modified_files():
    git_status_output = subprocess.check_output(["git","status","--porcelain"]).decode("utf-8")
    modified_files = [line[3:] for line in git_status_output.splitlines() if line.startswith("M")]
    for f in modified_files:
       if f[0] == '"':
           modified_files[modified_files.index(f)] = f[1:-1]

    return modified_files

'''
if __name__ == '__main__':
    for root, dirs, files in os.walk('.'):
#        for f in dirs:
#            if f.endswith('_vba'):
#                shutil.rmtree(os.path.join(root, f))

        for f in files:
            if f.endswith(EXCEL_FILE_EXTENSIONS) and f in get_modified_files():
                parse(os.path.join(root, f))

'''
'''
#Alternative Main:
if __name__ == '__main__':
   root = os.getcwd()
   modified_files = get_modified_files()
   for file in modified_files:
      if file.endswith(EXCEL_FILE_EXTENSIONS):
         parse(file)
'''
