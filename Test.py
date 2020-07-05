
def lala(x,y):
    return x+y

# This is for quick rename files in a folder by a excel list in it
print("You should put the excel'New Name List' and the target files and this program in the same folder\n\
    then Press Enter\n\
    (also remind that: the target files in the folder will always be ordered by names)")
input()

# Get the new names from the excel
from openpyxl import load_workbook
wb = load_workbook('New Name List.xlsx')
ws1 = wb["Sheet1"]
new_name_list=[]
for cell in ws1['A']:
    new_name_list.append(cell.value)
wb.close

# Input the target file format
target_format=input("Please input the target file format.xxx,like:txt,jpg >>> ")

# Get the target files in the folder
import os
# Count the target files
target_file_list=[]
for target_file in os.listdir():
    if target_file.endswith(f"{target_format}") and target_file != "New Name List.xlsx"\
        and target_file != "RenameByExcel.exe":
        target_file_list.append(target_file)
if len(target_file_list) != len(new_name_list):
    if input("The target files count differs from the new names count,do you still want to continue?\n\
        Yes(y) or No(n) >>>")=="n":
        import sys
        sys.exit()

# Rename the target files
import shutil
renamed_file_count=0
for i in range(min(len(target_file_list),len(new_name_list))):
    shutil.move(target_file_list[i],new_name_list[i]+f".{target_format}")
    renamed_file_count+=1
print(f"{renamed_file_count} files've been renamed successfully~~~!")
input()