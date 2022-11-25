##############################################################################
#
# File Converter: TXT to XLSX
# Each line in the file will be placed in a new row in xlsx file
# Each string tab-separated will be placed in a new col in xlsx file
# By: Robinson Coutinho - 20/11/2022
#

import sys
import xlsxwriter

if (len(sys.argv) != 3):
    print('Invalid number of arguments')
    print('Usage: python XlsxWriter.py SourceFile.txt DestFile.xlsx')
    sys.exit(1)
    
if(not(".txt" in sys.argv[1].lower())):
    print('Invalid argument "' + sys.argv[1] + '"')
    print('Usage: python XlsxWriter.py SourceFile.txt DestFile.xlsx')
    sys.exit(1)

if(not(".xlsx" in sys.argv[2].lower())):
    print('Invalid argument "' + sys.argv[2] + '"')
    print('Usage: python XlsxWriter.py SourceFile.txt DestFile.xlsx')
    sys.exit(1)

try:
    input_file = open(sys.argv[1], "r")
except IOError:
    print("Error reading file " + sys.argv[1])
    sys.exit(1)
    
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook(sys.argv[2])
worksheet = workbook.add_worksheet()

col= 0
row= 0
while(True):
    line= input_file.readline()
    if not line:
        break;
    splits = line.split('\t');    
    for split in splits:
        if(".bmp" in split.lower())or(".png" in split.lower())or(".jpg" in split.lower()):
            worksheet.insert_image(row, col, split)
        else:    
            worksheet.write_string(row, col, split)
        #print(split);
        col= col + 1
    row = row + 1
    col = 0
input_file.close()
workbook.close()
print("Converter finished")
sys.exit(0)

