from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
import argparse

#Copy range of cells as a nested list
#Takes: start cell, end cell, and sheet you want to copy from.
#src: https://yagisanatode.com/2017/11/18/copy-and-paste-ranges-in-excel-with-openpyxl-and-python-3/
def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
 
    return rangeSelected

#Paste range
#Paste data from copyRange into template sheet
#src: https://yagisanatode.com/2017/11/18/copy-and-paste-ranges-in-excel-with-openpyxl-and-python-3/
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData):
    countRow = 0
    for i in range(startRow,endRow+1):
        countCol = 0
        for j in range(startCol,endCol+1):
            
            sheetReceiving.cell(row = i, column = j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1

def split_workbook(file, dir_out, split_size):
    # Load workbook
    wb = load_workbook('test.xlsx')

    # Get first sheet
    sheet = wb.worksheets[0]

    # Get num of cols
    total_cols = sheet.max_column

    # Get max num of rows
    max_rows = sheet.max_row

    # Get number of cols each subfile should hold, and remainder
    (num_files, remainder) = divmod(total_cols, split_size)

    if remainder > 0:
        print('Warning: the last subfile will have {} + {} columns'.format(split_size, remainder))
    
    # Create subfiles containing only split_size number of cols of original
    for subfile in range(num_files):
        # Create new workbook name
        dest_filename = '{}_test.xlsx'.format(subfile)
        # Create new workbook
        new_file = Workbook()
        new_sheet = new_file.active
        # Get data to copy
        if subfile == 0:
            start_col = 1
            end_col = subfile*split_size+split_size
        elif subfile == len(range(num_files))-1:
            start_col = subfile*split_size+1
            end_col = subfile*split_size+split_size+remainder
        else:
            start_col = subfile*split_size+1
            end_col = subfile*split_size+split_size
        # Copy data
        selected_data =  copyRange(start_col, 1, end_col, max_rows, sheet)
        # Paste data
        if subfile == len(range(num_files))-1:
            pasteRange(1, 1, split_size+remainder, max_rows, new_sheet, selected_data)
        else:
            pasteRange(1, 1, split_size, max_rows, new_sheet, selected_data)
        print('Filled subfile {}'.format(subfile))
        new_file.save(filename = dest_filename)


if __name__ == '__main__':
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Splits 1 excel file into X smaller files, as equally as possible based on # of columns')
    parser.add_argument('file', help='Excel file to split')
    parser.add_argument('dir_out', help='Directory to save split files to')
    parser.add_argument('split_size', help='Number of columns to split by')

    # Get arguments
    args = parser.parse_args()

    # Save arguments locally
    file = args.file
    dir_out = args.dir_out
    split_size = args.split_size

    split_workbook(file, dir_out, int(split_size))
