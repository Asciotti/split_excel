from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import argparse

#Copy range of cells as a nested list
#Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
 
    return rangeSelected

#Paste range
#Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData):
    countRow = 0
    for i in range(startRow,endRow+1,1):
        countCol = 0
        for j in range(startCol,endCol+1,1):
            
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
    
    # Create subfiles containing only split_size number of cols of original
    for subfile in range(num_files):
        if subfile == 0:
            start_col = get_column_letter(1)
            end_col = get_column_letter(subfile*split_size+split_size)
        else:
            start_col = get_column_letter(subfile*split_size+1)
            end_col = get_column_letter(subfile*split_size+split_size)
        # Copy data
        selected_data =  copyRange(start_col, end_col, 1, 20, sheet)
        # Paste data


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

