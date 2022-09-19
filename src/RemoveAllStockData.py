import os

def main():
    excelFileName = 'all_stock_data.xlsx'
    currentPath = str(os.path.abspath(os.getcwd()))
    desiredPath = currentPath.replace('src', 'US')
    excelFilePath = f'{desiredPath}/{excelFileName}'

    if os.path.exists(excelFilePath):
        os.remove(excelFilePath)
        print(f'\nDeleted old {excelFilePath} file.') 
    else:
        print('\nNo pre-existing stock data file exists. ')

if __name__ == "__main__":
    main()