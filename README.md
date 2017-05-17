### xl2csv  
Translate a xlsx table with wide-short data to a csv with narrow-long data.  

#### Dependencies  
Python 3.3 or newer  
XLRD library

#### Install Instructions  
Create a new directory.  
    python3 pip install xlrd  
    git clone https://github.com/gnfrazier/xl2csv.git  
  
#### Transforming a Sheet  
Put excel file in the same working directory.  
Excel sheet must have column names in row 1 or be formatted as a table.  
Open a python command prompt (REPL)  
    import xl2csv as xl  
      
    workbook = 'excel-file-to-open.xlsx'
    name = 'Sheet1' #sheet name  
    ignore = ['columnname1', 'columnname2', 'columnname3'] #list of columns to ignore
    path = xl.get_path(workbook)  
    xlfile = xl.open_file(path)  
    sheet = xl.get_sheet(xlfile, name)  
    header = xl.get_header(sheet)  
    xl.to_columns(sheet, header, ignore)  
  
#### More information
Heavily influenced by https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/  


