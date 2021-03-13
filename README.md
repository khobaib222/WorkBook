# WorkBook

## Frequent operations with Excel WorkBooks  
###### 1. Read all WorkBooks - List WorkBooks and their WorkSheets name along with indicating emtpy WorkSheets.  
###### 2. Read a single WorkBook - List a WorkBook and its WorkSheets along with indicating empty WorkSheets.  
###### 3. Delete a list of WorkSheets from a WorkBook.  
###### 4. Add a list of WorkSheets from one WorkBook to another.  
###### 5. Convert a list of xlsx format WorkBooks to xls.  
###### 6. Convert a list of xls format WorkBooks to xlsx.  
###### 7. Aggregate all the WorkSeets across WorkBooks to a single WorkBook.  
###### 8. Aggredate list of WorkSheets who have similar headers into a single sheet.  

This project can make all these operations to be performed a lot easier.  

## Steps  
###### 1. Make sure latest node and npm versions are installed on your system.  
###### 2. Run **npm install** in project directory.  
###### 3. Make a folder with name **commandBook**.  
###### 4. Make a folder with name **resourceBook**.  
###### 5. Inside commandBook folder make a WorkBook with name **commandBook.xlsx** having a sheet with name **commandSheet**.  
###### 6. Pull all the WorkBooks on which you want to perform the operations into **resourceBook** folder.  

commandSheet contains all the commands to perform the operations on workbooks in your resourceBook folder.  

## Commands  

###### 1. READ  
This will read all the workbooks in you resouceBook folder and list it down in your commandSheet.  

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111022508-22ddf900-83f9-11eb-8c15-633ad4604bb4.png)|![image](https://user-images.githubusercontent.com/24797779/111022659-18702f00-83fa-11eb-9e71-04ff24c17ab2.png)

This will give the result in the order **WorkBook Name -> WorkSheets -> ...**. 

WorkSheet names will be ```','``` separated. The left side of ```'|'``` contains all the sheets which do not have all cells empty while those in the right side of it will be all empty sheets.  

###### 2. READ_BOOK
This will read a single workbook and list down its worksheet names. 

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111022819-430eb780-83fb-11eb-81a6-f38c265b40c1.png)|![image](https://user-images.githubusercontent.com/24797779/111022834-57eb4b00-83fb-11eb-9515-6434c47da70a.png)

###### 3. DELETE
This will Delete all the given list of sheet names from the workbook. 

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111022887-b1ec1080-83fb-11eb-839e-5cc30d690e88.png)|![image](https://user-images.githubusercontent.com/24797779/111022892-c7613a80-83fb-11eb-9747-872406f6e9d8.png)

**Note: The two READ_BOOK commands are given just for illustration purpose to check if the operation has been performed correctly or not and also in the first READ_BOOK command the next cells after the book name are just the result of previous operation they are not arguments to be passed**. 

###### 4. ADD
This will add a list of worksheets from once workbook to another

The arguments in differs cells are source book followed by source sheets followed by destination book.  

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111023103-1196eb80-83fd-11eb-8efe-918942d964cd.png)|![image](https://user-images.githubusercontent.com/24797779/111023144-528f0000-83fd-11eb-9a4c-b0ce48a4f9ea.png)

If the destination book already contains a work sheet with similar name then the sheet getting added will be assigned new name: here in this example Sheet1-1.  





