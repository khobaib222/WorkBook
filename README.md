# WorkBook

## Frequent operations with Excel WorkBooks  
###### 1. Read all WorkBooks - List WorkBooks and their WorkSheets name along with indicating emtpy WorkSheets.  
###### 2. Read a single WorkBook - List a WorkBook and its WorkSheets along with indicating empty WorkSheets.  
###### 3. Delete a list of WorkSheets from a WorkBook.  
###### 4. Add a list of WorkSheets from one WorkBook to another.  
###### 5. Convert a list of xlsx format WorkBooks to xls.  
###### 6. Convert a list of xls format WorkBooks to xlsx.  
###### 7. Aggregate all the WorkSeets across WorkBooks to a single WorkBook.  
###### 8. Aggregate list of WorkSheets who have similar headers into a single sheet.  
###### 9. Rename WorkSheets in a WorkBook.  

This project can make all these operations to be performed a lot easier.  

## Steps  
###### 1. Make sure latest node and npm versions are installed on your system.  
###### 2. Run ```npm install``` in project directory.  
###### 3. Make a folder with name ```commandBook```.  
###### 4. Make a folder with name ```resourceBook```.  
###### 5. Inside ```commandBook``` folder make a WorkBook with name ```commandBook.xlsx``` having a sheet with name ```commandSheet```.  
###### 6. Pull all the WorkBooks on which you want to perform the operations into ```resourceBook``` folder. 
###### 7. Fill ```commandSheet``` with commands.  
###### 8. Close ```commandBook``` workbook and all other books of ```resouceBook``` folder.  
###### 9. Run ```node index.js```. 

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

The arguments in different cells are source book followed by source sheets followed by destination book.  

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111023103-1196eb80-83fd-11eb-8efe-918942d964cd.png)|![image](https://user-images.githubusercontent.com/24797779/111023144-528f0000-83fd-11eb-9a4c-b0ce48a4f9ea.png)

If the destination book already contains a work sheet with similar name then the sheet getting added will be assigned new name: here in this example Sheet1-1.  

###### 5. RENAME
This will rename the worksheets of a workbook. 

The arguments in different cells are source book followed by ```','``` separated sheets with their original name and new name separated by ```'|'```.  


**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111024147-0e066300-8403-11eb-9013-00d91b23a4d2.png)|![image](https://user-images.githubusercontent.com/24797779/111024154-22e2f680-8403-11eb-94a9-e8514434612e.png)

###### 6. AGG-WB
This will aggregate all the worksheets in the given workbooks.  

The arguments in different cells are ```','``` separated workbook names followed by the aggregated workbook name.  


**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111024305-1c08b380-8404-11eb-9830-37d6949cde05.png)|![image](https://user-images.githubusercontent.com/24797779/111024328-3b074580-8404-11eb-9b2e-4f911bb722be.png)

###### 7. AGG-WS
This will aggregate the work sheets with similar header into a single worksheet of a workbook. 

The arguments in different cells are workbook name followed by the ```','``` separated sheet names. If sheet names are not provided then all the sheets will be aggregated with similar headers of that work book. 

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111024616-b4ebfe80-8405-11eb-9788-99916eff631b.png)|![image](https://user-images.githubusercontent.com/24797779/111024601-a3a2f200-8405-11eb-8820-3f339cfc78b8.png)

**Note: This operation works only on simple sheets with simple headers. Having only first row as header without any merge cells. Also the name of the generated sheets can be any unique name. Since the use won't be knowing how many different sheets can be generated because of different headers so specifying arguments is not possbile for new sheet names. Although a list of new sheet names can be given equal to the total number of current sheets and then only some of the first names would be picked up as the new names. But this is not supported yet**.  

###### 8. TO_XLS
This will convert the provided workbooks from xls/xlsx to xls. 

The arguments would be the list of ```','``` separted workbook names.  If no argument is provided all the workbooks will be converted to xls format. 
The same workbook is not converted rather a new book is added with same name and xls format. 

**Input** | **Result**
-----------|-----------
![image](https://user-images.githubusercontent.com/24797779/111024901-16609d00-8407-11eb-80b2-26998653ee71.png)|![image](https://user-images.githubusercontent.com/24797779/111024957-5d4e9280-8407-11eb-966c-ddacfc722791.png)



###### 9. TO_XLSX
This will convert the provided workbooks from xlsx/xls to xlsx. 

The arguments would be the list of ```','``` separted workbook names.  If no argument is provided all the workbooks will be converted to xlsx format. 
The same workbook is not converted rather a new book is added with same name and xlsx format. 


