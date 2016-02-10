#Refreshable OData Queries in Excel Add-In

This sample demonstrates how to create an Excel add-in that performs OData queries and inserts data into worksheets. Each entity from the OData source is placed into a new worksheet. When reloading the data, the existing worksheet is cleared and new data is inserted.
![](http://i.imgur.com/XogNcYC.png)

##Running the sample
1. Run this sample by downloading the sample code using `git clone https://github.com/dougperkes/Office-Add-Excel-RefreshableWorksheets.git`.
2. Open the solution file ExcelWorksheets\ExcelWorksheets.sln
3. Press F5 to start Excel
4. Click Yes on the security alert window.
5. Click one of the entity names to load a worksheet with the data.

##Note on Security Alert when running the add-in

This sample uses Northwind data from odata.org. Unfortunately, the SSL certificate has expired and you will be presented with the following warning when loading the add-in. Click **Yes** to continue.
![](http://i.imgur.com/rm1GwjJ.png) 