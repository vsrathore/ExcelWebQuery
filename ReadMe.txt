You can use modified webquery to get Stock Data in Excel.

For example you want to get price of Google for Date: 23rd Jun 2015

Url for that would be: "http://finance.yahoo.com/q/hp?s=GOOG&a=05&b=23&c=2015&d=05&e=23&f=2015&g=d"

Here variable(s) in Url are "StockName" & "Dates".

Now create a new webquery for this url request:

#########
WEB
1
http://finance.yahoo.com/q/hp?s=["Stk",""]&a=["stM",""]&b=["stD",""]&c=["stY",""]&d=["endM",""]&e=["endD",""]&f=["endY",""]&g=d

Selection=15
Formatting=None
PreFormattedTextToColumns=True
ConsecutiveDelimitersAsOne=True
SingleBlockTextImport=False
DisableDateRecognition=False
DisableRedirections=False

#########

Copy above text between '#' marks and paste in blank notebook, and save as "yahoo.iqy" on your desktop.
Here .iqy is Excel WebQuery File Ext.

1: Now open a new blank workbook.
2: Then Choose Data->Exixting Connection->Yahoo.iqy
3: A dialogue box will open asking you to select , browse to the webquery(Yahoo.iqy) you just saved on desktop and select it.
4: Then ‘Import Data’ box open. Click 'OK'.
5: Then it will ask for variable(s). Fill all required parameters and Done.


Example worksheet is on my GitHub:
vsrathore/ExcelWebQuery

Download Excel and WebQuery file.

About ExcelFile:
B1= "Stock_Name"

B/C/D::4/5 = "Date Parameters"

WebQuery in Cell A8

Note: We can use same method for get Stock data for a period(StartDate to EndDate). Here i use StartDate=EndDate for single data point in excel workbook.

Feel free for ask more query (or WebQuery) :P
