# VBA Utilities
This repository contains useful and frequently used functions. 

Below are a list of useful fragments of code that could prove useful.

# Get Username
```
strName =Environ("username")
```

# Modulo
```
i Mod 35
```
# Create array
```
Dim aiArr As Variant
stringArr = Array(" AI ", "artificial intelligence", "artificial int")
```

# Click into cell and hit enter
```
Sub clickAndEnter()
	Dim wkSht As Worksheet
	Set wkSht = Worksheets("Survey form")

	Dim rCell As Range
	Set rCells = Range(wkSht.Range("A1"), wkSht.Range("A1").End(xlDown))


	Dim rowCount, iRow As Long
	rowCount = rCells.Rows.Count

	    For iRow = 2 To rowCount
		Application.SendKeys "{F2}"
		Application.SendKeys "{ENTER}"
	    Next iRow
End Sub
```

# Errors

Resumes next in Loop 
```
On Error Resume Next
```
Causes break in code

```
On Error GoTo 0
```

Goes to specified section and runs code

```
On Error GoTo alertMsg

alertMsg:
	Msgbox(“Alert!”)
```
# Set Workbook
```
Set activeWkb=ThisWorkbook
 ```
 
# Status bar updating
```
Application.DisplayStatusBar=True
Application.StatusBar = x & “%”
Application.StatusBar = False
```
# Create New Workbook
```
Workbooks.Add
```
# Worksheet Function
```
Debug.Print ("length of columns: " & Application.WorksheetFunction.CountA(dataSht.Range("A:A")))
```
# Count worksheets
```
MsgBox Worksheets.Count
```
# Select Multiple Ranges
```
Range(“A1:A2,B2:C4”).value=10
```
# Cells in range:
```
Range(Cells(1,1),Cells(4,1))=5
```
# EmptyClipboard
```
Application.CutCopyMode=False
```
# Clear Contents
```
Workrange.ClearContents
```
# Filter – returns an array of elements in the filtered array that match
```
Arr_tem=Filter(Original_array, Some_variable_to_look_up)
```
# Join – join strings from array
```
Join(Array_temp,”+”)
```
# Offset
```
Range(“A1”).offset(row,column).value=”offsetRow”
```
# Split (text,delimeter)
```
Split(“A brand new string”,” “)
```
# Set cell formulas & filldown
```
Wksht.range(“A1”).formula=”=Sum(B1:D1)”
Wksht.range(“A1:A” & 50).filldown
```

# Text to columns with any character
```
Set workrange = Range("A1:A5")
delimTxt = "|"
workrange.TextToColumns Destination:=workrange, DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
:=delimTxt
```

# Yes/No
```
Dim Answ As Integer
Answ=MsgBox(“Continue?”,vbYesNo)
```

# Find Substring
```
If InStr(1, Range("A1”).Value, "Chicago") <> 0 Then
    Debug.print("Range A1 contains Chicago")
End If
```
# Put a break into your code
```
Stop
```

# VBA Timer
```
Dim x As Double
x = Timer
Debug.Print Timer - x
```

# Pivot Table General use
```
For Each PT In sht.PivotTables        '<~~ Loop all pivot tables in worksheet
      PT.ClearAllFilters '<~~ Clear filters from pivot table
      PT.HasAutoFormat = False	'<~~ Stop auto formatting
      PT.PivotCache.Refresh	'<~~ Refresh
      PT.PivotFields ("Sum")
      PT.Calculation = xlPercentOfRow
      PT.NumberFormat = "0,0%"
Next PT
```

# Group - show lower level of group -or show all grouped
```
sht.Outline.ShowLevels 2  '<~~ shows level 2
sht.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1 '<~~ shows highest level (collapses group)
```

# Frequently used Formulas
## Non-Data Analysis Regression

Select 3 cells horizontally and press ctrl+Shift+Enter – gives x2+x+Constant
```
=Linest(y-range,xrange^{1,2},TRUE,FALSE)
```
## Rank
for last argument: 0 - descending, 1 - ascending
```
=RANK.AVG(a1,$a$1:$a$5,0)
```
## Sub-Total - shows sum of filtered rows
```
=SUBTOTAL(9,A2:A100)
```

## Count of Unique Values
```
=Sumproduct(1/countif(a2:a5,a2:a5))
```

## Reverse String
```
=CONCAT(MID(A1,SEQUENCE(LEN(A1),,LEN(A1),-1),1))
```

## Get Last Name - find first space in reversed string
```
=RIGHT(A1,FIND(" ",CONCAT(MID(A1,SEQUENCE(LEN(A1),,LEN(A1),-1),1)))-1)
```

## Weekday
```
=text(A1,”dddd”)   - Where A1=01/01/2017
```

## Add 12 Months
```
=Edate(oldDate,12)
```

## Array Formula - below is superseded by MAXIFS
```
{=Max(If(A2:A4=”Dog”,B2:B4))}
```

## Indirect Formula
```
=indirect(“Calculation!A1”)
```

## File name
```
=cell(“filename”)
```

## Get worksheet name
```
=replace(cell(“filename”),1,find(“]”,cell(“filename”)),””)
```

## Date Difference
```
=DateDif(a1,a2,”m”) – can also be d/m/y
```
## Year Frac
```
=yearfrac(a1,a2)
```
## Errors?
```
=iferror(A1,”There’s an error”)
```

## Format text
```
=text(a1,”00”)
```

## Classic index match
```
=index(D:D,match(“Cat”,”A:A,0))
```

## Substitute
```
=substitute(A1,”ReplacementString”,A2)
```

## Format US to UK Date - assumes you've got the year
```
=IF(TYPE(A2)=2,DATE(RIGHT(a2,4),LEFT(A2,FIND("/",A2)-1),MID(A2,1+FIND("/",A2),FIND("/",SUBSTITUTE(A2,"/","",1))+1-(1+FIND("/",A2)))),DATE(RIGHT(a2,4),DAY(A2),MONTH(A2)))
```

## Create a 5 year age bracket
From an age i.e. 45 - converts its to 45-49
```
=FLOOR.MATH(A1/5)*5&"-"&FLOOR.MATH(A1/5+1)*5-1
```


## Data Analysis
1.	Click the File tab, click Options, and then click the Add-Ins category.
2.	In the Manage box, select Excel Add-ins and then click Go.
3.	In the Add-Ins available box, select the Analysis ToolPak check box, and then click OK.

# DAX Notes

## Create a measure to act as a percentage of a sub-total
### In below split figures I wanted for a bar chart split by Gender
```
PercentageGenderFTE = 
CALCULATE(
	SUM(Staff[FTE]),
	ALLSELECTED(Staff[Year])
)/CALCULATE(
	SUM(Staff[FTE]),
	ALLEXCEPT(Staff,Staff[Year],Staff[Gender])
)
```

## Sumif as a measure
```
sumif = SUMX(FILTER(TBL,TBL[Category]="Complete"),Tbl[value])
```

## Sumifs for a column
```
SUMIFS = calculate(SUM(TBL[value]),FILTER(TBL,TBL[Gender]=currentTbl[Gender] && TBL[Category] IN {"Good","Bad","Ugly"}),TBL[value])
```

## Countifs as measure
```
countifs = COUNTAX(filter(TBL,TBL[Category]="Complete" && TBL[Include]=True),TBL[Category])
```

## AverageIf as measure
```
averageif = AVERAGEX(Filter(TBL,TBL[Category]="Complete"),[value])
```

## medianIf as measure
```
medianif = MEDIANX(Filter(TBL,TBL[Category]="Complete"),[value])
```

## percentage measure
What percent are Complete?
```
PercentMeasure = COUNTAX(filter(Tbl,Tbl[Category]="Complete"),Tbl[Category])/COUNTA(Tbl[Category])
```
## Count rows All Except
Lets you count everything that has been filtered - you may need to have multiple columns if this is part of a graph
In the below example we're counting the rows of where Table[Col1] = "Value 1", but we're trying to create a line graph with Table[Col2] on the X-Axis, so need to have that excluded too, or it will create a constant line
```
CountAllExcept = CALCULATE(COUNTrows(filter(Table,Table[Col1]="Value 1")),AllEXCEPT(Table,Table[Col1],Table[Col2]))
```

## Have a Slicer to switch between Measures

Step 1 - create a table with the name of the measures you want like the below. To create this use (table will be called Selector):

```
Selector = UNION(ROW("Name","This period's sales","Code",1),ROW("Name","Year to date","Code",2),ROW("Name","Quarter to date","Code",3),ROW("Name","Same period last year","Code",4),ROW("Name","Cost","Code",5))
```
Will create a table like the below
![image](https://user-images.githubusercontent.com/29797377/129887957-edf69404-dc10-4c74-a963-fcf463199d3a.png)

Step 2 - Add a Measure to the above table (when you use a slicer on the name column of this table it'll affect your measures)
```
Selected Measure = SELECTEDVALUE('Selector'[Code],1)
```
Step 3 - In the table with your existing measures add the below measure - (this will be the value in whatever graph you create)

```
Sales All Measures =
SWITCH(
Selector[Selected Measure],
1,[Sales],
2,[Sales YTD],
3,[Sales QTD],
4,[Sales Same Period Last Year],
5,[Cost]
)
```
Step 4 - Add your 'Sales All Measures' as a value in your graph & add a slicer with 'Selected Measure' as the value

Optional Extra 1 - You can use your selected category to calculate a variable using a certain variable
```
Proportion (Dynamic) = 
VAR CategorySelection = SELECTEDVALUE('SelectorTbl'[Selection],"Gender")
RETURN
SWITCH(
    TRUE(),
    CategorySelection="Age Group",CALCULATE([Proportion],TREATAS(VALUES('SelectorTbl'[Category]), DataTable[Age Group])),
    CategorySelection="Gender",CALCULATE([Proportion],TREATAS(VALUES('SelectorTbl'[Category]), DataTable[Gender]))
)
```

Optional Extra 2 - you can split out a graph using small multiples, so if you have a table called 'Slicer Table' organised like:

### Selection	###Category
Gender	F
Gender	M
Age Young
Age Old

```
HeightCategorised= 
VAR CategorySelection = SELECTEDVALUE('Slicer Table'[Selection],"Gender")
RETURN
SWITCH(
    TRUE(),
    CategorySelection="Gender",CALCULATE([AverageHeight],TREATAS(VALUES('Slicer Table'[Category]),Data[Gender])),
    CategorySelection="Age",CALCULATE([AverageHeight],TREATAS(VALUES('Slicer Table'[Category]),Data[Age]))))
)
```

## Summarize & Max of summarize

In the below we create a summarised table - this is equivalent to Grouping by City and Year. In the second part the table taken as an argument is a filtered table, where we want to exclude "London"
```
SUMMARIZE(Staff,Staff[City],Staff[Year],"Turnover",Staff[Turnover])

SUMMARIZE(FILTER(Staff, Staff[City]<>"London"),Staff[City],Staff[Year],"Turnover",Staff[Turnover])
```

In the below example we want to find the max turnover for each city and gender. This creates a Union table summarising by city, year & gender, year - then finds the maximum turnover. This can be used as one argument
```
MaxTurnover = 
	MAXX(
	UNION(
		SUMMARIZE(FILTER(Staff, Staff[City]<>"Unknown"),Staff[City],Staff[Year],"Turnover",Staff[Turnover]),
		SUMMARIZE(Staff,Staff[Gender],Staff[Year],"Turnover",Staff[Turnover]))
		,[Turnover])
```

## Add a slider as a slicer
You will probably need to add a table with a series. Below we create a table called DateParameters
```
DateParameters = GENERATESERIES(-37, -1, 1)
```

You then need to use that as a Slicer and somehow relate that to the table you are slicing. You might want to do it by adding a column with
```
MonthsPrevious = (DATEDIFF([Date],edate(Max([Date]),1),MONTH)*-1)
```
## Create a manual table
```
NewTable = UNION(
		ROW("Row Number",1,"Animal","Cat","Special Ability","Scratch"),
		ROW("Row Number",2,"Animal","Dog","Special Ability","Bark"),
		ROW("Row Number",3,"Animal","Kraken","Special Ability","Engulf")
		)
```
OR
```
NewTable = DATATABLE("Animal",STRING,"Weight",DOUBLE,{{"Dog",55.5},{"Cat",32.4},{"Crocodile",165}})
```

## Python Compatability issues with Python:

- I had to install anaconda on my laptop
- I then had to create an old conda of Python 3.5 and install all the packages there. I then had to change the 
python package that Power BI uses to this old 3.5 version - named old_python_version. It'll be in
C:\Users\Anaconda3\envs\old_python_3.5
```
	conda create --name old_python_version python=3.5 pip
	conda activate old_python_version
	pip install pandas
	pip install numpy
	pip install scipy
	pip install matplotlib
```
