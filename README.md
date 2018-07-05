# json-excel
Convert JSON string to cell values and Cell Ranges to JSON string

INCOMPLETE but usuable

Simply paste a JSON formatted string into a cell and use the worksheet functions to extract the properties.

Maximum JSON string length is 32767.

## Install

After compiling, add the AddIn (from the Options menu or via the Alt+t, i shortcut) 
Browse to JsonExcel-AddIn-packed.xll (in bin/Debug or bin/Release)


## From JSON to cells

* =JsonToArray("{ 'title': 'my title', 'summary': "sum"}")
	=> 2x2 array (CTRL-SHIFT-ENTER for formula arrays)

* =JsonLookup("{ 'title': 'my title', 'summary': 'sum'}","title")  
	=> "title"

* =JsonLookup("{ 'title': 'my title', 'count': 23}","count")  
	=> 23

## From Cells to JSON
* =JsonFromCells(A1:C3) 
	=> { 'a1 value': 'b1 value'}

### JsonToArray
* json string

* orientation
	0 => vertical
	1 => horizontal
