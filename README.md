# json-excel
Convert JSON string to cell values and Cell Ranges to JSON string

INCOMPLETE but usuable

=JsonFromCells(A1:C3) => { 'a1 value': 'b1 value'}

=JsonToArray("{ 'title': 'a json string', 'summary': 23}", oritentation=0|1) => 2x2 array (CTRL-SHIFT-ENTER for formula arrays)

=JsonLookup("{ 'title': 'a json string'}","title")  => "a json string
