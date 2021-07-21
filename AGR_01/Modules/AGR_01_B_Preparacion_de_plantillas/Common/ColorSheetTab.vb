Sub ColorSheetTab(ByVal SheetName As String)
	'declare a variable
	Dim ws As Worksheet
	Set ws = Worksheets(SheetName)
	'color a worksheet, named Sheet1, in green
	ws.Tab.ColorIndex = 4
End Sub