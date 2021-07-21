Sub FindReplaceRange(ByVal SheetName As String, ByVal SearchColumn As String, ByVal SearchElement As String, ByVal ReplaceForElement As String)
	Worksheets(SheetName).Columns(SearchColumn).Replace _ 
	What:=SearchElement, Replacement:=ReplaceForElement, _ 
	SearchOrder:=xlByColumns, MatchCase:=True
End Sub