Sub ProtectExcelSheet(ByVal SheetName As String, ByVal Password As String)

Worksheets(SheetName).Protect Password

End Sub