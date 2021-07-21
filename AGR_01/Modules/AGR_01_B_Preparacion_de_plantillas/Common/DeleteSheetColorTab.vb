Sub DeleteSheetColorTab()
    Dim xSheet As Worksheet
    For Each xSheet In ActiveWorkbook.Worksheets
        xSheet.Tab.ColorIndex = xlColorIndexNone
    Next xSheet
End Sub