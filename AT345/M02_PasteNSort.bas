Attribute VB_Name = "Module3"
Sub M02_PasteNSort()
'
' M02_PasteNSort Макрос
'

'
    ActiveSheet.Paste
    Columns("B:B").Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A:D")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("D:D").Select
    ActiveSheet.Range("$A:$D").RemoveDuplicates Columns:=4, Header:=xlNo
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub


