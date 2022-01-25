Attribute VB_Name = "Module2"
Sub M01_MakeNumPage()
'
' M01_MakeNumPage Макрос
'

'
	' сделать лист, состоящий из чисел с 8-и знаками после запятой
    Cells.Select
    Selection.NumberFormat = "0.00000000"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 19.29
    Range("A1").Select
End Sub
