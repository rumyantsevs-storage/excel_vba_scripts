Attribute VB_Name = "Module1"
Sub M00_Combined1n2()
'
' M00_Combined1n2 Макрос
'

'
    ' Запоминание введенного пользователем
    Dim var As Double
    var = Selection
    var = (var - 10#) * (-1#)
    
    ' Удаление старого содержимого
    Sheets(1).Select
    Columns("A:D").Select
    Selection.ClearContents
    
    ' Удаление расчетных доп. листов
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        For Each ws In ActiveWorkbook.Sheets
            If ws.Index <> 1 Then
                ws.Delete
            End If
        Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Создание нового листа
    Sheets.Add After:=Sheets(Sheets.Count)
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
    
    ' Вставка и обработка данных
    Sheets(2).Paste
    Columns("B:B").Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A:E")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("E:E").Select
    ActiveSheet.Range("$A:$E").RemoveDuplicates Columns:=5, Header:=xlNo
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    ' Преобразование по формуле
    Sheets(2).Select
    Range("D1").Select
    Range("D1").FormulaLocal = "=ЕСЛИ(E1<>""""; (E1-10.0)*(-1.0); """")"
    Range("D1").Select
    Selection.AutoFill Destination:=Range("D1:D1000"), Type:=xlFillDefault
    Range("D1:D1000").Select
    
    ' Формулы - в значения
    Dim rng As Range
    For Each rng In Selection
        If rng.HasFormula Then
            rng.Formula = rng.Value
        End If
    Next rng
    
    ' Копирование данных
    Range("C1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Cells(1)).Select
    Selection.Copy
    
    ' Вставка в первый лист
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Paste
    
    ' Изменение первой ячейки в "G1 X"
    Range("A1").Value = "G1 X"
    Range("D1").Value = var
    ' копирование всего
    'Selection.Copy
End Sub




