Attribute VB_Name = "Module1"
Sub M00_Combined1n2()
'
' M00_Combined1n2 ������
'

'
    ' ����������� ���������� �������������
    Dim var As Double
    var = Selection
    var = (var - 10#) * (-1#)
    
    ' �������� ������� �����������
    Sheets(1).Select
    Columns("A:D").Select
    Selection.ClearContents
    
    ' �������� ��������� ���. ������
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
    
    ' �������� ������ �����
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
    
    ' ������� � ��������� ������
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
    
    ' �������������� �� �������
    Sheets(2).Select
    Range("D1").Select
    Range("D1").FormulaLocal = "=����(E1<>""""; (E1-10.0)*(-1.0); """")"
    Range("D1").Select
    Selection.AutoFill Destination:=Range("D1:D1000"), Type:=xlFillDefault
    Range("D1:D1000").Select
    
    ' ������� - � ��������
    Dim rng As Range
    For Each rng In Selection
        If rng.HasFormula Then
            rng.Formula = rng.Value
        End If
    Next rng
    
    ' ����������� ������
    Range("C1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Cells(1)).Select
    Selection.Copy
    
    ' ������� � ������ ����
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Paste
    
    ' ��������� ������ ������ � "G1 X"
    Range("A1").Value = "G1 X"
    Range("D1").Value = var
    ' ����������� �����
    'Selection.Copy
End Sub




