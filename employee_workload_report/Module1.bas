Attribute VB_Name = "Module1"
Sub Сформировать_отчет_по_загрузке()
Attribute Сформировать_отчет_по_загрузке.VB_Description = "Составитель отчета по загрузке сотрудников отделов. Использует данные активной книги Excel, выводит данные в документ Word. Если документ ""Отчет по загрузке.docx"" присутствует в каталоге данного дакумента, - предупреждает о перезаписи."
Attribute Сформировать_отчет_по_загрузке.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Сформировать_отчет_по_загрузке Макрос
'

'
    Dim Ids() As String
    Dim Fios() As String
    Dim Divisions() As Integer
    Dim IdTasks() As Integer
    Dim DivisionTasks() As Integer
    
    Dim IdCount As Integer

    ' Заполнение массива c таб.номерами сотрудников.
    Set NumberCells = Worksheets("Сотрудники").Range("A2:A" & CStr(Rows.count))
    
    i = 0
    For Each Cell In NumberCells
        If Cell.Value <> "" Then
            ReDim Preserve Ids(i)
            Ids(i) = Trim(Cell.Value)
            i = i + 1
        Else
            Exit For
        End If
    Next
    
    IdCount = UBound(Ids()) + 1
    
    ' Заполнение массива с именами сотрудников.
    Set FamiliyaCells = Worksheets("Сотрудники").Range("B2:B" & CStr(Rows.count))
    
    i = 0
    For Each Cell In FamiliyaCells
        If i < IdCount Then
            ReDim Preserve Fios(i)
            Fios(i) = Trim(Cell.Value)
            Set ImyaCell = Cell.Offset(0, 1)
            If ImyaCell.Value <> "" Then
                Fios(i) = Fios(i) & " "
                Fios(i) = Fios(i) & Left(Trim(ImyaCell.Value), 1) & "."
                Set OtchestvoCell = Cell.Offset(0, 2)
                If OtchestvoCell.Value <> "" Then
                    Fios(i) = Fios(i) & Left(Trim(OtchestvoCell.Value), 1) & "."
                End If
            End If
            i = i + 1
        Else
            Exit For
        End If
    Next
    
    ' Заполнение массива с номерами отделов
    Set DivisionCells = Worksheets("Сотрудники").Range("F2:F" & CStr(Rows.count))
    
    i = 0
    For Each Cell In DivisionCells
        If i < IdCount Then
            ReDim Preserve Divisions(i)
            Divisions(i) = CInt(Trim(Cell.Value))
            i = i + 1
        Else
            Exit For
        End If
    Next
    
    ' Заполнение массива с количеством задач
    For i = 0 To IdCount - 1
        ReDim Preserve IdTasks(i)
        IdTasks(i) = GetSumOfIdTasks(Ids(i))
    Next
    
    ' Сортировка 4 массивов: 1) по кол-вам задач сотрудников, 2) по номерам отделов
    Call SortingOf4Arrays(IdTasks(), Ids(), Divisions(), Fios(), 0, IdCount - 1)
    Call SortingOf4Arrays(Divisions(), Ids(), IdTasks(), Fios(), 0, IdCount - 1)
    
    ' Заполнение массива с количествами задач отдела
    
    For i = 0 To IdCount - 1
        ReDim Preserve DivisionTasks(i)
        DivisionTasks(i) = GetSumOfDivisionTasks(Divisions(i), Divisions(), IdTasks())
    Next
    
    ' Сортировка 5 массивов по по кол-вам задач отделов
    Call SortingOf5Arrays(DivisionTasks(), Divisions(), Ids(), IdTasks(), Fios(), 0, IdCount - 1)
        
    ' Проверка на наличие старого отчета Word
    Dim strFileName As String
    Dim strFileExists As String
     
    strFileName = ActiveWorkbook.Path & "\Отчет по загрузке.docx"
    strFileExists = Dir(strFileName)
    
    If strFileExists <> "" Then
        If MsgBox("Отчет уже был сформирован, хотите обновить его?", vbYesNo + vbQuestion) = vbYes Then
            If IsFileLocked(strFileName) Then
                MsgBox "Отчет сейчас открыт. Запись невозможна.", vbExclamation
                End
            End If
            Kill (strFileName)
        Else
            End
        End If
    End If
    
    ' Создание отчета Word
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    
    With objWord
        .Visible = True
        .Activate
    End With

    With objWord.Selection
        .Paragraphs.SpaceAfter = 8
        .Paragraphs.SpaceAfterAuto = False
        .Paragraphs.LineSpacingRule = wdLineSpaceMultiple
        .Paragraphs.LineSpacing = LinesToPoints(1.08)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Name = "Calibri"
        .Font.Size = 14
        .TypeText "Отчет по загрузке" & vbCrLf
    End With
    
    With objWord.Selection
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 11
        .TypeText vbCrLf
    End With
    
    ' Создание таблицы
    Set objTable = objDoc.Tables.Add(objWord.Selection.Range, 1, 2)
    
    objTable.Borders.Enable = True
    
    ' добавление шапки таблицы
    Set objRow = objTable.Rows(objTable.Rows.count)
    
    objRow.Shading.BackgroundPatternColor = wdColorGray50
    
    objDoc.Tables(1).Rows(1).Select
    With objWord.Selection
        .Paragraphs.SpaceBefore = 0
        .Paragraphs.SpaceAfter = 0
        .Paragraphs.SpaceAfterAuto = False
        .Paragraphs.LineSpacingRule = wdLineSpaceMultiple
        .Paragraphs.LineSpacing = LinesToPoints(1#)
        .Cells(1).Width = CentimetersToPoints(8.24)
        .Cells(2).Width = CentimetersToPoints(8.24)
    End With
    
    objWord.Selection.HomeKey Unit:=wdStory
    
    With objRow.Cells(1).Range
        .Font.Color = wdColorWhite
        .Bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Size = 11
        .Text = "Отдел"
    End With
    
    With objRow.Cells(2).Range
        .Font.Color = wdColorWhite
        .Bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Size = 11
        .Text = "Количество задач"
    End With

    Dim Division As Integer
    
    Division = -1
    
    For j = 0 To IdCount - 1
        If Division <> Divisions(j) Then
            Division = Divisions(j)
            
            ' добавление записи отдела и суммы его заданий
            objTable.Rows.Add
            
            Set objRow = objTable.Rows(objTable.Rows.count)
            
            objRow.Shading.BackgroundPatternColor = wdColorGray15
            
            With objRow.Cells(1).Range
                .Font.Color = wdAuto
                .Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                .Text = "Отдел " & CStr(Division)
            End With
            
            With objRow.Cells(2).Range
                .Font.Color = wdAuto
                .Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Text = CStr(DivisionTasks(j))
            End With
            
            ' добавление сотрудника вышеперечисленного отдела и суммы его заданий
            For i = 0 To IdCount - 1
                If Divisions(i) = Division Then
                    objTable.Rows.Add
                    
                    Set objRow = objTable.Rows(objTable.Rows.count)
                    
                    objRow.Shading.BackgroundPatternColor = wdColorWhite
                    
                    With objRow.Cells(1).Range
                        .Font.ColorIndex = wdAuto
                        .Bold = False
                        .ParagraphFormat.Alignment = wdAlignParagraphLeft
                        .Text = Fios(i)
                    End With
                    
                    With objRow.Cells(2).Range
                        .Font.ColorIndex = wdAuto
                        .Bold = False
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .Text = IdTasks(i)
                    End With
                End If
            Next i
        End If
    Next j
    
    ' Сохранить отчет Word на жестком диске
    objDoc.SaveAs (strFileName)
End Sub

Public Sub SortingOf4Arrays(vArray As Variant, vArray1 As Variant, vArray2 As Variant, _
                      vArray3 As Variant, iArrLow As Integer, iArrHigh As Integer)
    Dim vTmp As Variant
    Dim vTmp1 As Variant
    Dim vTmp2 As Variant
    Dim vTmp3 As Variant
    Dim q As Integer
    
    Do
       q = 0
       For i = iArrLow To iArrHigh - 1
           If vArray(i + 1) > vArray(i) Then
              q = -1
              vTmp = vArray(i)
              vArray(i) = vArray(i + 1)
              vArray(i + 1) = vTmp
              
              vTmp1 = vArray1(i)
              vArray1(i) = vArray1(i + 1)
              vArray1(i + 1) = vTmp1
              
              vTmp2 = vArray2(i)
              vArray2(i) = vArray2(i + 1)
              vArray2(i + 1) = vTmp2
              
              vTmp3 = vArray3(i)
              vArray3(i) = vArray3(i + 1)
              vArray3(i + 1) = vTmp3
           End If
       Next i
       If q = 0 Then Exit Do
    Loop
End Sub

Public Sub SortingOf5Arrays(vArray As Variant, vArray1 As Variant, vArray2 As Variant, _
                      vArray3 As Variant, vArray4 As Variant, iArrLow As Integer, iArrHigh As Integer)
    Dim vTmp As Variant
    Dim vTmp1 As Variant
    Dim vTmp2 As Variant
    Dim vTmp3 As Variant
    Dim vTmp4 As Variant
    Dim q As Integer
    
    Do
       q = 0
       For i = iArrLow To iArrHigh - 1
           If vArray(i + 1) > vArray(i) Then
              q = -1
              vTmp = vArray(i)
              vArray(i) = vArray(i + 1)
              vArray(i + 1) = vTmp
              
              vTmp1 = vArray1(i)
              vArray1(i) = vArray1(i + 1)
              vArray1(i + 1) = vTmp1
              
              vTmp2 = vArray2(i)
              vArray2(i) = vArray2(i + 1)
              vArray2(i + 1) = vTmp2
              
              vTmp3 = vArray3(i)
              vArray3(i) = vArray3(i + 1)
              vArray3(i + 1) = vTmp3
              
              vTmp4 = vArray4(i)
              vArray4(i) = vArray4(i + 1)
              vArray4(i + 1) = vTmp4
           End If
       Next i
       If q = 0 Then Exit Do
    Loop
End Sub

Public Function GetSumOfIdTasks(Id As Variant)
    Dim Sum As Integer
    
    Sum = 0
    
    Set IdCells = Sheets("Задачи").Range("B2:B" & CStr(Rows.count))
    
    For Each Cell In IdCells
        If Cell.Value <> "" Then
            If Trim(Cell.Value) = Id Then
                Sum = Sum + 1
            End If
        Else
            Exit For
        End If
    Next
    
    GetSumOfIdTasks = Sum
End Function

Public Function GetSumOfDivisionTasks(DivisionNumber As Integer, vDivisions As Variant, vIdTasks As Variant)
    Dim Sum As Integer
    Dim i As Integer
    
    Sum = 0
    
    i = 0
    For Each Division In vDivisions
        If DivisionNumber = Division Then
            Sum = Sum + vIdTasks(i)
        End If
        i = i + 1
    Next
    
    GetSumOfDivisionTasks = Sum
End Function

Function IsFileLocked(sFile As String) As Boolean
    On Error Resume Next
    Open sFile For Binary Access Read Write Lock Read Write As #1
    Close #1
    
    If Err.Number <> 0 Then
        IsFileLocked = True
        Err.Clear
    Else
        IsFileLocked = False
        Err.Clear
    End If
End Function
