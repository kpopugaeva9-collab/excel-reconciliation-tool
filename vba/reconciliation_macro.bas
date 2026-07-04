Option Explicit

Sub ObnovitOtchet()

    Dim pathZayavki As Variant
    Dim pathNevyvoz As Variant
    
    Dim wbZ As Workbook
    Dim wbN As Workbook
    Dim wsZ As Worksheet
    Dim wsN As Worksheet
    Dim wsOut As Worksheet
    
    Dim colZ_ID As Long
    Dim colZ_Obem As Long
    Dim colZ_Address As Long
    Dim colZ_Waste As Long
    
    Dim colN_ID As Long
    Dim colN_Obem As Long
    
    Dim rowZ_Header As Long
    Dim rowN_Header As Long
    Dim tmpRow As Long
    
    Dim lastRowZ As Long
    Dim lastRowN As Long
    Dim r As Long
    
    Dim planDict As Object
    Dim nevyvozDict As Object
    Dim wasteDict As Object
    Dim writtenDict As Object
    
    Dim id As String
    Dim plan As Double
    Dim nevyvoz As Double
    Dim fact As Double
    
    Dim colPlan As Long
    Dim colNevyvoz As Long
    Dim colFact As Long
    Dim colStatus As Long
    Dim colComment As Long
    
    Dim statusText As String
    Dim commentText As String
    Dim addressText As String
    Dim wasteText As String
    Dim hasNevyvoz As Boolean
    
    Dim outRow As Long
    Dim outPath As String
    Dim oldCalc As XlCalculation
    Dim stepName As String
    
    Dim errNumber As Long
    Dim errDescription As String
    
    Dim wasteTypes As Variant
    Dim statusTypes As Variant
    Dim startCol As Long
    Dim startRow As Long
    Dim summaryLastCol As Long
    Dim summaryTotalRow As Long
    Dim i As Long
    Dim j As Long
    Dim rr As Long
    Dim cnt As Long
    Dim totalByStatus As Long
    Dim totalByWaste As Long
    
    On Error GoTo Fail
    
    stepName = "Выбор файла заявок"
    
    pathZayavki = Application.GetOpenFilename( _
        "Excel files (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm", _
        , _
        "Выберите файл ЗАЯВОК" _
    )
    
    If VarType(pathZayavki) = vbBoolean Then Exit Sub
    
    stepName = "Выбор файла отчёта по невывозу"
    
    pathNevyvoz = Application.GetOpenFilename( _
        "Excel files (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm", _
        , _
        "Выберите файл ОТЧЁТА ПО НЕВЫВОЗУ" _
    )
    
    If VarType(pathNevyvoz) = vbBoolean Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    stepName = "Открытие файла заявок"
    
    Set wbZ = Application.Workbooks.Open( _
        fileName:=CStr(pathZayavki), _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        IgnoreReadOnlyRecommended:=True _
    )
    
    If wbZ Is Nothing Then
        Err.Raise vbObjectError + 3001, , "Не открылся файл заявок. Проверьте, что он закрыт."
    End If
    
    stepName = "Открытие файла отчёта по невывозу"
    
    Set wbN = Application.Workbooks.Open( _
        fileName:=CStr(pathNevyvoz), _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True _
    )
    
    If wbN Is Nothing Then
        Err.Raise vbObjectError + 3002, , "Не открылся файл отчёта по невывозу. Проверьте, что он закрыт."
    End If
    
    stepName = "Поиск первого листа"
    
    Set wsZ = wbZ.Worksheets(1)
    Set wsN = wbN.Worksheets(1)
    
    stepName = "Поиск столбцов в заявках"
    
    colZ_ID = FindHeader(wsZ, "ID площадки", rowZ_Header)
    colZ_Obem = FindHeader(wsZ, "Объем, м3", tmpRow)
    
    If tmpRow <> rowZ_Header Then
        Err.Raise vbObjectError + 1001, , "В файле заявок заголовки ID площадки и Объем, м3 находятся в разных строках."
    End If
    
    colZ_Address = FindHeader(wsZ, "Адрес площадки", tmpRow)
    
    If tmpRow <> rowZ_Header Then
        Err.Raise vbObjectError + 1003, , "В файле заявок заголовок Адрес площадки находится в другой строке."
    End If
    
    colZ_Waste = FindHeader(wsZ, "Вид отходов", tmpRow)
    
    If tmpRow <> rowZ_Header Then
        Err.Raise vbObjectError + 1004, , "В файле заявок заголовок Вид отходов находится в другой строке."
    End If
    
    stepName = "Поиск столбцов в отчёте по невывозу"
    
    colN_ID = FindHeader(wsN, "ID МНО", rowN_Header)
    colN_Obem = FindHeader(wsN, "Объем невывоза, м3", tmpRow)
    
    If tmpRow <> rowN_Header Then
        Err.Raise vbObjectError + 1002, , "В файле невывоза заголовки ID МНО и Объем невывоза, м3 находятся в разных строках."
    End If
    
    Set planDict = CreateObject("Scripting.Dictionary")
    Set nevyvozDict = CreateObject("Scripting.Dictionary")
    Set wasteDict = CreateObject("Scripting.Dictionary")
    
    stepName = "Сбор суммы заявок"
    
    lastRowZ = LastUsedRow(wsZ)
    
    For r = rowZ_Header + 1 To lastRowZ
        
        id = NormalizeID(wsZ.Cells(r, colZ_ID).Value)
        
        If Len(id) > 0 Then
            
            If Not planDict.Exists(id) Then planDict.Add id, 0#
            planDict(id) = planDict(id) + ToNumber(wsZ.Cells(r, colZ_Obem).Value)
            
            wasteText = NormalizeText(wsZ.Cells(r, colZ_Waste).Value)
            
            If Len(wasteText) > 0 Then
                If Not wasteDict.Exists(id) Then
                    wasteDict.Add id, wasteText
                Else
                    wasteDict(id) = AddUniqueText(CStr(wasteDict(id)), wasteText)
                End If
            End If
            
        End If
        
    Next r
    
    stepName = "Сбор суммы невывоза"
    
    lastRowN = LastUsedRow(wsN)
    
    For r = rowN_Header + 1 To lastRowN
        
        id = NormalizeID(wsN.Cells(r, colN_ID).Value)
        
        If Len(id) > 0 Then
            If Not nevyvozDict.Exists(id) Then nevyvozDict.Add id, 0#
            nevyvozDict(id) = nevyvozDict(id) + ToNumber(wsN.Cells(r, colN_Obem).Value)
        End If
        
    Next r
    
    stepName = "Добавление итоговых столбцов в исходные заявки"
    
    colPlan = GetOrCreateHeader(wsZ, rowZ_Header, "План по площадке, м3")
    colNevyvoz = GetOrCreateHeader(wsZ, rowZ_Header, "Невывоз по площадке, м3")
    colFact = GetOrCreateHeader(wsZ, rowZ_Header, "Вывезено, м3")
    colStatus = GetOrCreateHeader(wsZ, rowZ_Header, "статус факт")
    colComment = GetOrCreateHeader(wsZ, rowZ_Header, "Комментарий сверки")
    
    stepName = "Проставление результата в исходные заявки"
    
    For r = rowZ_Header + 1 To lastRowZ
        
        id = NormalizeID(wsZ.Cells(r, colZ_ID).Value)
        
        If Len(id) = 0 Then
        
            wsZ.Cells(r, colPlan).Value = ""
            wsZ.Cells(r, colNevyvoz).Value = ""
            wsZ.Cells(r, colFact).Value = ""
            wsZ.Cells(r, colStatus).Value = "проверить"
            wsZ.Cells(r, colComment).Value = "Пустой ID площадки"
            
        Else
            
            plan = 0#
            nevyvoz = 0#
            hasNevyvoz = nevyvozDict.Exists(id)
            
            If planDict.Exists(id) Then plan = planDict(id)
            If hasNevyvoz Then nevyvoz = nevyvozDict(id)
            
            fact = plan - nevyvoz
            
            If Not hasNevyvoz Then
                statusText = "вывоз"
                commentText = "ID не найден в отчёте по невывозу"
                
            ElseIf Abs(plan - nevyvoz) < 0.000001 Then
                statusText = "невывоз"
                commentText = "Невывоз равен плану"
                
            ElseIf nevyvoz < plan Then
                statusText = "частичный вывоз"
                commentText = "Невывоз " & Format(nevyvoz, "0.###") & " из " & Format(plan, "0.###") & " м3"
                
            Else
                statusText = "невывоз"
                commentText = "Невывоз больше плана — проверить"
            End If
            
            wsZ.Cells(r, colPlan).Value = Round(plan, 3)
            wsZ.Cells(r, colNevyvoz).Value = Round(nevyvoz, 3)
            
            If fact < 0 Then
                wsZ.Cells(r, colFact).Value = 0
            Else
                wsZ.Cells(r, colFact).Value = Round(fact, 3)
            End If
            
            wsZ.Cells(r, colStatus).Value = statusText
            wsZ.Cells(r, colComment).Value = commentText
            
        End If
        
    Next r
    
    wsZ.Columns(colPlan).NumberFormat = "0.000"
    wsZ.Columns(colNevyvoz).NumberFormat = "0.000"
    wsZ.Columns(colFact).NumberFormat = "0.000"
    
    stepName = "Создание листа Итог по площадкам"
    
    On Error Resume Next
    Set wsOut = wbZ.Worksheets("Итог по площадкам")
    On Error GoTo Fail
    
    If wsOut Is Nothing Then
        Set wsOut = wbZ.Worksheets.Add(After:=wsZ)
        wsOut.Name = "Итог по площадкам"
    Else
        wsOut.Cells.Clear
    End If
    
    wsOut.Cells(1, 1).Value = "ID площадки"
    wsOut.Cells(1, 2).Value = "Адрес площадки"
    wsOut.Cells(1, 3).Value = "Вид отходов"
    wsOut.Cells(1, 4).Value = "План по площадке, м3"
    wsOut.Cells(1, 5).Value = "Невывоз по площадке, м3"
    wsOut.Cells(1, 6).Value = "Вывезено, м3"
    wsOut.Cells(1, 7).Value = "статус факт"
    wsOut.Cells(1, 8).Value = "Комментарий сверки"
    
    Set writtenDict = CreateObject("Scripting.Dictionary")
    outRow = 2
    
    For r = rowZ_Header + 1 To lastRowZ
        
        id = NormalizeID(wsZ.Cells(r, colZ_ID).Value)
        
        If Len(id) > 0 Then
            
            If Not writtenDict.Exists(id) Then
                
                writtenDict.Add id, True
                
                plan = 0#
                nevyvoz = 0#
                hasNevyvoz = nevyvozDict.Exists(id)
                
                If planDict.Exists(id) Then plan = planDict(id)
                
                If hasNevyvoz Then
                    nevyvoz = nevyvozDict(id)
                Else
                    nevyvoz = 0#
                End If
                
                fact = plan - nevyvoz
                
                If Not hasNevyvoz Then
                    statusText = "вывоз"
                    commentText = "ID не найден в отчёте по невывозу"
                    
                ElseIf Abs(plan - nevyvoz) < 0.000001 Then
                    statusText = "невывоз"
                    commentText = "Невывоз равен плану"
                    
                ElseIf nevyvoz < plan Then
                    statusText = "частичный вывоз"
                    commentText = "Невывоз " & Format(nevyvoz, "0.###") & " из " & Format(plan, "0.###") & " м3"
                    
                Else
                    statusText = "невывоз"
                    commentText = "Невывоз больше плана — проверить"
                End If
                
                addressText = CStr(wsZ.Cells(r, colZ_Address).Value)
                
                If wasteDict.Exists(id) Then
                    wasteText = CStr(wasteDict(id))
                Else
                    wasteText = ""
                End If
                
                wsOut.Cells(outRow, 1).Value = id
                wsOut.Cells(outRow, 2).Value = addressText
                wsOut.Cells(outRow, 3).Value = wasteText
                wsOut.Cells(outRow, 4).Value = Round(plan, 3)
                wsOut.Cells(outRow, 5).Value = Round(nevyvoz, 3)
                
                If fact < 0 Then
                    wsOut.Cells(outRow, 6).Value = 0
                Else
                    wsOut.Cells(outRow, 6).Value = Round(fact, 3)
                End If
                
                wsOut.Cells(outRow, 7).Value = statusText
                wsOut.Cells(outRow, 8).Value = commentText
                
                outRow = outRow + 1
                
            End If
            
        End If
        
    Next r
    
    wsOut.Columns(4).NumberFormat = "0.000"
    wsOut.Columns(5).NumberFormat = "0.000"
    wsOut.Columns(6).NumberFormat = "0.000"
    
    wsOut.Rows(1).Font.Bold = True
    wsOut.Range("A1:H1").AutoFilter
    
    stepName = "Создание сводки по видам отходов"
    
    wasteTypes = Array( _
        "Смешанные отходы", _
        "Раздельно собранные отходы (Вторсырье)", _
        "Крупногабаритные отходы", _
        "Спил, смет" _
    )
    
    statusTypes = Array( _
        "вывоз", _
        "частичный вывоз", _
        "невывоз" _
    )
    
    startCol = 10
    startRow = 1
    summaryLastCol = startCol + 1 + UBound(wasteTypes) + 1
    summaryTotalRow = startRow + 2 + UBound(statusTypes) + 1
    
    wsOut.Cells(startRow, startCol).Value = "Сводка: количество площадок"
    
    wsOut.Cells(startRow + 1, startCol).Value = "статус факт"
    
    For j = 0 To UBound(wasteTypes)
        wsOut.Cells(startRow + 1, startCol + 1 + j).Value = wasteTypes(j)
    Next j
    
    wsOut.Cells(startRow + 1, summaryLastCol).Value = "Всего площадок"
    
    For i = 0 To UBound(statusTypes)
        
        wsOut.Cells(startRow + 2 + i, startCol).Value = statusTypes(i)
        
        totalByStatus = 0
        
        For rr = 2 To outRow - 1
            If NormalizeText(wsOut.Cells(rr, 7).Value) = NormalizeText(statusTypes(i)) Then
                totalByStatus = totalByStatus + 1
            End If
        Next rr
        
        For j = 0 To UBound(wasteTypes)
            
            cnt = 0
            
            For rr = 2 To outRow - 1
                
                If NormalizeText(wsOut.Cells(rr, 7).Value) = NormalizeText(statusTypes(i)) Then
                    If ContainsTextInList(CStr(wsOut.Cells(rr, 3).Value), CStr(wasteTypes(j))) Then
                        cnt = cnt + 1
                    End If
                End If
                
            Next rr
            
            wsOut.Cells(startRow + 2 + i, startCol + 1 + j).Value = cnt
            
        Next j
        
        wsOut.Cells(startRow + 2 + i, summaryLastCol).Value = totalByStatus
        
    Next i
    
    wsOut.Cells(summaryTotalRow, startCol).Value = "Итого по виду"
    
    For j = 0 To UBound(wasteTypes)
        
        totalByWaste = 0
        
        For rr = 2 To outRow - 1
            If ContainsTextInList(CStr(wsOut.Cells(rr, 3).Value), CStr(wasteTypes(j))) Then
                totalByWaste = totalByWaste + 1
            End If
        Next rr
        
        wsOut.Cells(summaryTotalRow, startCol + 1 + j).Value = totalByWaste
        
    Next j
    
    wsOut.Cells(summaryTotalRow, summaryLastCol).Value = outRow - 2
    
    wsOut.Range(wsOut.Cells(startRow + 1, startCol), wsOut.Cells(summaryTotalRow, summaryLastCol)).Borders.LineStyle = xlContinuous
    wsOut.Range(wsOut.Cells(startRow + 1, startCol), wsOut.Cells(startRow + 1, summaryLastCol)).Font.Bold = True
    wsOut.Range(wsOut.Cells(summaryTotalRow, startCol), wsOut.Cells(summaryTotalRow, summaryLastCol)).Font.Bold = True
    
    wsOut.Cells(summaryTotalRow + 2, startCol).Value = "Примечание: если у одной площадки несколько видов отходов, она считается в каждом соответствующем виде."
    
    wsOut.Columns.AutoFit
    
    stepName = "Сохранение результата"
    
    outPath = BuildOutputPath(CStr(pathZayavki))
    
    wbZ.SaveAs fileName:=outPath, FileFormat:=xlOpenXMLWorkbook
    
    wbN.Close SaveChanges:=False
    
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Готово!" & vbCrLf & vbCrLf & _
           "Файл сохранён:" & vbCrLf & outPath, vbInformation

    Exit Sub

Fail:
    errNumber = Err.Number
    errDescription = Err.Description
    
    On Error Resume Next
    
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If Not wbN Is Nothing Then wbN.Close SaveChanges:=False
    
    MsgBox "Ошибка " & errNumber & " на шаге:" & vbCrLf & _
           stepName & vbCrLf & vbCrLf & _
           errDescription, vbCritical

End Sub


Private Function FindHeader(ByVal ws As Worksheet, ByVal headerName As String, ByRef headerRow As Long) As Long
    Dim cell As Range
    Dim target As String
    Dim current As String
    
    target = NormalizeHeader(headerName)
    
    For Each cell In ws.UsedRange.Cells
        
        current = NormalizeHeader(CStr(cell.Value))
        
        If current = target Then
            headerRow = cell.Row
            FindHeader = cell.Column
            Exit Function
        End If
        
    Next cell
    
    Err.Raise vbObjectError + 2002, , _
        "Не найден столбец """ & headerName & """ на листе """ & ws.Name & """."
End Function


Private Function GetOrCreateHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim target As String
    
    target = NormalizeHeader(headerName)
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        If NormalizeHeader(CStr(ws.Cells(headerRow, c).Value)) = target Then
            GetOrCreateHeader = c
            Exit Function
        End If
    Next c
    
    GetOrCreateHeader = lastCol + 1
    ws.Cells(headerRow, GetOrCreateHeader).Value = headerName
End Function


Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    
    Set foundCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )
    
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function


Private Function NormalizeHeader(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, Chr(10), "")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, "ё", "е")
    s = Replace(s, "?", "3")
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    
    NormalizeHeader = s
End Function


Private Function NormalizeID(ByVal v As Variant) As String
    Dim s As String
    
    If IsError(v) Then
        NormalizeID = ""
        Exit Function
    End If
    
    If IsEmpty(v) Then
        NormalizeID = ""
        Exit Function
    End If
    
    If IsNumeric(v) Then
        NormalizeID = Format$(CDbl(v), "0")
    Else
        s = CStr(v)
        s = Replace(s, Chr(160), " ")
        s = Trim$(s)
        NormalizeID = s
    End If
End Function


Private Function NormalizeText(ByVal v As Variant) As String
    Dim s As String
    
    If IsError(v) Then
        NormalizeText = ""
        Exit Function
    End If
    
    If IsEmpty(v) Then
        NormalizeText = ""
        Exit Function
    End If
    
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Trim$(s)
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeText = s
End Function


Private Function ToNumber(ByVal v As Variant) As Double
    Dim s As String
    
    If IsError(v) Then
        ToNumber = 0#
        Exit Function
    End If
    
    If IsEmpty(v) Then
        ToNumber = 0#
        Exit Function
    End If
    
    If IsNumeric(v) Then
        ToNumber = CDbl(v)
        Exit Function
    End If
    
    s = CStr(v)
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, "м3", "")
    s = Replace(s, "м?", "")
    s = Replace(s, ",", ".")
    
    ToNumber = Val(s)
End Function


Private Function AddUniqueText(ByVal existingText As String, ByVal newText As String) As String
    Dim parts() As String
    Dim i As Long
    
    existingText = NormalizeText(existingText)
    newText = NormalizeText(newText)
    
    If Len(newText) = 0 Then
        AddUniqueText = existingText
        Exit Function
    End If
    
    If Len(existingText) = 0 Then
        AddUniqueText = newText
        Exit Function
    End If
    
    parts = Split(existingText, ";")
    
    For i = LBound(parts) To UBound(parts)
        If LCase$(NormalizeText(parts(i))) = LCase$(newText) Then
            AddUniqueText = existingText
            Exit Function
        End If
    Next i
    
    AddUniqueText = existingText & "; " & newText
End Function


Private Function ContainsTextInList(ByVal listText As String, ByVal targetText As String) As Boolean
    Dim parts() As String
    Dim i As Long
    
    listText = NormalizeText(listText)
    targetText = NormalizeText(targetText)
    
    If Len(listText) = 0 Or Len(targetText) = 0 Then
        ContainsTextInList = False
        Exit Function
    End If
    
    parts = Split(listText, ";")
    
    For i = LBound(parts) To UBound(parts)
        If LCase$(NormalizeText(parts(i))) = LCase$(targetText) Then
            ContainsTextInList = True
            Exit Function
        End If
    Next i
    
    ContainsTextInList = False
End Function


Private Function BuildOutputPath(ByVal sourcePath As String) As String
    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim dotPos As Long
    
    folderPath = Left$(sourcePath, InStrRev(sourcePath, "\"))
    fileName = Mid$(sourcePath, InStrRev(sourcePath, "\") + 1)
    dotPos = InStrRev(fileName, ".")
    
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
    Else
        baseName = fileName
    End If
    
    BuildOutputPath = folderPath & baseName & "_со_статусом_факт_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
End Function

