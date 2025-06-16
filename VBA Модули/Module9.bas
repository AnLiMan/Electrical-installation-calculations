Attribute VB_Name = "Module9"
'Расчет длины кабелей для жгутов

Sub PrepareCalculatePathMultiple()
    Dim wsExport As Worksheet, wsCalc As Worksheet, wsAux As Worksheet
    Set wsExport = ThisWorkbook.Sheets("Экспорт")
    Set wsCalc = ThisWorkbook.Sheets("Расчет жгута")
    Set wsAux = ThisWorkbook.Sheets("Вспомогательные данные")
    
    Dim rowExp As Long: rowExp = 2
    Dim harnessName As String, lastUsedRow As Long
    Dim graphRow As Long: graphRow = 52 ' Начальная строка построения графов
    
    Do While Trim(wsExport.Cells(rowExp, 1).Value) <> ""
        harnessName = Trim(CStr(wsExport.Cells(rowExp, 1).Value))
        
        Dim known As Boolean: known = False
        Dim i As Long
        
        ' Проверяем, уже есть ли такой жгут
        For i = 20 To 40
            If wsAux.Cells(i, 3).Value = harnessName Then
                known = True
                graphRow = wsAux.Cells(i, 4).Value
                Exit For
            End If
        Next i
        
        ' Если не найден, создаем структуру и записываем
        If Not known Then
            Dim searchRange As Range, cell As Range, baseCell As Range
            Set searchRange = wsCalc.Range("D5:AD100")
            Set baseCell = Nothing
            For Each cell In searchRange
                If Trim(CStr(cell.Value)) = harnessName Then
                    Set baseCell = cell
                    Exit For
                End If
            Next cell
            
            If baseCell Is Nothing Then
                wsExport.Cells(rowExp, 4).Value = "Жгут не найден"
                rowExp = rowExp + 1
                GoTo NextRow
            End If
            
            ' Координаты области
            Dim r0 As Long, c0 As Long
            r0 = baseCell.Row
            c0 = baseCell.Column
            
            Dim rStart As Long, rEnd As Long, cEnd As Long
            rStart = Application.Max(5, r0 - 7)
            rEnd = Application.Min(100, r0 + 7)
            cEnd = Application.Min(wsCalc.Columns.count, c0 + 50)
            
            ' Копируем структуру
            Dim r As Long, c As Long, rowIndex As Long
            rowIndex = graphRow
            For r = rStart To rEnd
                Dim colIndex As Long
                colIndex = 1
                For c = c0 To cEnd
                    Dim val As Variant
                    val = wsCalc.Cells(r, c).Value
                    If Not IsEmpty(val) Then
                        wsAux.Cells(rowIndex, colIndex).Value = val
                    End If
                    colIndex = colIndex + 1
                Next c
                rowIndex = rowIndex + 1
            Next r
            
            ' Сохраняем информацию о жгуте
            For i = 20 To 40
                If wsAux.Cells(i, 3).Value = "" Then
                    wsAux.Cells(i, 3).Value = harnessName
                    wsAux.Cells(i, 4).Value = graphRow
                    Exit For
                End If
            Next i
        End If

        ' Чтение начальной и конечной точки
        Dim startNode As String, endNode As String
        startNode = Trim(CStr(wsExport.Cells(rowExp, 2).Value))
        endNode = Trim(CStr(wsExport.Cells(rowExp, 3).Value))
        
        Dim Path As Collection
        Set Path = FindPathInList(wsAux, startNode, endNode, graphRow)
        
        If Path Is Nothing Then
            wsExport.Cells(rowExp, 4).Value = "Путь не найден"
        Else
            Dim s As String, totalWeight As Double, nodeVal
            totalWeight = 0
            For Each nodeVal In Path
                s = s & nodeVal & " > "
                If IsNumeric(nodeVal) Then totalWeight = totalWeight + CDbl(nodeVal)
            Next nodeVal
            s = Left(s, Len(s) - 3)
            wsExport.Cells(rowExp, 4).Value = totalWeight
            wsExport.Cells(rowExp, 5).Value = s ' Путь
        End If
        
        graphRow = graphRow + 15 ' Сдвиг на 15 строк для следующего нового графа
NextRow:
        rowExp = rowExp + 1
    Loop
End Sub

Function FindPathInList(ws As Worksheet, startVal As String, endVal As String, baseRow As Long) As Collection
    Dim r As Long, c As Long
    Dim nodeDict As Object: Set nodeDict = CreateObject("Scripting.Dictionary")
    
    ' Сканируем таблицу и сохраняем значения с координатами
    For r = baseRow To baseRow + 30
        For c = 1 To 50
            Dim val As String
            val = Trim(CStr(ws.Cells(r, c).Value))
            If val <> "" Then
                nodeDict(r & "_" & c) = val
            End If
        Next c
    Next r
    
    ' Поиск стартовой позиции
    Dim found As Boolean: found = False
    Dim sr As Long, sc As Long
    Dim key As Variant
    For Each key In nodeDict.Keys
        If nodeDict(key) = startVal Then
            sr = CLng(Split(key, "_")(0))
            sc = CLng(Split(key, "_")(1))
            found = True
            Exit For
        End If
    Next key
    
    If Not found Then Exit Function
    
    ' Рекурсивный поиск
    Dim visited As Object: Set visited = CreateObject("Scripting.Dictionary")
    Dim Path As Collection: Set Path = New Collection
    If RecursiveSearch(nodeDict, sr, sc, endVal, visited, Path) Then
        Set FindPathInList = Path
    End If
End Function

Function RecursiveSearch(nodeDict As Object, r As Long, c As Long, _
                           endVal As String, visited As Object, _
                           Path As Collection) As Boolean
    Dim key As String
    key = r & "_" & c
    If visited.exists(key) Then Exit Function
    If Not nodeDict.exists(key) Then Exit Function
    
    visited.Add key, True
    Path.Add nodeDict(key)
    
    If nodeDict(key) = endVal Then
        RecursiveSearch = True
        Exit Function
    End If
    
    Dim i As Long
    Dim dr As Variant, dc As Variant
    dr = Array(-1, 1, 0, 0)
    dc = Array(0, 0, -1, 1)
    
    For i = 0 To 3
        Dim nr As Long, nc As Long
        nr = r + dr(i)
        nc = c + dc(i)
        If RecursiveSearch(nodeDict, nr, nc, endVal, visited, Path) Then
            RecursiveSearch = True
            Exit Function
        End If
    Next i
    
    Path.Remove Path.count ' Backtrack
End Function

