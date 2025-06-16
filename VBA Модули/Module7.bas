Attribute VB_Name = "Module7"
'Выбор оптимального диаметра гофры, исходя из итогового значения диаметра

Sub FindOptimalGofraDiameter()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim searchValue As Range, resultCell As Range, extDiamCell As Range
    Dim diametersRange As Range, extDiametersRange As Range
    Dim selectedOption As String
    Dim targetDiameter As Double, foundDiameter As Double, foundExtDiameter As Double
    Dim i As Long
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет гофры")
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные")
    
    ' Проверяем значение в G10
    Set searchValue = wsCalc.Range("G10")
    selectedOption = Trim(searchValue.Value)
    
    ' Ячейки для результатов
    Set resultCell = wsCalc.Range("I7") ' Внутренний диаметр
    Set extDiamCell = wsCalc.Range("J7") ' Внешний диаметр
    
    ' Очищаем старые результаты
    resultCell.ClearContents
    extDiamCell.ClearContents
    
    ' Получаем целевой диаметр для поиска (из H7)
    If Not IsNumeric(wsCalc.Range("H7").Value) Then
        MsgBox "Сначала рассчитайте диаметр описанной окружности!", vbExclamation
        Exit Sub
    End If
    targetDiameter = wsCalc.Range("H7").Value
    
    ' Выбираем диапазоны в зависимости от варианта
    If selectedOption = "Да" Then
        Set diametersRange = wsData.Range("Q9:Q33").SpecialCells(xlCellTypeConstants)
        Set extDiametersRange = wsData.Range("R9:R33").SpecialCells(xlCellTypeConstants)
    ElseIf selectedOption = "Нет" Then
        Set diametersRange = wsData.Range("T9:T20").SpecialCells(xlCellTypeConstants)
        Set extDiametersRange = wsData.Range("U9:U20").SpecialCells(xlCellTypeConstants)
    Else
        MsgBox "Выберите 'Да' или 'Нет' в ячейке G10", vbExclamation
        Exit Sub
    End If
    
    ' Ищем ближайший больший внутренний диаметр
    foundDiameter = 0
    foundExtDiameter = 0
    
    For i = 1 To diametersRange.Cells.count
        If IsNumeric(diametersRange.Cells(i).Value) Then
            If diametersRange.Cells(i).Value >= targetDiameter Then
                If foundDiameter = 0 Or diametersRange.Cells(i).Value < foundDiameter Then
                    foundDiameter = diametersRange.Cells(i).Value
                    foundExtDiameter = extDiametersRange.Cells(i).Value
                End If
            End If
        End If
    Next i
    
    ' Если не нашли - берем максимальный доступный
    If foundDiameter = 0 Then
        foundDiameter = Application.WorksheetFunction.Max(diametersRange)
        foundExtDiameter = extDiametersRange.Cells(diametersRange.Cells.count).Value
        MsgBox "Подходящий диаметр не найден. Использован максимальный доступный: " & foundDiameter, vbInformation
    End If
    
    ' Записываем результаты
    resultCell.Value = foundDiameter
    resultCell.NumberFormat = "0.00"
    
    extDiamCell.Value = foundExtDiameter
    extDiamCell.NumberFormat = "0.00"
    
    MsgBox "Подобран внутренний диаметр гофры: " & foundDiameter & vbCrLf & _
           "Соответствующий внешний диаметр: " & foundExtDiameter, vbInformation
End Sub
