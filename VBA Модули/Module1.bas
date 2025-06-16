Attribute VB_Name = "Module1"
'Функция перерасчета AWG в квадратные миллиметры

Sub CalculateAWGtoSquareMM()
    Dim awgValue As Double
    Dim squareMM As Double
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim standardsRange As Range
    Dim resultCell As Range
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")             ' Лист с вводом AWG
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные") ' Лист со стандартными сечениями
    
    ' Ячейки
    awgValue = wsCalc.Range("E2").Value                       ' Ячейка с AWG на листе "Расчет"
    Set standardsRange = wsData.Range("A10:A29")               ' Диапазон стандартов на вспомогательном листе
    Set resultCalculateCell = wsCalc.Range("E3")               ' Ячейка для результата с расчетным значением
    Set resultCell = wsCalc.Range("E5")                       ' Ячейка для результата со стандартным значением
    
    ' Проверка, что введено число
    If Not IsNumeric(awgValue) Then
        MsgBox "Введите числовое значение AWG в ячейку E2!", vbExclamation
        Exit Sub
    End If
    
    ' Формула перевода AWG в мм кв.
    squareMM = 0.012668 * (92 ^ ((36 - awgValue) / 19.5))
    resultCalculateCell.Value = squareMM
    
    ' Поиск ближайшего стандартного значения
    Dim closestStandard As Double
    closestStandard = FindClosestStandard(squareMM, standardsRange)
    
    ' Вывод результата
    resultCell.Value = closestStandard
    MsgBox "Результат: AWG " & awgValue & " = " & closestStandard & " мм кв", vbInformation
End Sub

' Функция поиска ближайшего стандартного значения по ПУЭ

Function FindClosestStandard(targetValue As Double, standardsRange As Range) As Double
    Dim cell As Range
    Dim closest As Double
    Dim minDiff As Double
    Dim tolerance As Double
    
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")
    
    tolerance = wsCalc.Range("E4")
    minDiff = 1E+30
    closest = 0
    
    ' Сначала ищем ближайшее большее в пределах 5%
    For Each cell In standardsRange
        If IsNumeric(cell.Value) Then
            Dim currentVal As Double
            currentVal = cell.Value
            
            ' Если значение в пределах ±5% (или иного заданного значения) и >= целевому
            If currentVal >= targetValue * (1 - tolerance) Then
                If currentVal <= targetValue * (1 + tolerance) Then
                    If currentVal < closest Or closest = 0 Then
                        closest = currentVal
                    End If
                End If
            End If
        End If
    Next cell
    
    ' Если не нашли в пределах 5% (или иного заданного значения), берём ближайшее большее
    If closest = 0 Then
        For Each cell In standardsRange
            If IsNumeric(cell.Value) Then
                currentVal = cell.Value
                If currentVal >= targetValue Then
                    If currentVal < closest Or closest = 0 Then
                        closest = currentVal
                    End If
                End If
            End If
        Next cell
    End If
    
    ' Если все стандарты меньше, берём максимальный
    If closest = 0 Then
        closest = Application.WorksheetFunction.Max(standardsRange)
    End If
    
    FindClosestStandard = closest
End Function
