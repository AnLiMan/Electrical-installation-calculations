Attribute VB_Name = "Module3"
'Функция перерасчета квадратных миллиметров в AWG

Sub CalculateSquareMMtoAWG()
    Dim squareMM As Double
    Dim awgValue As Double
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim awgStandardsRange As Range
    Dim resultCalculateCell As Range, resultCell As Range
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные")
    
    ' Ячейки
    squareMM = wsCalc.Range("H2").Value                  ' Ячейка с мм кв на листе "Расчет"
    Set awgStandardsRange = wsData.Range("A33:A48")       ' Диапазон AWG стандартов
    Set resultCalculateCell = wsCalc.Range("H3")          ' Расчетное значение AWG
    Set resultCell = wsCalc.Range("H5")                  ' Стандартное значение AWG
    
    ' Проверка ввода
    If Not IsNumeric(squareMM) Or squareMM <= 0 Then
        MsgBox "Введите числовое значение сечения!", vbExclamation
        Exit Sub
    End If
    
    ' Формула перевода мм кв. в AWG
    awgValue = 36 - (19.5 * Log(squareMM / 0.012668) / Log(92))
    resultCalculateCell.Value = awgValue
    
    ' Поиск ближайшего стандартного AWG
    Dim closestAWG As Double
    closestAWG = FindClosestAWG(awgValue, awgStandardsRange)
    
    ' Вывод результата
    resultCell.Value = closestAWG
    MsgBox "Результат: " & squareMM & " мм кв. = AWG " & closestAWG, vbInformation
End Sub

'Функция поиска ближайшего стандартного значения из ряда AWG

Function FindClosestAWG(targetAWG As Double, awgStandardsRange As Range) As Double
    Dim cell As Range
    Dim closest As Double
    Dim minDiff As Double
    Dim tolerance As Double
    
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")
    
    tolerance = wsCalc.Range("H4").Value  ' Допуск из ячейки H4
    minDiff = 1E+30
    closest = 0
    
    ' Ищем ближайшее значение AWG (меньшее = большее сечение)
    For Each cell In awgStandardsRange
        If IsNumeric(cell.Value) Then
            Dim currentAWG As Double
            currentAWG = cell.Value
            
            ' Для AWG: чем меньше номер, тем больше сечение
            ' Ищем значение в пределах допуска
            If currentAWG >= targetAWG * (1 - tolerance) And _
               currentAWG <= targetAWG * (1 + tolerance) Then
                ' Выбираем ближайшее меньшее (по модулю разницы)
                If Abs(currentAWG - targetAWG) < minDiff Then
                    minDiff = Abs(currentAWG - targetAWG)
                    closest = currentAWG
                End If
            End If
        End If
    Next cell
    
    ' Если не нашли в пределах допуска, берем ближайшее
    If closest = 0 Then
        minDiff = 1E+30
        For Each cell In awgStandardsRange
            If IsNumeric(cell.Value) Then
                currentAWG = cell.Value
                If Abs(currentAWG - targetAWG) < minDiff Then
                    minDiff = Abs(currentAWG - targetAWG)
                    closest = currentAWG
                End If
            End If
        Next cell
    End If
    
    FindClosestAWG = closest
End Function

