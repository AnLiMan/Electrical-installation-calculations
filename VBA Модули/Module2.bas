Attribute VB_Name = "Module2"
' Функция расчета сопротивления для выбора сечения провода

Sub CalculateMaxResistance()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim voltageDropCell As Range, voltageCell As Range, currentCell As Range, materialCell As Range
    Dim lengthCell As Range, temperatureCell As Range
    Dim resultResistanceCell As Range, resultResistanceCellTemp As Range, resultSectionCell As Range
    Dim resultFinalSectionCell As Range, resultVoltageDrop As Range, resultMaxCurrentCell As Range
    Dim voltageDrop As Double, voltage As Double, current As Double, length As Double
    Dim maxResistance As Double, resistivity As Double, tempCoeff As Double, temperature As Double
    Dim currentResistance As Double, calculatedSection As Double, finalSection As Double
    Dim calculateVoltageDrop As Double, maxCurrentForSection As Double
    Dim selectedMaterial As String
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные")
    
    ' Ячейки ввода
    Set voltageDropCell = wsCalc.Range("B7") ' Ячейка с допустимым падением напряжения
    Set voltageCell = wsCalc.Range("B6") ' Ячейка с напряжением сети
    Set currentCell = wsCalc.Range("B4") ' Ячейка с током
    Set materialCell = wsCalc.Range("B2") ' Ячейка с материалом
    Set lengthCell = wsCalc.Range("B3") ' Ячейка с длиной
    Set temperatureCell = wsCalc.Range("B5") ' Ячейка с температурой
    
    ' Ячейки вывода
    Set resultResistanceCell = wsCalc.Range("B13")
    Set resultResistanceCellTemp = wsCalc.Range("B14")
    Set resultSectionCell = wsCalc.Range("B15")
    Set resultFinalSectionCell = wsCalc.Range("B16")
    Set resultVoltageDrop = wsCalc.Range("B17")
    
    ' Проверка данных
    If Not IsNumeric(voltageDropCell.Value) Or Not IsNumeric(voltageCell.Value) Or _
       Not IsNumeric(currentCell.Value) Or Not IsNumeric(lengthCell.Value) Or _
       Not IsNumeric(temperatureCell.Value) Then
        MsgBox "Проверьте введённые значения: все они должны быть числами!", vbExclamation
        Exit Sub
    End If
    
    'Присваиваем переменным числовые значения
    voltageDrop = voltageDropCell.Value
    voltage = voltageCell.Value
    current = currentCell.Value
    length = lengthCell.Value
    temperature = temperatureCell.Value
    selectedMaterial = materialCell.Value
    
    'Проверка длины провода и тока на нулевое значение
    If current = 0 Or length = 0 Then
        MsgBox "Ток и длина проводника не могут быть нулевыми!", vbExclamation
        Exit Sub
    End If
    
    ' 1. Расчёт максимального сопротивления
    maxResistance = (voltageDrop * voltage) / current
    
    ' 2. Получение характеристик материала
    Dim resistRange As Range, tempRange As Range
    Set resistRange = wsData.Range("A2:B4")
    Set tempRange = wsData.Range("D2:E4")
    
    Dim i As Integer
    Dim materialFound As Boolean: materialFound = False
    
    For i = 1 To resistRange.Rows.count
        If Trim(resistRange.Cells(i, 1).Value) = Trim(selectedMaterial) Then
            resistivity = resistRange.Cells(i, 2).Value
            materialFound = True
            Exit For
        End If
    Next i
    
    If materialFound Then
        materialFound = False
        For i = 1 To tempRange.Rows.count
            If Trim(tempRange.Cells(i, 1).Value) = Trim(selectedMaterial) Then
                tempCoeff = tempRange.Cells(i, 2).Value
                materialFound = True
                Exit For
            End If
        Next i
    End If
    
    If Not materialFound Then
        MsgBox "Материал '" & selectedMaterial & "' не найден в таблицах!", vbExclamation
        Exit Sub
    End If
    
    ' 3. Расчёт сопротивления с температурной поправкой
    currentResistance = maxResistance / (1 + tempCoeff * (temperature - 20))
    
    ' 4. Расчёт сечения
    calculatedSection = resistivity * length / currentResistance
    
    ' 5. Подбор стандартного сечения
    Dim standardSections As Range
    Set standardSections = wsData.Range("A10:A30")
    finalSection = Application.WorksheetFunction.Max(standardSections)
    
    For i = 1 To standardSections.Rows.count
        If standardSections.Cells(i, 1).Value >= calculatedSection Then
            finalSection = standardSections.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    ' 6. Проверка максимального тока для выбранного сечения
    Dim currentTable As Range
    Dim sectionFound As Boolean: sectionFound = False
    maxCurrentForSection = 0
    
    Set currentTable = wsData.Range("F10:G29") ' Таблица сечение-ток (F - сечение, G - ток)
    
    For i = 1 To currentTable.Rows.count
        If currentTable.Cells(i, 1).Value = finalSection Then
            maxCurrentForSection = currentTable.Cells(i, 2).Value
            sectionFound = True
            Exit For
        End If
    Next i
    
    If Not sectionFound Then
        MsgBox "Для сечения " & finalSection & " мм кв. не найдено значение максимального тока!", vbExclamation
    End If
    
    ' 7. Вывод результатов
    resultResistanceCell.Value = maxResistance 'Сопротивление кабеля максмальное, Ом (раб. температура)
    resultResistanceCell.NumberFormat = "0.000"
    
    resultResistanceCellTemp.Value = currentResistance 'Сопротивление кабеля максмальное, Ом (раб. температура)
    resultResistanceCellTemp.NumberFormat = "0.000"
    
    resultSectionCell.Value = calculatedSection 'Расчетное сечение, мм?
    resultSectionCell.NumberFormat = "0.000"
    
    resultFinalSectionCell.Value = finalSection 'Сечение кабеля по ПУЭ, мм?
    resultFinalSectionCell.NumberFormat = "0.00"
    
    correctedResistance = resistivity * length / finalSection
    calculateVoltageDrop = current * correctedResistance 'Величина падения напряжения, В
    resultVoltageDrop.Value = calculateVoltageDrop
    resultVoltageDrop.NumberFormat = "0.000"
    
    ' Проверка на превышение максимального тока
    Dim warningMsg As String
    warningMsg = ""
    
    If current > maxCurrentForSection And maxCurrentForSection > 0 Then
        warningMsg = vbCrLf & "ВНИМАНИЕ! Ток нагрузки (" & current & " А) превышает максимальный для этого сечения (" & maxCurrentForSection & " А). Уменьшите нагрузку на проводник!"
    End If
    
    ' Информационное сообщение
    MsgBox "Расчёт завершён:" & vbCrLf & _
           "Расчётное сечение: " & Format(calculatedSection, "0.0000") & " мм кв." & vbCrLf & _
           "Стандартное сечение: " & finalSection & " мм кв." & vbCrLf & _
           "Макс. ток для сечения: " & maxCurrentForSection & " А" & warningMsg, _
           vbInformation
End Sub
