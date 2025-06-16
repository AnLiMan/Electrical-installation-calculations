Attribute VB_Name = "Module5"
'Расчет сечения для сериии кабелей

Sub CalculateMultipleResistancesWithPUEandTemp()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim materialCell As Range, voltageCell As Range, voltageDropCell As Range, tempCell As Range
    Dim lengthsRange As Range, currentsRange As Range, resultsRange As Range, pueSectionsRange As Range
    Dim selectedMaterial As String, voltage As Double, voltageDrop As Double, temperature As Double
    Dim resistivity As Double, tempCoeff As Double
    Dim i As Long, lastRow As Long
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные")
    
    ' Базовые параметры
    Set materialCell = wsCalc.Range("B2") ' Ячейка с материалом
    Set voltageCell = wsCalc.Range("B6") ' Ячейка с напряжением сети
    Set voltageDropCell = wsCalc.Range("B7") ' Ячейка с допустимым уровнем падения напряжения
    Set tempCell = wsCalc.Range("B5") ' Ячейка с температурой
    
    ' Диапазон стандартных сечений по ПУЭ
    Set pueSectionsRange = wsData.Range("A10:A30")
    
    ' Определяем диапазоны данных
    lastRow = wsCalc.Cells(wsCalc.Rows.count, "B").End(xlUp).Row
    If lastRow < 26 Then lastRow = 26
    
    Set lengthsRange = wsCalc.Range("B26:B" & lastRow)
    Set currentsRange = wsCalc.Range("C26:C" & lastRow)
    Set resultsRange = wsCalc.Range("D26:D" & lastRow)
    
    ' Проверка базовых параметров
    If Not IsNumeric(voltageCell.Value) Or Not IsNumeric(voltageDropCell.Value) Or _
       Not IsNumeric(tempCell.Value) Then
        MsgBox "Проверьте напряжение, допустимые потери и температуру в B5-B7!", vbExclamation
        Exit Sub
    End If
    
    'Присваиваем переменным числовые значения
    voltage = voltageCell.Value
    voltageDrop = voltageDropCell.Value
    temperature = tempCell.Value
    selectedMaterial = materialCell.Value
    
    ' Получаем характеристики материала
    Dim resistRange As Range, tempCoeffRange As Range
    Set resistRange = wsData.Range("A2:B4")
    Set tempCoeffRange = wsData.Range("D2:E4")
    
    Dim materialFound As Boolean: materialFound = False
    
    ' Поиск удельного сопротивления
    For i = 1 To resistRange.Rows.count
        If Trim(resistRange.Cells(i, 1).Value) = Trim(selectedMaterial) Then
            resistivity = resistRange.Cells(i, 2).Value
            materialFound = True
            Exit For
        End If
    Next i
    
    ' Поиск температурного коэффициента
    If materialFound Then
        materialFound = False
        For i = 1 To tempCoeffRange.Rows.count
            If Trim(tempCoeffRange.Cells(i, 1).Value) = Trim(selectedMaterial) Then
                tempCoeff = tempCoeffRange.Cells(i, 2).Value
                materialFound = True
                Exit For
            End If
        Next i
    End If
    
    If Not materialFound Then
        MsgBox "Материал '" & selectedMaterial & "' не найден в таблицах!", vbExclamation
        Exit Sub
    End If
    
    ' Очищаем старые результаты
    resultsRange.ClearContents
    
    ' Обработка каждой строки с учетом температуры
    For i = 1 To lengthsRange.Rows.count
        If IsNumeric(lengthsRange.Cells(i, 1).Value) And _
           IsNumeric(currentsRange.Cells(i, 1).Value) And _
           currentsRange.Cells(i, 1).Value <> 0 Then
            
            Dim length As Double, current As Double
            Dim maxResistance As Double, currentResistance As Double
            Dim calculatedSection As Double, standardSection As Double
            
            length = lengthsRange.Cells(i, 1).Value
            current = currentsRange.Cells(i, 1).Value
            
            ' 1. Расчёт максимального сопротивления без учета температуры
            maxResistance = (voltageDrop * voltage) / current
            
            ' 2. Коррекция сопротивления с учетом температуры
            currentResistance = maxResistance / (1 + tempCoeff * (temperature - 20))
            
            ' 3. Расчёт сечения с учетом температурного сопротивления
            calculatedSection = resistivity * length / currentResistance
            
            ' 4. Подбор стандартного сечения по ПУЭ
            standardSection = FindClosestPueSection(calculatedSection, pueSectionsRange)
            
            ' Записываем результат
            resultsRange.Cells(i, 1).Value = standardSection
            resultsRange.Cells(i, 1).NumberFormat = "0.00"
        Else
            resultsRange.Cells(i, 1).Value = "-"
        End If
    Next i
    
    MsgBox "Расчет завершен для " & lengthsRange.Rows.count & " строк." & vbCrLf & _
           "Учтена рабочая температура: " & temperature & "°C" & vbCrLf & _
           "Температурный коэффициент: " & tempCoeff & " 1/°C", vbInformation
End Sub

'Функция поиска ближайшего значения сечения по ПУЭ

Function FindClosestPueSection(targetSection As Double, pueSectionsRange As Range) As Double
    Dim cell As Range
    Dim closestSection As Double
    
    closestSection = Application.WorksheetFunction.Max(pueSectionsRange)
    
    For Each cell In pueSectionsRange
        If IsNumeric(cell.Value) Then
            If cell.Value >= targetSection And cell.Value < closestSection Then
                closestSection = cell.Value
            End If
        End If
    Next cell
    
    FindClosestPueSection = closestSection
End Function

