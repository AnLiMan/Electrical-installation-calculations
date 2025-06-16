Attribute VB_Name = "Module4"
'Функция подбора кабеля в зависимости от марки по расчетному сечению

Sub SelectCableSectionByType()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim cableTypeCell As Range, currentSectionCell As Range, resultSectionCell As Range
    Dim cableType As String, currentSection As Double
    Dim cableTypesRange As Range, sectionsTable As Range
    Dim targetColumn As Range, resultSection As Double
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет")
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные")
    
    ' Ячейки ввода/вывода
    Set cableTypeCell = wsCalc.Range("B19")          ' Ячейка с выпадающим списком марок кабеля
    Set currentSectionCell = wsCalc.Range("B15")     ' Текущее сечение (мм кв) для подбора
    Set resultSectionCell = wsCalc.Range("B20")      ' Результат: подобранное сечение
    
    ' Проверка ввода
    If Not IsNumeric(currentSectionCell.Value) Or currentSectionCell.Value <= 0 Then
        MsgBox "Введите корректное значение сечения (положительное число)!", vbExclamation
        Exit Sub
    End If
    
    cableType = cableTypeCell.Value
    currentSection = currentSectionCell.Value
    
    ' Поиск столбца с сечениями для выбранной марки кабеля
    Set cableTypesRange = wsData.Range("B9:D17")     ' Заголовки с марками кабелей в 1-й строке
    Set targetColumn = Nothing
    
    On Error Resume Next
    Set targetColumn = cableTypesRange.Find(What:=cableType, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If targetColumn Is Nothing Then
        MsgBox "Марка кабеля '" & cableType & "' не найдена!", vbExclamation
        Exit Sub
    End If
    
    ' Диапазон сечений для выбранной марки (данные начинаются со 2-й строки)
    Set sectionsTable = wsData.Range(wsData.Cells(2, targetColumn.Column), _
                        wsData.Cells(100, targetColumn.Column)).SpecialCells(xlCellTypeConstants)
    
    ' Поиск ближайшего большего сечения
    resultSection = FindClosestSection(currentSection, sectionsTable)
    
    ' Вывод результата
    resultSectionCell.Value = resultSection
    MsgBox "Для марки " & cableType & vbCrLf & _
           "Требуемое сечение: " & currentSection & " мм кв." & vbCrLf & _
           "Рекомендуемое сечение: " & resultSection & " мм кв.", vbInformation
End Sub

'Поиск ближайшего значения сечения кабеля в зависимости от выбранной марки

Function FindClosestSection(targetSection As Double, sectionsRange As Range) As Double
    Dim cell As Range
    Dim closest As Double
    
    closest = Application.WorksheetFunction.Max(sectionsRange) ' По умолчанию - максимальное
    
    For Each cell In sectionsRange
        If IsNumeric(cell.Value) Then
            If cell.Value >= targetSection And cell.Value < closest Then
                closest = cell.Value
            End If
        End If
    Next cell
    
    FindClosestSection = closest
End Function
