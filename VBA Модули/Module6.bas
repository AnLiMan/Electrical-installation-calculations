Attribute VB_Name = "Module6"
'Расчет гофры для списка поводов с листа "Расчеты"

Sub CalculateWireParametersWithVisualization()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim sectionCol As Range, brandCol As Range
    Dim resultsDiaRange As Range, resultsAreaRange As Range
    Dim lastRow As Long, i As Long
    Dim dataTable As Range, headerRow As Range
    Dim summSection As Range, resultSummSection As Double
    Dim diameter As Double, area As Double
    Dim section As Double, brand As String
    
    ' Настройка листов
    Set wsCalc = ThisWorkbook.Worksheets("Расчет гофры")
    Set wsData = ThisWorkbook.Worksheets("Вспомогательные данные")
    
    'Расчет количества проводов
    lastRow = wsCalc.Cells(wsCalc.Rows.count, "B").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    
    Set sectionCol = wsCalc.Range("B2:B" & lastRow)
    Set brandCol = wsCalc.Range("C2:C" & lastRow)
    Set resultsDiaRange = wsCalc.Range("D2:D" & lastRow)
    Set resultsAreaRange = wsCalc.Range("E2:E" & lastRow)
    Set summSection = wsCalc.Range("F2")
    resultSummSection = 0

    
    Set dataTable = wsData.Range("K8").CurrentRegion
    Set headerRow = dataTable.Rows(1)

    resultsDiaRange.ClearContents
    resultsAreaRange.ClearContents

    On Error Resume Next
    Dim sh As Shape
    For Each sh In wsCalc.Shapes
        If sh.name Like "Wire_*" Or sh.name = "CircumscribedCircle" Then
            sh.Delete
        End If
    Next sh
    On Error GoTo 0

    For i = 1 To sectionCol.Rows.count
        If IsNumeric(sectionCol.Cells(i, 1).Value) And Not IsEmpty(brandCol.Cells(i, 1).Value) Then
            section = sectionCol.Cells(i, 1).Value
            brand = Trim(brandCol.Cells(i, 1).Value)
            diameter = FindDiameter(section, brand, dataTable, headerRow)
            
            If diameter > 0 Then
                area = WorksheetFunction.Pi * (diameter ^ 2) / 4
                resultsDiaRange.Cells(i, 1).Value = diameter
                resultsAreaRange.Cells(i, 1).Value = area
                resultSummSection = resultSummSection + area
                DrawWire wsCalc, diameter, i
            Else
                resultsDiaRange.Cells(i, 1).Value = "Нет данных"
                resultsAreaRange.Cells(i, 1).Value = "Нет данных"
            End If
        Else
            resultsDiaRange.Cells(i, 1).Value = "-"
            resultsAreaRange.Cells(i, 1).Value = "-"
        End If
    Next i
    
    summSection.Value = resultSummSection
    summSection.NumberFormat = "0.000"

    ' Вызов плотной отрисовки и описанной окружности
    Call RepackAndRedrawCircumscribedCircle(wsCalc, 20)
    
    MsgBox "Расчет завершен для " & sectionCol.Rows.count & " строк.", vbInformation
End Sub


Function FindDiameter(section As Double, brand As String, dataTable As Range, headerRow As Range) As Double
    Dim i As Long, j As Long
    Dim currentSection As Double
    Dim currentBrand As String
    
    For i = 2 To dataTable.Rows.count
        If IsNumeric(dataTable.Cells(i, 1).Value) Then
            currentSection = dataTable.Cells(i, 1).Value
            If Abs(currentSection - section) < 0.0001 Then
                For j = 2 To dataTable.Columns.count
                    currentBrand = Trim(headerRow.Cells(1, j).Value)
                    If StrComp(currentBrand, brand, vbTextCompare) = 0 Then
                        FindDiameter = dataTable.Cells(i, j).Value
                        Exit Function
                    End If
                Next j
            End If
        End If
    Next i
    
    FindDiameter = 0
End Function

Sub DrawWire(ws As Worksheet, diameter As Double, wireNumber As Long)
    Dim scaleFactor As Double: scaleFactor = 20 ' Масштабный коэффициент
    Dim leftPos As Double: leftPos = 300 + (wireNumber - 1) * 50
    Dim topPos As Double: topPos = 100
    Dim shapeName As String: shapeName = "Wire_" & wireNumber
    
    ' Удаляем старое изображение если есть
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
    
    ' Рисуем круг, представляющий провод
    With ws.Shapes.AddShape(msoShapeOval, leftPos, topPos, diameter * scaleFactor, diameter * scaleFactor)
        .name = shapeName
        .Fill.ForeColor.RGB = RGB(200, 200, 255)
        .Line.ForeColor.RGB = RGB(0, 0, 128)
        .Line.weight = 1.5
        .TextFrame2.TextRange.Characters.Text = Format(diameter, "0.00")
        .TextFrame2.TextRange.Characters.Font.Size = 8
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
    End With
End Sub

Sub DrawCircumscribedCircle(ws As Worksheet, diameter As Double)
    Dim scaleFactor As Double: scaleFactor = 20
    Dim leftPos As Double: leftPos = 250
    Dim topPos As Double: topPos = 50
    
    ' Рисуем описанную окружность
    With ws.Shapes.AddShape(msoShapeOval, leftPos, topPos, diameter * scaleFactor, diameter * scaleFactor)
        .name = "CircumscribedCircle"
        .Fill.Transparency = 1 ' Прозрачная заливка
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.weight = 2
        .Line.DashStyle = msoLineDash
        .TextFrame2.TextRange.Characters.Text = "Описанная окружность"
        .TextFrame2.TextRange.Characters.Font.Size = 10
        .TextFrame2.TextRange.Characters.Font.Bold = msoTrue
        .TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .TextFrame2.VerticalAnchor = msoAnchorTop
    End With
End Sub

Sub RunTightLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Расчет гофры")
    
    RepackAndRedrawCircumscribedCircle ws, 20
End Sub

' Основная процедура
Sub RepackAndRedrawCircumscribedCircle(ws As Worksheet, scaleFactor As Double)
    Dim wires() As Variant
    Dim centers() As Variant
    Dim radii() As Double
    Dim count As Long
    Dim i As Long, j As Long
    Dim x As Double, y As Double
    Dim tempX As Double, tempY As Double
    Dim placed As Boolean
    
    ' Шаги при поиске позиции
    Dim angleStep As Double: angleStep = WorksheetFunction.Pi() / 12
    
    ' Собираем все провода
    Dim sh As Shape
    For Each sh In ws.Shapes
        If sh.name Like "Wire_*" Then
            count = count + 1
            ReDim Preserve wires(1 To count)
            ReDim Preserve radii(1 To count)
            ReDim Preserve centers(1 To count)
            Set wires(count) = sh
            radii(count) = sh.Width / 2
        End If
    Next sh
    
    If count = 0 Then Exit Sub
    
    ' Первый провод в центр
    ReDim centers(1 To count)
    centers(1) = Array(900#, 300#)
    
    wires(1).Left = centers(1)(0) - radii(1)
    wires(1).Top = centers(1)(1) - radii(1)
    
    ' Расставляем остальные провода по касанию
    For i = 2 To count
        placed = False
        For j = 1 To i - 1
            Dim r As Double
            r = radii(i) + radii(j)
            
            For angle = 0 To 2 * WorksheetFunction.Pi() Step angleStep
                tempX = centers(j)(0) + r * Cos(angle)
                tempY = centers(j)(1) + r * Sin(angle)
                
                If Not IsOverlapping(tempX, tempY, radii(i), centers, radii, i - 1) Then
                    centers(i) = Array(tempX, tempY)
                    wires(i).Left = tempX - radii(i)
                    wires(i).Top = tempY - radii(i)
                    placed = True
                    Exit For
                End If
            Next angle
            If placed Then Exit For
        Next j
    Next i
    
' Находим максимальное расстояние от центра до краёв окружностей
Dim maxDist As Double: maxDist = 0
Dim dist As Double

For i = 1 To count
    dist = Sqr((centers(i)(0) - 900) ^ 2 + (centers(i)(1) - 300) ^ 2) + radii(i)
    If dist > maxDist Then maxDist = dist
Next i

' Добавим 5% запас
Dim margin As Double: margin = 0.05
Dim radiusWithMargin As Double: radiusWithMargin = maxDist * (1 + margin)

' Удаляем старую окружность
On Error Resume Next
ws.Shapes("CircumscribedCircle").Delete
On Error GoTo 0

' Рисуем новую описанную окружность
With ws.Shapes.AddShape(msoShapeOval, 900 - radiusWithMargin, 300 - radiusWithMargin, 2 * radiusWithMargin, 2 * radiusWithMargin)
    .name = "CircumscribedCircle"
    .Fill.Visible = msoFalse
    .Line.ForeColor.RGB = RGB(255, 0, 0)
    .Line.weight = 1.5
End With

' Перевод в мм
Dim boundingMM As Double
boundingMM = (2 * radiusWithMargin) / scaleFactor
ws.Range("F7").Value = boundingMM
ws.Range("F7").NumberFormat = "0.00"

End Sub

' Проверка пересечений
Function IsOverlapping(x As Double, y As Double, r As Double, centers() As Variant, radii() As Double, count As Long) As Boolean
    Dim i As Long
    For i = 1 To count
        Dim dx As Double, dy As Double, dist As Double
        dx = x - centers(i)(0)
        dy = y - centers(i)(1)
        dist = Sqr(dx * dx + dy * dy)
        If dist < (r + radii(i) - 0.1) Then
            IsOverlapping = True
            Exit Function
        End If
    Next i
    IsOverlapping = False
End Function

