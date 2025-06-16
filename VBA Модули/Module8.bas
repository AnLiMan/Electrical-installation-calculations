Attribute VB_Name = "Module8"
' Создание линейной структуры жгута

Sub CreateHorizontalHarnessVisualization()
    Dim ws As Worksheet
    Dim harnessCount As Integer
    Dim i As Integer, j As Integer
    Dim startRow As Integer
    Dim nodeCount As Integer
    Dim harnessName As String
    Dim startName As String, endName As String
    Dim endCol As Integer
    
    ' Настройка листа
    Set ws = ThisWorkbook.Worksheets("Расчет жгута")
    
    ' Получаем количество жгутов из A2
    If Not IsNumeric(ws.Range("A2").Value) Or ws.Range("A2").Value < 1 Then
        MsgBox "Укажите корректное количество жгутов в ячейке A2", vbExclamation
        Exit Sub
    End If
    harnessCount = ws.Range("A2").Value
    
    ' Полностью очищаем предыдущую визуализацию (E10:XFD100)
    With ws.Range("E10:AF100")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    ' Начальная позиция для первого жгута
    startRow = 12 ' Начинаем с E12
    
    ' Создаем визуализацию для каждого жгута
    For i = 1 To harnessCount
        ' Получаем название жгута из столбца C
        harnessName = ws.Range("C" & 1 + i).Value ' C2 для первого жгута
        If harnessName = "" Then harnessName = "Жгут " & i
        
        ' Получаем количество узлов из B8 со смещением на 5 строк для каждого следующего жгута
        nodeCount = ws.Range("B" & 8 + (i - 1) * 5).Value
        If nodeCount < 1 Then nodeCount = 1 ' Минимум 1 узел
        
        ' Получаем названия для "начала" и "конца"
        startName = ws.Range("B" & 6 + (i - 1) * 5).Value
        If startName = "" Then startName = "Начало"
        
        endName = ws.Range("B" & 7 + (i - 1) * 5).Value
        If endName = "" Then endName = "Конец"
        
        ' Записываем название жгута (в той же строке, слева от начала)
        With ws.Cells(startRow, 4) ' Колонка D
            .Value = harnessName
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        
        ' Начало жгута (E12, E22 и т.д.)
        With ws.Cells(startRow, 5) ' Колонка E
            .Value = startName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        ' Создаем узлы (горизонтально вправо)
        For j = 1 To nodeCount * 2 ' Умножаем на 2 (пустая + узел)
            If j Mod 2 = 0 Then
                ' Четные столбцы - узлы (черные)
                With ws.Cells(startRow, 5 + j)
                    .Value = 0 ' Номер узла
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Interior.Color = RGB(0, 0, 0)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                End With
            ' Нечетные столбцы остаются пустыми
            End If
        Next j
        
        ' Конец жгута
        endCol = 5 + nodeCount * 2 + 1
        With ws.Cells(startRow, endCol)
            .Value = endName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        ' Форматирование границ для всего жгута
        With ws.Range(ws.Cells(startRow, 5), ws.Cells(startRow, endCol))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders.weight = xlMedium
            .ColumnWidth = 8
        End With
        
        ' Смещаемся для следующего жгута (+15 строк)
        startRow = startRow + 15
    Next i
    
    ' Настраиваем ширину столбца D для названий
    ws.Columns("D").ColumnWidth = 15
    
    MsgBox "Горизонтальная визуализация создана для " & harnessCount & " жгутов", vbInformation
End Sub


