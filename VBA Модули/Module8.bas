Attribute VB_Name = "Module8"
' �������� �������� ��������� �����

Sub CreateHorizontalHarnessVisualization()
    Dim ws As Worksheet
    Dim harnessCount As Integer
    Dim i As Integer, j As Integer
    Dim startRow As Integer
    Dim nodeCount As Integer
    Dim harnessName As String
    Dim startName As String, endName As String
    Dim endCol As Integer
    
    ' ��������� �����
    Set ws = ThisWorkbook.Worksheets("������ �����")
    
    ' �������� ���������� ������ �� A2
    If Not IsNumeric(ws.Range("A2").Value) Or ws.Range("A2").Value < 1 Then
        MsgBox "������� ���������� ���������� ������ � ������ A2", vbExclamation
        Exit Sub
    End If
    harnessCount = ws.Range("A2").Value
    
    ' ��������� ������� ���������� ������������ (E10:XFD100)
    With ws.Range("E10:AF100")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    ' ��������� ������� ��� ������� �����
    startRow = 12 ' �������� � E12
    
    ' ������� ������������ ��� ������� �����
    For i = 1 To harnessCount
        ' �������� �������� ����� �� ������� C
        harnessName = ws.Range("C" & 1 + i).Value ' C2 ��� ������� �����
        If harnessName = "" Then harnessName = "���� " & i
        
        ' �������� ���������� ����� �� B8 �� ��������� �� 5 ����� ��� ������� ���������� �����
        nodeCount = ws.Range("B" & 8 + (i - 1) * 5).Value
        If nodeCount < 1 Then nodeCount = 1 ' ������� 1 ����
        
        ' �������� �������� ��� "������" � "�����"
        startName = ws.Range("B" & 6 + (i - 1) * 5).Value
        If startName = "" Then startName = "������"
        
        endName = ws.Range("B" & 7 + (i - 1) * 5).Value
        If endName = "" Then endName = "�����"
        
        ' ���������� �������� ����� (� ��� �� ������, ����� �� ������)
        With ws.Cells(startRow, 4) ' ������� D
            .Value = harnessName
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        
        ' ������ ����� (E12, E22 � �.�.)
        With ws.Cells(startRow, 5) ' ������� E
            .Value = startName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        ' ������� ���� (������������� ������)
        For j = 1 To nodeCount * 2 ' �������� �� 2 (������ + ����)
            If j Mod 2 = 0 Then
                ' ������ ������� - ���� (������)
                With ws.Cells(startRow, 5 + j)
                    .Value = 0 ' ����� ����
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Interior.Color = RGB(0, 0, 0)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                End With
            ' �������� ������� �������� �������
            End If
        Next j
        
        ' ����� �����
        endCol = 5 + nodeCount * 2 + 1
        With ws.Cells(startRow, endCol)
            .Value = endName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        ' �������������� ������ ��� ����� �����
        With ws.Range(ws.Cells(startRow, 5), ws.Cells(startRow, endCol))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders.weight = xlMedium
            .ColumnWidth = 8
        End With
        
        ' ��������� ��� ���������� ����� (+15 �����)
        startRow = startRow + 15
    Next i
    
    ' ����������� ������ ������� D ��� ��������
    ws.Columns("D").ColumnWidth = 15
    
    MsgBox "�������������� ������������ ������� ��� " & harnessCount & " ������", vbInformation
End Sub


