Attribute VB_Name = "Module7"
'����� ������������ �������� �����, ������ �� ��������� �������� ��������

Sub FindOptimalGofraDiameter()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim searchValue As Range, resultCell As Range, extDiamCell As Range
    Dim diametersRange As Range, extDiametersRange As Range
    Dim selectedOption As String
    Dim targetDiameter As Double, foundDiameter As Double, foundExtDiameter As Double
    Dim i As Long
    
    ' ��������� ������
    Set wsCalc = ThisWorkbook.Worksheets("������ �����")
    Set wsData = ThisWorkbook.Worksheets("��������������� ������")
    
    ' ��������� �������� � G10
    Set searchValue = wsCalc.Range("G10")
    selectedOption = Trim(searchValue.Value)
    
    ' ������ ��� �����������
    Set resultCell = wsCalc.Range("I7") ' ���������� �������
    Set extDiamCell = wsCalc.Range("J7") ' ������� �������
    
    ' ������� ������ ����������
    resultCell.ClearContents
    extDiamCell.ClearContents
    
    ' �������� ������� ������� ��� ������ (�� H7)
    If Not IsNumeric(wsCalc.Range("H7").Value) Then
        MsgBox "������� ����������� ������� ��������� ����������!", vbExclamation
        Exit Sub
    End If
    targetDiameter = wsCalc.Range("H7").Value
    
    ' �������� ��������� � ����������� �� ��������
    If selectedOption = "��" Then
        Set diametersRange = wsData.Range("Q9:Q33").SpecialCells(xlCellTypeConstants)
        Set extDiametersRange = wsData.Range("R9:R33").SpecialCells(xlCellTypeConstants)
    ElseIf selectedOption = "���" Then
        Set diametersRange = wsData.Range("T9:T20").SpecialCells(xlCellTypeConstants)
        Set extDiametersRange = wsData.Range("U9:U20").SpecialCells(xlCellTypeConstants)
    Else
        MsgBox "�������� '��' ��� '���' � ������ G10", vbExclamation
        Exit Sub
    End If
    
    ' ���� ��������� ������� ���������� �������
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
    
    ' ���� �� ����� - ����� ������������ ���������
    If foundDiameter = 0 Then
        foundDiameter = Application.WorksheetFunction.Max(diametersRange)
        foundExtDiameter = extDiametersRange.Cells(diametersRange.Cells.count).Value
        MsgBox "���������� ������� �� ������. ����������� ������������ ���������: " & foundDiameter, vbInformation
    End If
    
    ' ���������� ����������
    resultCell.Value = foundDiameter
    resultCell.NumberFormat = "0.00"
    
    extDiamCell.Value = foundExtDiameter
    extDiamCell.NumberFormat = "0.00"
    
    MsgBox "�������� ���������� ������� �����: " & foundDiameter & vbCrLf & _
           "��������������� ������� �������: " & foundExtDiameter, vbInformation
End Sub
