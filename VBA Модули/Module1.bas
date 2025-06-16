Attribute VB_Name = "Module1"
'������� ����������� AWG � ���������� ����������

Sub CalculateAWGtoSquareMM()
    Dim awgValue As Double
    Dim squareMM As Double
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim standardsRange As Range
    Dim resultCell As Range
    
    ' ��������� ������
    Set wsCalc = ThisWorkbook.Worksheets("������")             ' ���� � ������ AWG
    Set wsData = ThisWorkbook.Worksheets("��������������� ������") ' ���� �� ������������ ���������
    
    ' ������
    awgValue = wsCalc.Range("E2").Value                       ' ������ � AWG �� ����� "������"
    Set standardsRange = wsData.Range("A10:A29")               ' �������� ���������� �� ��������������� �����
    Set resultCalculateCell = wsCalc.Range("E3")               ' ������ ��� ���������� � ��������� ���������
    Set resultCell = wsCalc.Range("E5")                       ' ������ ��� ���������� �� ����������� ���������
    
    ' ��������, ��� ������� �����
    If Not IsNumeric(awgValue) Then
        MsgBox "������� �������� �������� AWG � ������ E2!", vbExclamation
        Exit Sub
    End If
    
    ' ������� �������� AWG � �� ��.
    squareMM = 0.012668 * (92 ^ ((36 - awgValue) / 19.5))
    resultCalculateCell.Value = squareMM
    
    ' ����� ���������� ������������ ��������
    Dim closestStandard As Double
    closestStandard = FindClosestStandard(squareMM, standardsRange)
    
    ' ����� ����������
    resultCell.Value = closestStandard
    MsgBox "���������: AWG " & awgValue & " = " & closestStandard & " �� ��", vbInformation
End Sub

' ������� ������ ���������� ������������ �������� �� ���

Function FindClosestStandard(targetValue As Double, standardsRange As Range) As Double
    Dim cell As Range
    Dim closest As Double
    Dim minDiff As Double
    Dim tolerance As Double
    
    Set wsCalc = ThisWorkbook.Worksheets("������")
    
    tolerance = wsCalc.Range("E4")
    minDiff = 1E+30
    closest = 0
    
    ' ������� ���� ��������� ������� � �������� 5%
    For Each cell In standardsRange
        If IsNumeric(cell.Value) Then
            Dim currentVal As Double
            currentVal = cell.Value
            
            ' ���� �������� � �������� �5% (��� ����� ��������� ��������) � >= ��������
            If currentVal >= targetValue * (1 - tolerance) Then
                If currentVal <= targetValue * (1 + tolerance) Then
                    If currentVal < closest Or closest = 0 Then
                        closest = currentVal
                    End If
                End If
            End If
        End If
    Next cell
    
    ' ���� �� ����� � �������� 5% (��� ����� ��������� ��������), ���� ��������� �������
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
    
    ' ���� ��� ��������� ������, ���� ������������
    If closest = 0 Then
        closest = Application.WorksheetFunction.Max(standardsRange)
    End If
    
    FindClosestStandard = closest
End Function
