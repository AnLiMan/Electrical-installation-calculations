Attribute VB_Name = "Module3"
'������� ����������� ���������� ����������� � AWG

Sub CalculateSquareMMtoAWG()
    Dim squareMM As Double
    Dim awgValue As Double
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim awgStandardsRange As Range
    Dim resultCalculateCell As Range, resultCell As Range
    
    ' ��������� ������
    Set wsCalc = ThisWorkbook.Worksheets("������")
    Set wsData = ThisWorkbook.Worksheets("��������������� ������")
    
    ' ������
    squareMM = wsCalc.Range("H2").Value                  ' ������ � �� �� �� ����� "������"
    Set awgStandardsRange = wsData.Range("A33:A48")       ' �������� AWG ����������
    Set resultCalculateCell = wsCalc.Range("H3")          ' ��������� �������� AWG
    Set resultCell = wsCalc.Range("H5")                  ' ����������� �������� AWG
    
    ' �������� �����
    If Not IsNumeric(squareMM) Or squareMM <= 0 Then
        MsgBox "������� �������� �������� �������!", vbExclamation
        Exit Sub
    End If
    
    ' ������� �������� �� ��. � AWG
    awgValue = 36 - (19.5 * Log(squareMM / 0.012668) / Log(92))
    resultCalculateCell.Value = awgValue
    
    ' ����� ���������� ������������ AWG
    Dim closestAWG As Double
    closestAWG = FindClosestAWG(awgValue, awgStandardsRange)
    
    ' ����� ����������
    resultCell.Value = closestAWG
    MsgBox "���������: " & squareMM & " �� ��. = AWG " & closestAWG, vbInformation
End Sub

'������� ������ ���������� ������������ �������� �� ���� AWG

Function FindClosestAWG(targetAWG As Double, awgStandardsRange As Range) As Double
    Dim cell As Range
    Dim closest As Double
    Dim minDiff As Double
    Dim tolerance As Double
    
    Set wsCalc = ThisWorkbook.Worksheets("������")
    
    tolerance = wsCalc.Range("H4").Value  ' ������ �� ������ H4
    minDiff = 1E+30
    closest = 0
    
    ' ���� ��������� �������� AWG (������� = ������� �������)
    For Each cell In awgStandardsRange
        If IsNumeric(cell.Value) Then
            Dim currentAWG As Double
            currentAWG = cell.Value
            
            ' ��� AWG: ��� ������ �����, ��� ������ �������
            ' ���� �������� � �������� �������
            If currentAWG >= targetAWG * (1 - tolerance) And _
               currentAWG <= targetAWG * (1 + tolerance) Then
                ' �������� ��������� ������� (�� ������ �������)
                If Abs(currentAWG - targetAWG) < minDiff Then
                    minDiff = Abs(currentAWG - targetAWG)
                    closest = currentAWG
                End If
            End If
        End If
    Next cell
    
    ' ���� �� ����� � �������� �������, ����� ���������
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

