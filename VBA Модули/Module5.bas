Attribute VB_Name = "Module5"
'������ ������� ��� ������ �������

Sub CalculateMultipleResistancesWithPUEandTemp()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim materialCell As Range, voltageCell As Range, voltageDropCell As Range, tempCell As Range
    Dim lengthsRange As Range, currentsRange As Range, resultsRange As Range, pueSectionsRange As Range
    Dim selectedMaterial As String, voltage As Double, voltageDrop As Double, temperature As Double
    Dim resistivity As Double, tempCoeff As Double
    Dim i As Long, lastRow As Long
    
    ' ��������� ������
    Set wsCalc = ThisWorkbook.Worksheets("������")
    Set wsData = ThisWorkbook.Worksheets("��������������� ������")
    
    ' ������� ���������
    Set materialCell = wsCalc.Range("B2") ' ������ � ����������
    Set voltageCell = wsCalc.Range("B6") ' ������ � ����������� ����
    Set voltageDropCell = wsCalc.Range("B7") ' ������ � ���������� ������� ������� ����������
    Set tempCell = wsCalc.Range("B5") ' ������ � ������������
    
    ' �������� ����������� ������� �� ���
    Set pueSectionsRange = wsData.Range("A10:A30")
    
    ' ���������� ��������� ������
    lastRow = wsCalc.Cells(wsCalc.Rows.count, "B").End(xlUp).Row
    If lastRow < 26 Then lastRow = 26
    
    Set lengthsRange = wsCalc.Range("B26:B" & lastRow)
    Set currentsRange = wsCalc.Range("C26:C" & lastRow)
    Set resultsRange = wsCalc.Range("D26:D" & lastRow)
    
    ' �������� ������� ����������
    If Not IsNumeric(voltageCell.Value) Or Not IsNumeric(voltageDropCell.Value) Or _
       Not IsNumeric(tempCell.Value) Then
        MsgBox "��������� ����������, ���������� ������ � ����������� � B5-B7!", vbExclamation
        Exit Sub
    End If
    
    '����������� ���������� �������� ��������
    voltage = voltageCell.Value
    voltageDrop = voltageDropCell.Value
    temperature = tempCell.Value
    selectedMaterial = materialCell.Value
    
    ' �������� �������������� ���������
    Dim resistRange As Range, tempCoeffRange As Range
    Set resistRange = wsData.Range("A2:B4")
    Set tempCoeffRange = wsData.Range("D2:E4")
    
    Dim materialFound As Boolean: materialFound = False
    
    ' ����� ��������� �������������
    For i = 1 To resistRange.Rows.count
        If Trim(resistRange.Cells(i, 1).Value) = Trim(selectedMaterial) Then
            resistivity = resistRange.Cells(i, 2).Value
            materialFound = True
            Exit For
        End If
    Next i
    
    ' ����� �������������� ������������
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
        MsgBox "�������� '" & selectedMaterial & "' �� ������ � ��������!", vbExclamation
        Exit Sub
    End If
    
    ' ������� ������ ����������
    resultsRange.ClearContents
    
    ' ��������� ������ ������ � ������ �����������
    For i = 1 To lengthsRange.Rows.count
        If IsNumeric(lengthsRange.Cells(i, 1).Value) And _
           IsNumeric(currentsRange.Cells(i, 1).Value) And _
           currentsRange.Cells(i, 1).Value <> 0 Then
            
            Dim length As Double, current As Double
            Dim maxResistance As Double, currentResistance As Double
            Dim calculatedSection As Double, standardSection As Double
            
            length = lengthsRange.Cells(i, 1).Value
            current = currentsRange.Cells(i, 1).Value
            
            ' 1. ������ ������������� ������������� ��� ����� �����������
            maxResistance = (voltageDrop * voltage) / current
            
            ' 2. ��������� ������������� � ������ �����������
            currentResistance = maxResistance / (1 + tempCoeff * (temperature - 20))
            
            ' 3. ������ ������� � ������ �������������� �������������
            calculatedSection = resistivity * length / currentResistance
            
            ' 4. ������ ������������ ������� �� ���
            standardSection = FindClosestPueSection(calculatedSection, pueSectionsRange)
            
            ' ���������� ���������
            resultsRange.Cells(i, 1).Value = standardSection
            resultsRange.Cells(i, 1).NumberFormat = "0.00"
        Else
            resultsRange.Cells(i, 1).Value = "-"
        End If
    Next i
    
    MsgBox "������ �������� ��� " & lengthsRange.Rows.count & " �����." & vbCrLf & _
           "������ ������� �����������: " & temperature & "�C" & vbCrLf & _
           "������������� �����������: " & tempCoeff & " 1/�C", vbInformation
End Sub

'������� ������ ���������� �������� ������� �� ���

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

