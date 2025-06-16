Attribute VB_Name = "Module2"
' ������� ������� ������������� ��� ������ ������� �������

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
    
    ' ��������� ������
    Set wsCalc = ThisWorkbook.Worksheets("������")
    Set wsData = ThisWorkbook.Worksheets("��������������� ������")
    
    ' ������ �����
    Set voltageDropCell = wsCalc.Range("B7") ' ������ � ���������� �������� ����������
    Set voltageCell = wsCalc.Range("B6") ' ������ � ����������� ����
    Set currentCell = wsCalc.Range("B4") ' ������ � �����
    Set materialCell = wsCalc.Range("B2") ' ������ � ����������
    Set lengthCell = wsCalc.Range("B3") ' ������ � ������
    Set temperatureCell = wsCalc.Range("B5") ' ������ � ������������
    
    ' ������ ������
    Set resultResistanceCell = wsCalc.Range("B13")
    Set resultResistanceCellTemp = wsCalc.Range("B14")
    Set resultSectionCell = wsCalc.Range("B15")
    Set resultFinalSectionCell = wsCalc.Range("B16")
    Set resultVoltageDrop = wsCalc.Range("B17")
    
    ' �������� ������
    If Not IsNumeric(voltageDropCell.Value) Or Not IsNumeric(voltageCell.Value) Or _
       Not IsNumeric(currentCell.Value) Or Not IsNumeric(lengthCell.Value) Or _
       Not IsNumeric(temperatureCell.Value) Then
        MsgBox "��������� �������� ��������: ��� ��� ������ ���� �������!", vbExclamation
        Exit Sub
    End If
    
    '����������� ���������� �������� ��������
    voltageDrop = voltageDropCell.Value
    voltage = voltageCell.Value
    current = currentCell.Value
    length = lengthCell.Value
    temperature = temperatureCell.Value
    selectedMaterial = materialCell.Value
    
    '�������� ����� ������� � ���� �� ������� ��������
    If current = 0 Or length = 0 Then
        MsgBox "��� � ����� ���������� �� ����� ���� ��������!", vbExclamation
        Exit Sub
    End If
    
    ' 1. ������ ������������� �������������
    maxResistance = (voltageDrop * voltage) / current
    
    ' 2. ��������� ������������� ���������
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
        MsgBox "�������� '" & selectedMaterial & "' �� ������ � ��������!", vbExclamation
        Exit Sub
    End If
    
    ' 3. ������ ������������� � ������������� ���������
    currentResistance = maxResistance / (1 + tempCoeff * (temperature - 20))
    
    ' 4. ������ �������
    calculatedSection = resistivity * length / currentResistance
    
    ' 5. ������ ������������ �������
    Dim standardSections As Range
    Set standardSections = wsData.Range("A10:A30")
    finalSection = Application.WorksheetFunction.Max(standardSections)
    
    For i = 1 To standardSections.Rows.count
        If standardSections.Cells(i, 1).Value >= calculatedSection Then
            finalSection = standardSections.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    ' 6. �������� ������������� ���� ��� ���������� �������
    Dim currentTable As Range
    Dim sectionFound As Boolean: sectionFound = False
    maxCurrentForSection = 0
    
    Set currentTable = wsData.Range("F10:G29") ' ������� �������-��� (F - �������, G - ���)
    
    For i = 1 To currentTable.Rows.count
        If currentTable.Cells(i, 1).Value = finalSection Then
            maxCurrentForSection = currentTable.Cells(i, 2).Value
            sectionFound = True
            Exit For
        End If
    Next i
    
    If Not sectionFound Then
        MsgBox "��� ������� " & finalSection & " �� ��. �� ������� �������� ������������� ����!", vbExclamation
    End If
    
    ' 7. ����� �����������
    resultResistanceCell.Value = maxResistance '������������� ������ �����������, �� (���. �����������)
    resultResistanceCell.NumberFormat = "0.000"
    
    resultResistanceCellTemp.Value = currentResistance '������������� ������ �����������, �� (���. �����������)
    resultResistanceCellTemp.NumberFormat = "0.000"
    
    resultSectionCell.Value = calculatedSection '��������� �������, ��?
    resultSectionCell.NumberFormat = "0.000"
    
    resultFinalSectionCell.Value = finalSection '������� ������ �� ���, ��?
    resultFinalSectionCell.NumberFormat = "0.00"
    
    correctedResistance = resistivity * length / finalSection
    calculateVoltageDrop = current * correctedResistance '�������� ������� ����������, �
    resultVoltageDrop.Value = calculateVoltageDrop
    resultVoltageDrop.NumberFormat = "0.000"
    
    ' �������� �� ���������� ������������� ����
    Dim warningMsg As String
    warningMsg = ""
    
    If current > maxCurrentForSection And maxCurrentForSection > 0 Then
        warningMsg = vbCrLf & "��������! ��� �������� (" & current & " �) ��������� ������������ ��� ����� ������� (" & maxCurrentForSection & " �). ��������� �������� �� ���������!"
    End If
    
    ' �������������� ���������
    MsgBox "������ ��������:" & vbCrLf & _
           "��������� �������: " & Format(calculatedSection, "0.0000") & " �� ��." & vbCrLf & _
           "����������� �������: " & finalSection & " �� ��." & vbCrLf & _
           "����. ��� ��� �������: " & maxCurrentForSection & " �" & warningMsg, _
           vbInformation
End Sub
