Attribute VB_Name = "Module4"
'������� ������� ������ � ����������� �� ����� �� ���������� �������

Sub SelectCableSectionByType()
    Dim wsCalc As Worksheet, wsData As Worksheet
    Dim cableTypeCell As Range, currentSectionCell As Range, resultSectionCell As Range
    Dim cableType As String, currentSection As Double
    Dim cableTypesRange As Range, sectionsTable As Range
    Dim targetColumn As Range, resultSection As Double
    
    ' ��������� ������
    Set wsCalc = ThisWorkbook.Worksheets("������")
    Set wsData = ThisWorkbook.Worksheets("��������������� ������")
    
    ' ������ �����/������
    Set cableTypeCell = wsCalc.Range("B19")          ' ������ � ���������� ������� ����� ������
    Set currentSectionCell = wsCalc.Range("B15")     ' ������� ������� (�� ��) ��� �������
    Set resultSectionCell = wsCalc.Range("B20")      ' ���������: ����������� �������
    
    ' �������� �����
    If Not IsNumeric(currentSectionCell.Value) Or currentSectionCell.Value <= 0 Then
        MsgBox "������� ���������� �������� ������� (������������� �����)!", vbExclamation
        Exit Sub
    End If
    
    cableType = cableTypeCell.Value
    currentSection = currentSectionCell.Value
    
    ' ����� ������� � ��������� ��� ��������� ����� ������
    Set cableTypesRange = wsData.Range("B9:D17")     ' ��������� � ������� ������� � 1-� ������
    Set targetColumn = Nothing
    
    On Error Resume Next
    Set targetColumn = cableTypesRange.Find(What:=cableType, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If targetColumn Is Nothing Then
        MsgBox "����� ������ '" & cableType & "' �� �������!", vbExclamation
        Exit Sub
    End If
    
    ' �������� ������� ��� ��������� ����� (������ ���������� �� 2-� ������)
    Set sectionsTable = wsData.Range(wsData.Cells(2, targetColumn.Column), _
                        wsData.Cells(100, targetColumn.Column)).SpecialCells(xlCellTypeConstants)
    
    ' ����� ���������� �������� �������
    resultSection = FindClosestSection(currentSection, sectionsTable)
    
    ' ����� ����������
    resultSectionCell.Value = resultSection
    MsgBox "��� ����� " & cableType & vbCrLf & _
           "��������� �������: " & currentSection & " �� ��." & vbCrLf & _
           "������������� �������: " & resultSection & " �� ��.", vbInformation
End Sub

'����� ���������� �������� ������� ������ � ����������� �� ��������� �����

Function FindClosestSection(targetSection As Double, sectionsRange As Range) As Double
    Dim cell As Range
    Dim closest As Double
    
    closest = Application.WorksheetFunction.Max(sectionsRange) ' �� ��������� - ������������
    
    For Each cell In sectionsRange
        If IsNumeric(cell.Value) Then
            If cell.Value >= targetSection And cell.Value < closest Then
                closest = cell.Value
            End If
        End If
    Next cell
    
    FindClosestSection = closest
End Function
