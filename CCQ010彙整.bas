Attribute VB_Name = "Module1"
Sub ����1()
Attribute ����1.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Rng As Range '�۰ʿz�ﵲ�G�d��
Dim Book1 As Workbook
Set Book1 = ActiveWorkbook

'���ɦW
Filename = Book1.Name
Filename = "CCQ010(" & Mid(Filename, 8, 7) & ")"

'�z��CCQ010�ýƻs
With Book1.ActiveSheet
    Set Rng = .UsedRange
    Rng.AutoFilter Field:=23, Criteria1:="=*CCQ010*", Operator:=xlAnd
    Set Rng = .UsedRange
    Rng.Copy
End With

'�K��s����ï
Workbooks.Add
ActiveSheet.Paste

'�Ƨ�(CSR>�ɶ�>����)
With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("U1")
    .SortFields.Add Key:=Range("O1")
    .SortFields.Add Key:=Range("B1")
    .SetRange Range("A1").CurrentRegion
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'�����ήɶ�-���Ƹ�Ƶ��O
Range("B:B,O:O").Select
Selection.FormatConditions.AddUniqueValues
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).DupeUnique = xlDuplicate
With Selection.FormatConditions(1).Font
    .Color = -16383844
    .TintAndShade = 0
End With

With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
End With
Selection.FormatConditions(1).StopIfTrue = False

'�����ήɶ�-�վ���e
Columns("B:B").EntireColumn.AutoFit
Columns("O:O").EntireColumn.AutoFit

Range("A1").Select

'��J�K�X&�t�s�ɮ�&�����ɮ�
ActiveSheet.Name = Filename & "��l��"
ActiveSheet.SaveAs Book1.Path & "\" & Filename & ".xlsx"

'DataPass= InputBox("�п�J�n�]�w���K�X","�]�w�ɮ׶}�ұK�X")
ActiveWorkbook.Password = "16850"
ActiveWorkbook.Save

Book1.Close False 'False:���b������ï���Ұ��������ܧ�

End Sub

