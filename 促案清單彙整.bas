Attribute VB_Name = "Module1"
Public xMonth

Sub �P��()

xMonth = InputBox("�п�J�s�@�~���Τ��" & vbNewLine & "�ҡG202305")
'xMonth = "202304"
Worksheets("�u�@��1").Name = "�Ҧ��P�ײM��"

'���J�iA��j-�`�ΡA���(VlookUp) �i����j���iO��j
Columns("A").Insert
Range("A1") = "�`��"

xLast = Columns("B").End(xlDown).Row
For x = 2 To xLast
    Range("A" & x) = "=VLOOKUP(D" & x & ",���!O:O,1,0)"
Next

'�������iF~I�BK�BO~AF�BAI~AL�BAN~AO�j
Range("F:I, K:K, O:AF, AI:AL, AN:AO").EntireColumn.Hidden = True

'�i�z��ýƻs�j�����쪺���(������#N/A)
Range("A1").AutoFilter Field:=1, Criteria1:="<>#N/A"
Columns("C:AM").Copy

'�K�W�s�u�@��iX������P�ײM��j
Worksheets.Add Before:=Worksheets("�Ҧ��P�ײM��")
Worksheets(1).Name = Right(xMonth, 1) & "������P�ײM��"
ActiveSheet.Paste Range("A1")
Application.CutCopyMode = False

'�i�Ҧ��P�ײM��j�u�@��-�Ѱ��i�z��j�C
Worksheets("�Ҧ��P�ײM��").AutoFilterMode = False

'�iX������P�ײM��j�u�@��
'�ƧǡG�i���~�O�j��A��Z�Ƨ�
Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

'�ƻs�iA~D��j��s�u�@��iX������P�ץN�X�j
xLast = Columns("A").End(xlDown).Row
Range("A1:D" & xLast).Copy
Worksheets.Add Before:=Worksheets(1)
Worksheets(1).Name = Right(xMonth, 1) & "������P�ץN�X"
ActiveSheet.Paste Range("A1")
Application.CutCopyMode = False

'�i����j�u�@��
' �N�iBG��j(����M�ץ����ꦬ�믲) �i�ഫ���Ʀr�j
xLast = Columns("BG").End(xlDown).Row
For x = 2 To xLast
   Worksheets("���").Range("BG" & x) = Worksheets("���").Range("BG" & x).Value
Next

Call �ϯä��R��

End Sub

Sub test()

End Sub


Public Sub �ϯä��R��()

Dim PTCache As PivotCache
Dim PT As PivotTable

    Set PTCache = ThisWorkbook.PivotCaches.Add _
        (SourceType:=xlDatabase, SourceData:=Sheets("���").Range("A1").CurrentRegion.Address)

    Set PT = PTCache.CreatePivotTable(TableDestination:="", TableName:="�ϯä��R��1")
         
    ActiveSheet.Name = "�ϯ�"
    
    With ActiveSheet.PivotTables("�ϯä��R��1")
        .PivotFields("�O�_����").Orientation = xlPageField '����z��
        .PivotFields("�O�_����").CurrentPage = "������"
        .PivotFields("�M�ץN�X").Orientation = xlRowField '�C����
        .PivotFields("�M�צW��").Orientation = xlRowField '�C����"
        .PivotFields("����M�ץ����ꦬ�믲").Orientation = xlColumnField '�����
        .PivotFields("����M�ץ����ꦬ�믲").AutoSort xlAscending, "����M�ץ����ꦬ�믲" '�Ƨ�
        .AddDataField ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("����"), "�������P�ױƦW�έp", xlCount '��
        .RowAxisLayout xlTabularRow '�H�C��覡���
        .PivotFields("�M�ץN�X").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)  '�����p�p
    End With

    Worksheets("�ϯ�").Range("A3").CurrentRegion.Copy
    Worksheets.Add Before:=Worksheets(1)
    Worksheets(1).Name = Right(xMonth, 1) & "������P�ױƦW"
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

End Sub

Sub Sav()

MsgBox (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & xMonth & "�ȪA����P�ײM��")
ActiveWorkbook.SaveAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & xMonth & "�ȪA����P�ײM��.xlsm"
'CreateObject("WScript.Shell").SpecialFolders("Desktop")  '�ୱ���|

End Sub

