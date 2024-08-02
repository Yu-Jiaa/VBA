Attribute VB_Name = "Module1"
Sub �����()
   Call �Ƨ�
   'Setting up the Excel variables.
   Dim olApp As Object
   Dim olMailItm As Object
 
   'Create the Outlook application and the empty email.
   Set olApp = CreateObject("Outlook.Application")
   Set olMailItm = olApp.CreateItem(0)
   
   Days = Range("K1")   'Debug.Print Days
   
   If Days = 1 Then
        Day1 = Format(Date - 1, "mm/dd")
   ElseIf Days > 1 Then
        Day1 = Format(Date - Days, "mm/dd") & "-" & Format(Date - 1, "mm/dd")
   End If

   mStr = "Dear Duty" & vbCrLf & "�Ш�U�G�F , ����" & vbCrLf & "1.�U��έp" & Day1 & _
        " KM�W�[��T�A���ѤH�����ѽs���B KM��m�Υi�ϥέ�������A�i�d�ߨ즹�h���ѡA�ФH���h�[�Q�ΡC" & vbCrLf & "2.�w�T�{�G�F���e�Ҥw���Ѹ�ƨӷ��C"
   LF = Cells(Rows.Count, "A").End(xlUp).Row
   
   For x = 1 To LF
    If Range("A" & x) = Date - Days Then
        Exit For
    End If
   Next
   
   With olMailItm
       .To = "CS-NH-Duty; CS-TY-Duty"
       .Cc = "²�\��; �H���x; �����; ���ɶ�; QA�p�� <QA??@aptg.com.tw>"
       .Subject = "�i�Ш�U�G�F�jKM�W�[�����_" & Format(Date, "yyyymmdd")
       .HTMLBody = mStr & RangetoHTML(Range("A1:H1,A" & x & ":H" & LF))
       .Display
   End With
   
   'Clean up the Outlook application.
   Set olMailItm = Nothing
   Set olApp = Nothing
   
   ActiveWorkbook.Save
   
End Sub

'�ޥ� Ron de Bruin �� RangetoHTML �{��
'Reference: http://www.rondebruin.nl/cdo.htm
Function RangetoHTML(Rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2007
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    Rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Sub �Ƨ�()

With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("A1") ', SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range("B1"), CustomOrder:="�@�~�y�{,�t�κ��@,��P�q��,�ȪAApp,�Ш|�V�m,�Τ�׺�,���a�s��"
    .SetRange Range("A1").CurrentRegion
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Cells.EntireRow.AutoFit
ActiveWorkbook.Save

End Sub

Sub ����()
Call �Ƨ�
ActiveSheet.Copy
Columns("I:Q").Delete
Range("A1").Select
ActiveWorkbook.SaveAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\KM�W�[�����(" & Worksheets(1).Name & ").xlsx"
'CreateObject("WScript.Shell").SpecialFolders("Desktop")  '�ୱ���|

'ActiveWorkbook.SaveAs "C:\Users\yujialin\Desktop\KM�W�[�����(" & Worksheets(1).Name & ").xlsx"

ActiveWorkbook.Close
MsgBox "�w�x�s�b�ୱ�I"

End Sub

Sub �ƻs()

Dim d As New DataObject

i = Selection.Row()
Str1 = "Hi Duty" & vbCrLf & vbCrLf & "��s�����p�U���|�G" & vbCrLf & Range("E" & i) & vbCrLf & "���ѽs���G" & Range("C" & i) _
& vbCrLf & vbCrLf & "CS-NH-Duty <cs-nh-duty@aptg.com.tw>; CS-TY-Duty <cs-ty-duty@aptg.com.tw>" _
& vbCrLf & "���ɶ� <evchen@aptg.com.tw>;  TQ�p�� <TQ-Group@aptg.com.tw>; QA�p�� <QA??@aptg.com.tw>" _
& vbCrLf & "�iKM��s�����j"


'vbCrLf:����

d.SetText Str1
d.PutInClipboard '��J�ŶKï

Range("J" & i).Interior.color = RGB(255, 0, 0) '�]�w����(��)
Application.Wait (Now + TimeValue("00:00:01")) '����1��
Range("J" & i).Interior.color = xlNone '�A�]�w����(�L��)

End Sub

Sub Username()

'MsgBox Environ("Username")

MsgBox CreateObject("WScript.Shell").SpecialFolders("Desktop")

End Sub


