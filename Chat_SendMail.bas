Attribute VB_Name = "Module1"
Sub Mail()

    If Range("C2") = "" Then
        x = Range("B2")
    Else:
        x = Range("C2")
    End If
    
    Select Case x
        Case "�x�\��", "��n�T", "���z�s", "�L�|��", "����", "�d���z", "���Hޱ"
            cc = "���ɶ�; QA�p�� <QA??@aptg.com.tw>"
        Case "�\�z�g", "�L���o", "���ئ�", "�_�s��", "�f����"
            cc = "�����"
        Case "�i�[�s", "���X��", "���j��", "�B�Ӯp", "������", "�d���@"
            cc = "������"
    End Select
        
        
    Dim olApp As Object
    Dim olMailItm As Object
    
    Columns("E:H").Hidden = False
        
    Set olApp = CreateObject("Outlook.Application")
    Set olMailItm = olApp.CreateItem(0)
   
    With olMailItm
        .To = "CS-NH-Duty; CS-TY-Duty"
        .cc = cc
        .Subject = Range("C7") & "-�iChat�d��нШ�U�w�Ʀ^�q�j" & Range("C4")
        .HTMLBody = RangetoHTML(Range("E1:H6"))
        .Display
    End With
       
    Set olMailItm = Nothing
    Set olApp = Nothing
    
    Columns("E:H").Hidden = True
    
End Sub

Sub �M��()

Range("C3:C7").ClearContents

End Sub

Sub user()

MsgBox Application.UserName

End Sub


Function RangetoHTML(Rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

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

