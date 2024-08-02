Attribute VB_Name = "Module1"
Sub Mail()

    If Range("C2") = "" Then
        x = Range("B2")
    Else:
        x = Range("C2")
    End If
    
    Select Case x
        Case "洪珮馨", "賴緯紘", "李慧貞", "林育嘉", "潘潔瑩", "吳美慧", "葉人瑄"
            cc = "陳怡雯; QA小組 <QA??@aptg.com.tw>"
        Case "許慧君", "林卓穎", "陳建成", "柯孟廷", "呂姿昀"
            cc = "楊姿璇"
        Case "張坤山", "曾琪婉", "王大衛", "劉志峰", "黃莉珊", "吳明昇"
            cc = "陳詠文"
    End Select
        
        
    Dim olApp As Object
    Dim olMailItm As Object
    
    Columns("E:H").Hidden = False
        
    Set olApp = CreateObject("Outlook.Application")
    Set olMailItm = olApp.CreateItem(0)
   
    With olMailItm
        .To = "CS-NH-Duty; CS-TY-Duty"
        .cc = cc
        .Subject = Range("C7") & "-【Chat留單煩請協助安排回電】" & Range("C4")
        .HTMLBody = RangetoHTML(Range("E1:H6"))
        .Display
    End With
       
    Set olMailItm = Nothing
    Set olApp = Nothing
    
    Columns("E:H").Hidden = True
    
End Sub

Sub 清除()

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

