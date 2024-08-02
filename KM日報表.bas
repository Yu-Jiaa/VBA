Attribute VB_Name = "Module1"
Sub 日報表()
   Call 排序
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

   mStr = "Dear Duty" & vbCrLf & "請協助佈達 , 謝謝" & vbCrLf & "1.下表統計" & Day1 & _
        " KM上架資訊，提供人員知識編號、 KM位置及可使用哪些關鍵，可查詢到此則知識，請人員多加利用。" & vbCrLf & "2.已確認佈達內容皆已提供資料來源。"
   LF = Cells(Rows.Count, "A").End(xlUp).Row
   
   For x = 1 To LF
    If Range("A" & x) = Date - Days Then
        Exit For
    End If
   Next
   
   With olMailItm
       .To = "CS-NH-Duty; CS-TY-Duty"
       .Cc = "簡珮玲; 沈妤庭; 李文綺; 陳怡雯; QA小組 <QA??@aptg.com.tw>"
       .Subject = "【請協助佈達】KM上架日報表_" & Format(Date, "yyyymmdd")
       .HTMLBody = mStr & RangetoHTML(Range("A1:H1,A" & x & ":H" & LF))
       .Display
   End With
   
   'Clean up the Outlook application.
   Set olMailItm = Nothing
   Set olApp = Nothing
   
   ActiveWorkbook.Save
   
End Sub

'引用 Ron de Bruin 的 RangetoHTML 程序
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

Sub 排序()

With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("A1") ', SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range("B1"), CustomOrder:="作業流程,系統維護,行銷通報,客服App,教育訓練,用戶終端,防災廣播"
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

Sub 附件()
Call 排序
ActiveSheet.Copy
Columns("I:Q").Delete
Range("A1").Select
ActiveWorkbook.SaveAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\KM上架日報表(" & Worksheets(1).Name & ").xlsx"
'CreateObject("WScript.Shell").SpecialFolders("Desktop")  '桌面路徑

'ActiveWorkbook.SaveAs "C:\Users\yujialin\Desktop\KM上架日報表(" & Worksheets(1).Name & ").xlsx"

ActiveWorkbook.Close
MsgBox "已儲存在桌面！"

End Sub

Sub 複製()

Dim d As New DataObject

i = Selection.Row()
Str1 = "Hi Duty" & vbCrLf & vbCrLf & "更新完成如下路徑：" & vbCrLf & Range("E" & i) & vbCrLf & "知識編號：" & Range("C" & i) _
& vbCrLf & vbCrLf & "CS-NH-Duty <cs-nh-duty@aptg.com.tw>; CS-TY-Duty <cs-ty-duty@aptg.com.tw>" _
& vbCrLf & "陳怡雯 <evchen@aptg.com.tw>;  TQ小組 <TQ-Group@aptg.com.tw>; QA小組 <QA??@aptg.com.tw>" _
& vbCrLf & "【KM更新完成】"


'vbCrLf:換行

d.SetText Str1
d.PutInClipboard '放入剪貼簿

Range("J" & i).Interior.color = RGB(255, 0, 0) '設定底色(紅)
Application.Wait (Now + TimeValue("00:00:01")) '等候1秒
Range("J" & i).Interior.color = xlNone '再設定底色(無填滿)

End Sub

Sub Username()

'MsgBox Environ("Username")

MsgBox CreateObject("WScript.Shell").SpecialFolders("Desktop")

End Sub


