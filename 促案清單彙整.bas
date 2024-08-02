Attribute VB_Name = "Module1"
Public xMonth

Sub 促案()

xMonth = InputBox("請輸入製作年份及月分" & vbNewLine & "例：202305")
'xMonth = "202304"
Worksheets("工作表1").Name = "所有促案清單"

'插入【A欄】-常用，比對(VlookUp) 【續約】的【O欄】
Columns("A").Insert
Range("A1") = "常用"

xLast = Columns("B").End(xlDown).Row
For x = 2 To xLast
    Range("A" & x) = "=VLOOKUP(D" & x & ",續約!O:O,1,0)"
Next

'隱藏欄位【F~I、K、O~AF、AI~AL、AN~AO】
Range("F:I, K:K, O:AF, AI:AL, AN:AO").EntireColumn.Hidden = True

'【篩選並複製】有比對到的資料(不等於#N/A)
Range("A1").AutoFilter Field:=1, Criteria1:="<>#N/A"
Columns("C:AM").Copy

'貼上新工作表【X月續約促案清單】
Worksheets.Add Before:=Worksheets("所有促案清單")
Worksheets(1).Name = Right(xMonth, 1) & "月續約促案清單"
ActiveSheet.Paste Range("A1")
Application.CutCopyMode = False

'【所有促案清單】工作表-解除【篩選】。
Worksheets("所有促案清單").AutoFilterMode = False

'【X月續約促案清單】工作表
'排序：【產品別】由A到Z排序
Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

'複製【A~D欄】到新工作表【X月續約促案代碼】
xLast = Columns("A").End(xlDown).Row
Range("A1:D" & xLast).Copy
Worksheets.Add Before:=Worksheets(1)
Worksheets(1).Name = Right(xMonth, 1) & "月續約促案代碼"
ActiveSheet.Paste Range("A1")
Application.CutCopyMode = False

'【續約】工作表
' 將【BG欄】(續約專案平均實收月租) 【轉換成數字】
xLast = Columns("BG").End(xlDown).Row
For x = 2 To xLast
   Worksheets("續約").Range("BG" & x) = Worksheets("續約").Range("BG" & x).Value
Next

Call 樞紐分析表

End Sub

Sub test()

End Sub


Public Sub 樞紐分析表()

Dim PTCache As PivotCache
Dim PT As PivotTable

    Set PTCache = ThisWorkbook.PivotCaches.Add _
        (SourceType:=xlDatabase, SourceData:=Sheets("續約").Range("A1").CurrentRegion.Address)

    Set PT = PTCache.CreatePivotTable(TableDestination:="", TableName:="樞紐分析表1")
         
    ActiveSheet.Name = "樞紐"
    
    With ActiveSheet.PivotTables("樞紐分析表1")
        .PivotFields("是否取消").Orientation = xlPageField '報表篩選
        .PivotFields("是否取消").CurrentPage = "未取消"
        .PivotFields("專案代碼").Orientation = xlRowField '列標籤
        .PivotFields("專案名稱").Orientation = xlRowField '列標籤"
        .PivotFields("續約專案平均實收月租").Orientation = xlColumnField '欄標籤
        .PivotFields("續約專案平均實收月租").AutoSort xlAscending, "續約專案平均實收月租" '排序
        .AddDataField ActiveSheet.PivotTables("樞紐分析表1").PivotFields("門號"), "續約成交促案排名統計", xlCount '值
        .RowAxisLayout xlTabularRow '以列表方式顯示
        .PivotFields("專案代碼").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)  '關閉小計
    End With

    Worksheets("樞紐").Range("A3").CurrentRegion.Copy
    Worksheets.Add Before:=Worksheets(1)
    Worksheets(1).Name = Right(xMonth, 1) & "月續約促案排名"
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

End Sub

Sub Sav()

MsgBox (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & xMonth & "客服續約促案清單")
ActiveWorkbook.SaveAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & xMonth & "客服續約促案清單.xlsm"
'CreateObject("WScript.Shell").SpecialFolders("Desktop")  '桌面路徑

End Sub

