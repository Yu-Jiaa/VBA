Attribute VB_Name = "Module1"
Sub 巨集1()
Attribute 巨集1.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Rng As Range '自動篩選結果範圍
Dim Book1 As Workbook
Set Book1 = ActiveWorkbook

'抓檔名
Filename = Book1.Name
Filename = "CCQ010(" & Mid(Filename, 8, 7) & ")"

'篩選CCQ010並複製
With Book1.ActiveSheet
    Set Rng = .UsedRange
    Rng.AutoFilter Field:=23, Criteria1:="=*CCQ010*", Operator:=xlAnd
    Set Rng = .UsedRange
    Rng.Copy
End With

'貼到新活頁簿
Workbooks.Add
ActiveSheet.Paste

'排序(CSR>時間>門號)
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

'門號及時間-重複資料註記
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

'門號及時間-調整欄寬
Columns("B:B").EntireColumn.AutoFit
Columns("O:O").EntireColumn.AutoFit

Range("A1").Select

'輸入密碼&另存檔案&關閉檔案
ActiveSheet.Name = Filename & "原始版"
ActiveSheet.SaveAs Book1.Path & "\" & Filename & ".xlsx"

'DataPass= InputBox("請輸入要設定的密碼","設定檔案開啟密碼")
ActiveWorkbook.Password = "16850"
ActiveWorkbook.Save

Book1.Close False 'False:放棄在此活頁簿中所做的任何變更

End Sub

