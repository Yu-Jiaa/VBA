Attribute VB_Name = "Module1"
Sub 未結案()

Worksheets(1).Name = "原始檔"
xLast = Columns("A").End(xlDown).Row

For x = 2 To xLast
    Range("Q" & x) = "=DATEVALUE(J" & x & ")"
    Range("Q" & x).NumberFormat = ("yyyy/mm/dd")
Next

Range("A1").AutoFilter Field:=17, Criteria1:="<" & Date
Range("A1").CurrentRegion.Copy
Worksheets.Add After:=Worksheets(1)
Worksheets(2).Name = "逾期未結案件"
Worksheets(2).Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False

Selection.Borders.LineStyle = xlContinuous
Range("A1:P1").Interior.Color = RGB(217, 217, 217)

With Cells
    .Font.Name = "微軟正黑體"
    .Font.Size = 12
    .EntireColumn.AutoFit
End With

Columns("D:F").Hidden = True
Columns("Q").Delete
Range("A1").Select

With Worksheets(1)
    .Columns("Q").Delete
    .AutoFilterMode = False
    '.Range("A1").Select
End With

ActiveWorkbook.Password = "16850"

End Sub


