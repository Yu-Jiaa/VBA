Attribute VB_Name = "Module2"
Sub 開關()

If Range("C1") = "狀態：一般模式" Then
    Range("C1") = "狀態：複製模式"
Else
    Range("C1") = "狀態：一般模式"
End If

End Sub


Sub cm1()

If Range("C1") = "狀態：複製模式" Then

    If Selection.Column() = "2" Then
        If Selection.Row() <> 1 Then
          
        i = Selection.Row()
        Dim data As New DataObject
    
        '無使用工具
        str1 = Range("B" & i)
        data.SetText str1
        data.PutInClipboard  '放入剪貼簿
        
        Range("A" & i).Interior.Color = RGB(230, 184, 183) '設定底色(紅)
        Application.Wait (Now + TimeValue("00:00:01")) '等候1秒
        Range("A" & i).Interior.Color = xlNone '再設定底色(無填滿)
        
    End If
    End If
    
End If

End Sub

Sub cm2()

'有使用小幫手模式

If Range("C1") = "狀態：複製模式" Then

    If Selection.Column() = "2" Then
        If Selection.Row() <> 1 Then
          
        i = Selection.Row()
        
        '有使用小工具
        Range("B" & i).Copy
        
        Range("A" & i).Interior.Color = RGB(230, 184, 183) '設定底色(紅)
        Application.Wait (Now + TimeValue("00:00:01")) '等候1秒
        Range("A" & i).Interior.Color = xlNone '再設定底色(無填滿)
        
    End If
    End If
    
End If

End Sub


