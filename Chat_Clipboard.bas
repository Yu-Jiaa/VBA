Attribute VB_Name = "Module2"
Sub �}��()

If Range("C1") = "���A�G�@��Ҧ�" Then
    Range("C1") = "���A�G�ƻs�Ҧ�"
Else
    Range("C1") = "���A�G�@��Ҧ�"
End If

End Sub


Sub cm1()

If Range("C1") = "���A�G�ƻs�Ҧ�" Then

    If Selection.Column() = "2" Then
        If Selection.Row() <> 1 Then
          
        i = Selection.Row()
        Dim data As New DataObject
    
        '�L�ϥΤu��
        str1 = Range("B" & i)
        data.SetText str1
        data.PutInClipboard  '��J�ŶKï
        
        Range("A" & i).Interior.Color = RGB(230, 184, 183) '�]�w����(��)
        Application.Wait (Now + TimeValue("00:00:01")) '����1��
        Range("A" & i).Interior.Color = xlNone '�A�]�w����(�L��)
        
    End If
    End If
    
End If

End Sub

Sub cm2()

'���ϥΤp����Ҧ�

If Range("C1") = "���A�G�ƻs�Ҧ�" Then

    If Selection.Column() = "2" Then
        If Selection.Row() <> 1 Then
          
        i = Selection.Row()
        
        '���ϥΤp�u��
        Range("B" & i).Copy
        
        Range("A" & i).Interior.Color = RGB(230, 184, 183) '�]�w����(��)
        Application.Wait (Now + TimeValue("00:00:01")) '����1��
        Range("A" & i).Interior.Color = xlNone '�A�]�w����(�L��)
        
    End If
    End If
    
End If

End Sub


