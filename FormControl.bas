Attribute VB_Name = "FormControl"












Public Sub ResizeControls_add(frm As Form)
Dim i As Integer
x_size_add = frm.height / iHeight_add
y_size_add = frm.width / iWidth_add

For i = 0 To UBound(list_add)
    For Each curr_obj In frm
        If list_add(i).support = "supported" Then
            With curr_obj
                .Left = list_add(i).Left * y_size_add
                .width = list_add(i).width * y_size_add
                
                .Top = list_add(i).Top * x_size_add
                On Local Error GoTo Hata:
                .height = list_add(i).height * x_size_add
Hata:
             End With
         End If
    Next curr_obj
Next i
End Sub


Public Sub GetLocation_add(frm As Form)
Dim i As Integer
i = 0
For Each curr_obj In frm
    ReDim Preserve list_add(i)

    Select Case curr_obj.Name
        Case "Timer1"
            With list_add(i)
                '.support = "unsupported"
            End With
        Case Else
            With list_add(i)
                .Index = curr_obj.TabIndex
                .Left = curr_obj.Left
                .Top = curr_obj.Top
                .width = curr_obj.width
                .height = curr_obj.height
                .support = "supported"
            End With
        i = i + 1
    End Select

    

Next curr_obj
'   This is what the object sizes will be compared to on rescaling.
    iHeight_add = frm.height
    iWidth_add = frm.width
End Sub



