Attribute VB_Name = "mFormatting"
Sub AlignLeft()
Attribute AlignLeft.VB_ProcData.VB_Invoke_Func = "l\n14"
    With Selection
        .HorizontalAlignment = xlLeft
    End With
End Sub

Sub AlignCenter()
Attribute AlignCenter.VB_ProcData.VB_Invoke_Func = "e\n14"
    With Selection
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub AlignRight()
Attribute AlignRight.VB_ProcData.VB_Invoke_Func = "r\n14"
    With Selection
        .HorizontalAlignment = xlRight
    End With
End Sub

Sub AlignCenterAcrossSelection()
Attribute AlignCenterAcrossSelection.VB_ProcData.VB_Invoke_Func = "E\n14"
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
End Sub
