'These short macros mimic the MS Word shortcuts for formatting Left/Right/Center
'You need to set up the shortcut keys yourself.
'Import this code into your Personal.xlsb workbook (or other location)
'In Excel, go to the Developer tab.
'	Select 'Macros' from the Ribbon
'	Select a macro
'	Select the 'Options...' button
'	Assign a Shortcut Key
'Recommended shortcut keys:
'	Ctrl+l = AlignLeft
'	Ctrl+e = AlighCenter
'	Ctrl+r = AlignRight
'	Ctrl+Shift+e = AlignCenterAcrossSelection
'----------------------------------------
' Author: Matthew B Milton
'----------------------------------------

Sub AlignLeft()
    With Selection
        .HorizontalAlignment = xlLeft
    End With
End Sub

Sub AlignCenter()
    With Selection
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub AlignRight()
    With Selection
        .HorizontalAlignment = xlRight
    End With
End Sub

Sub AlignCenterAcrossSelection()
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
End Sub
