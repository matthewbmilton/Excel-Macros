Sub InsertRows()
'----------------------------------------
' Select a row. Run macro (recommended: from button on ribbon/Quick Access Toolbar)
' Enter a number
' Insert that many rows
'----------------------------------------
' Author: Matthew B Milton
'----------------------------------------
    Dim NumRows As Integer
    NumRows = InputBox("Number of rows?")

    For i = 1 To NumRows
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i

End Sub
