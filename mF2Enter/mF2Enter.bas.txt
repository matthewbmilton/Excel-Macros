Sub F2Enter()
'Mimics pressing "<F2>, <Enter>" for each cell in selection
'Useful to force update formats, turn text into URLs, or numbers to text within Text-formatted fields
'----------------------------------------
' Author: Matthew B Milton
'----------------------------------------
    Dim i As Long
    i = 1

    For Each c In Selection
        c.Value = c.Value & ""
        i = i + 1
        Debug.Print i
        DoEvents
    Next c

End Sub
