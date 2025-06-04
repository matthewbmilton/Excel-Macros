Sub ColWidth()
'Takes each column and adjusts its width up to the nearest integer
'----------------------------------------
' Author: Matthew B Milton
'----------------------------------------

    For Each col In Selection.Columns
        col.ColumnWidth = WorksheetFunction.RoundUp(col.ColumnWidth, 0)
        Debug.Print col.ColumnWidth
        DoEvents
    Next

End Sub
