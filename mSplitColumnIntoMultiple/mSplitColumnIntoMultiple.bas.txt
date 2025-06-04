Sub SplitColumnIntoMuliple()
'Takes a long Column A,
'Copies it into a number of columns
'Based on how many items per column defined in SplitCount
' Outputs to the right of the selected column
'
'Example:
'   Col A has 13,000 items
'   Need to split into separate columns every 900 items
'   This will need 15 columns (14.4 rounded up)
'----------------------------------------
' Author: Matthew B Milton
'----------------------------------------
    
    'Declarations
        Dim RowCount As Long     'How many total rows
        Dim ColCount As Long     'Determine how many COLs are needed
        Dim SplitCount As Long   'How many items per column
        Dim SourceRange As Range 'Selection of where to copy from
        Dim CopyRange As Range   'Sub-selection of where to copy from, by chunk
        Dim RowStart             'Where to start copying from
        Dim RowEnd               'Where to stop copying from
        Dim iDefaultRows As Long 'How many rows to split into, by default
        
    'Initialize
        iDefaultRows = 999
    
    'Set SourceRange as current selection
        Set SourceRange = Selection
    
    'Input: How many rows per column?
        SplitCount = InputBox( _
                        "How many rows per column?", _
                        "SplitCount", _
                        iDefaultRows _
                        )
        RowCount = Selection.Rows.count
    
    'How many columns do we split this into?
        ColCount = _
            WorksheetFunction.Ceiling( _
                Arg1:=RowCount / SplitCount, _
                Arg2:=1 _
                )
    
    'Loop to fill in each column
        For i = 1 To ColCount
            'Start and end rows, based on current destination column
                RowStart = ((i - 1) * SplitCount) + 1
                RowEnd = (i) * SplitCount
            
            'Select the sub-range to copy from
                Set CopyRange = SourceRange.Range(Cells(RowStart, 1), Cells(RowEnd, 1))
            
            'Copy to the destination: rows 1 to SplitCount,
            'start in column 1 to the right of the source range
                CopyRange.Copy _
                    Destination:= _
                        SourceRange.Range(Cells(1, i + 1), Cells(SplitCount, i + 1))
    
        Next i

End Sub


