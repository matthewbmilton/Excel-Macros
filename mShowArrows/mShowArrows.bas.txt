
'----------------------------------------
' Collection of macros to toggle viewing of Precedent/Dependency arrows in Excel.
' Useful for formula tracing/auditing.
' Benefits over built-in Excel functionality:
'   Select a range of cells, trace all dependencts/precedents in one click, instead of per-cell.
'   Faster when needing to audit a block of cells to ensure they all have similar precedents.
'   Easily visually identify any formula cell that has a different precedent, noticing that the arrow patterns are different.
'----------------------------------------
' To set up:
'   Right-click on Ribbon.
'   Customize Ribbon.
'   Under "Customize the Ribbon" on the right:
'     Option 1: Select "New Tab" and create new tab, with name.
'     Option 2: Select an existing tab, such as "Formulas".
'
'   Select "New Group" and type in a Group name for these macro buttons. Example, "Formula Auditing"
'
'   Select the Tab and Group, expanding all arrows.
'
'   Under "Choose commands from:" on the left, select "Macros"
'   Locate the macros (name may be similar to 'PERSONAL.XLSB!ShowPrecedents').
'   Select the "Add >>" button, and repeat for each of the three macros (ShowPrecedents, ShowDependencies, ClearArrows).
'
'   Select "Rename" and choose display-friendly names and icons.
' Recommended:
'   "Show Precedents" double-left arrow <<-
'   "Show Depenencies" double-right arrow ->>
'   "Clear Arrows": big red X
'----------------------------------------
' Author: Matthew B Milton
'----------------------------------------

Sub ShowPrecedents()
    For Each c In Selection
        c.ShowPrecedents
    Next c
End Sub

Sub ShowDependencies()
    For Each c In Selection
        c.ShowDependents
    Next c
End Sub

Sub ClearArrows()
    ActiveSheet.ClearArrows
End Sub
