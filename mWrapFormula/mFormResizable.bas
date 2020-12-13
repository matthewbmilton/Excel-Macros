Attribute VB_Name = "mFormResizable"
'Code from here: https://www.teachexcel.com/excel-tutorial/2027/resizable-userformFormResizableFormResizable

Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As Long

Private Declare PtrSafe Function GetWindowLong _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) _
  As Long

Private Declare PtrSafe Function SetWindowLong _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) _
  As Long

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

Public Sub FormResizable()

Dim lStyle As Long
Dim hWnd As Long
Dim RetVal

hWnd = GetForegroundWindow

lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Sub
