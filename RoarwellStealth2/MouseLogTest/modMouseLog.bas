Attribute VB_Name = "modMouseLog"
Option Explicit
'WindowFromPoint determines the handle of the window located at a specific point on the screen.
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'GetCursorPos reads the current position of the mouse cursor. _
The x and y coordinates of the cursor (relative to the screen) _
are put into the variable passed as lpPoint. The function _
returns 0 if an error occured or 1 if it is successful.
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_TYPE) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'The POINT_TYPE structure holds the (x,y) coordinate of a point
Public Type POINT_TYPE
  x As Long
  y As Long
End Type

Public Function PollMouse() As String
' Display the title bar text of whatever window the mouse
' cursor is currently over.  Note that this could be a control on a program window.
Dim mousepos As POINT_TYPE  ' coordinates of the mouse cursor
Dim wintext As String, slength As Long  ' receive title bar text and its length
Dim hwnd As Long  ' handle to the window found at the point
Dim retval As Long  ' return value

' Determine the window the mouse cursor is over.
retval = GetCursorPos(mousepos)  ' get the location of the mouse
hwnd = WindowFromPoint(mousepos.x, mousepos.y)  ' determine the window that's there
If hwnd = 0 Then  ' error or no window at that point
  'Debug.Print "No window exists at that location."
  'End
  PollMouse = Chr(vbNull)
End If

' Display that window's title bar text
slength = GetWindowTextLength(hwnd)  ' get length of title bar text
wintext = Space(slength + 1)  ' make room in the buffer to receive the string
slength = GetWindowText(hwnd, wintext, slength + 1)  ' get the text
wintext = Left(wintext, slength)  ' extract the returned string from the buffer
PollMouse = "Title bar text of the window: " & wintext
End Function


