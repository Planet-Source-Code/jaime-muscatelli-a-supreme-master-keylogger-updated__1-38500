Attribute VB_Name = "mdlActiveWindow"
Option Explicit

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public nCAPTION As String
Public nTESTER As Long
Public nClass As String

Public Sub GetActiveWindowName()
nCAPTION = Space(256)
nClass = Space(256)

GetWindowText GetForegroundWindow, nCAPTION, Len(nCAPTION)
GetClassName GetForegroundWindow, nClass, Len(nClass)

If nTESTER = GetForegroundWindow Then Exit Sub
FRMLOG.txtENUMERATE.Text = FRMLOG.txtENUMERATE.Text & vbCrLf & Time & " " & nCAPTION
FRMLOG.txtENUMERATE.Text = FRMLOG.txtENUMERATE.Text & vbTab & nClass
nTESTER = GetForegroundWindow
End Sub
