Attribute VB_Name = "basCOMMON"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Public Sub ExecuteLink(LINK As String)
    On Error Resume Next
    Dim lRet As Long
    If LINK <> "" Then
        lRet = ShellExecute(0, "open", LINK, "", App.Path, SW_SHOWNORMAL)
        If lRet >= 0 And lRet <= 32 Then
            MsgBox Err.Number & ":" & Err.Description, vbCritical, "Error"
        End If
    End If
End Sub
Public Sub MakeTransparent(LhWnd As Long, bLevel As Byte)
    Dim lOldStyle As Long
    lOldStyle = GetWindowLong(LhWnd, GWL_EXSTYLE)
    SetWindowLong LhWnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes LhWnd, 0, bLevel, LWA_ALPHA
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever
' Date: January,02 2003 @ 10:58:09
'------------------------------------------------------------
Public Sub FadeOutControl(ctl As Object, pForm As Form)
    On Error GoTo ErrorFadeControl
    Dim frm As frmFADE, tc As Long, x As Long
    Dim oX As Long, oY As Long
    Dim r As RECT
    '------------------------------------------------------
    ' Store the postion of the control
    '------------------------------------------------------
    oX = ctl.Left
    oY = ctl.Top
    '------------------------------------------------------
    ' Create an instance of frmFADE
    '------------------------------------------------------
    Set frm = New frmFADE
    '------------------------------------------------------
    ' Get the screen position of the contorl
    '------------------------------------------------------
    GetWindowRect ctl.hwnd, r
    '------------------------------------------------------
    ' Move our instance of frmFADE to the same
    ' place as the control
    '------------------------------------------------------
    MoveWindow frm.hwnd, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, False
    '------------------------------------------------------
    ' Assign the instance of frmFADE as the new
    ' parent of the control
    '------------------------------------------------------
    SetParent ctl.hwnd, frm.hwnd
    '------------------------------------------------------
    ' Move the control to 0,0 on frmFADE
    '------------------------------------------------------
    ctl.Move 0, 0
    '------------------------------------------------------
    'Show the instance of frmFADe
    '------------------------------------------------------
    frm.Show , pForm
    pForm.SetFocus
    '------------------------------------------------------
    ' Loop backwards slowly making frmFADE
    ' fade away.
    '------------------------------------------------------
    For x = 255 To 0 Step -5
        MakeTransparent frm.hwnd, CByte(x)
        tc = GetTickCount
        While GetTickCount < tc + 1: Sleep 1: DoEvents: Wend
    Next x
    '------------------------------------------------------
    ' Once 100% transparent hide the form
    '------------------------------------------------------
    ctl.Visible = False
    '------------------------------------------------------
    ' Move the control back to the orginal form
    '------------------------------------------------------
    SetParent ctl.hwnd, pForm.hwnd
    '------------------------------------------------------
    ' Make sure it is in the same spot as before.
    '------------------------------------------------------
    ctl.Move oX, oY
    '------------------------------------------------------
    ' Unload our instance of frmFADE
    '------------------------------------------------------
    Unload frm
    Exit Sub
ErrorFadeControl:
    MsgBox Err & ":Error in call to FadeControl()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
