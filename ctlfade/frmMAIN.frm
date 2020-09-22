VERSION 5.00
Begin VB.Form frmMAIN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Form"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMAIN.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDEMO 
      Caption         =   "&Click me"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblLINK 
      BackStyle       =   0  'Transparent
      Caption         =   "http://vbasic.iscool.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMAIN.frx":B38E
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------
' For more of my stuff, you can visit my site
' at http://vbasic.iscool.net
'------------------------------------------------------

Private Sub cmdDEMO_Click()
    '------------------------------------------------------
    ' Call the sub to fade out the control.
    '------------------------------------------------------
    FadeOutControl cmdDEMO, Me
End Sub
Private Sub Form_Click()
    '------------------------------------------------------
    ' Show the button again.
    '------------------------------------------------------
    Me.cmdDEMO.Visible = True
End Sub
Private Sub Form_Load()
    On Error Resume Next
    '------------------------------------------------------
    ' make the background image fit the form.
    '------------------------------------------------------
    Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
Private Sub Label1_Click()
    '------------------------------------------------------
    ' Show the button again.
    '------------------------------------------------------
    Me.cmdDEMO.Visible = True
End Sub
Private Sub lblLINK_Click()
    On Error Resume Next
    '------------------------------------------------------
    ' Go to my web site.
    '------------------------------------------------------
    ExecuteLink lblLINK.Caption
End Sub
