VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Windows XP Command Buttons"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin XPUI.XPCommandButton XPCommandButton2 
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPUI.XPCommandButton XPCommandButton1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1508
      Caption         =   "Click to enable/disable Close button"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Events"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
      Begin VB.Label e 
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
XPCommandButton1.Width = Width - 360
With XPCommandButton2
.Top = Height - .Height - 600
.Left = Width - .Width - 240
End With
End Sub

Private Sub XPCommandButton1_Click()
e.Caption = "Click"
XPCommandButton2.Enabled = Not XPCommandButton2.Enabled
If XPCommandButton2.Enabled Then XPCommandButton1.Caption = "Click to disable Close button" Else XPCommandButton1.Caption = "Click to enable Close button"
End Sub

Private Sub XPCommandButton1_GotFocus()
e.Caption = "GotFocus"
End Sub

Private Sub XPCommandButton1_LostFocus()
e.Caption = "LostFocus"
End Sub

Private Sub XPCommandButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
e.Caption = "MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
End Sub

Private Sub XPCommandButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
e.Caption = "MouseMove(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
End Sub

Private Sub XPCommandButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
e.Caption = "MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
End Sub

Private Sub XPCommandButton2_Click()
Unload Me
End Sub
