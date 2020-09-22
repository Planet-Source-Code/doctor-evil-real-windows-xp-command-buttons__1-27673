VERSION 5.00
Begin VB.UserControl XPCommandButton 
   Appearance      =   0  'Flat
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   0  'User
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   3480
   End
   Begin VB.CommandButton StandardButton 
      Caption         =   "XPCommandButton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "XPCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WINDOWS XP COMMAND BUTTON CONTROL.
' (c) 2001 dr.-evil@mad.scientist.com.  All rights reserved.
' You may use this control in your applications free of charge,
' provided that you do not redistribute this source code without
' giving me credit for my work.  Of course, credit in your
' applications is always welcome.

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim MyState As BtnState
Dim LastState As BtnState
Dim MouseIsDown As Boolean
Dim IsKeyDown As Boolean

Dim s_Enabled As Boolean
Dim s_Caption As String
Dim s_Font As Font

Public Property Get Caption() As String
Caption = s_Caption
End Property

Public Property Get Enabled() As Boolean
Enabled = s_Enabled
End Property

Public Property Get Font() As Font
Set Font = s_Font
End Property

Public Property Let Caption(Val As String)
s_Caption = Val
StandardButton.Caption = Val
Draw True   ' force a redraw so the that the new caption is shown.
End Property

Public Property Let Enabled(Val As Boolean)
s_Enabled = Val
If Not Val Then
MyState = Disabled
UserControl.Enabled = False
ElseIf Val And (UserControl.Ambient.DisplayAsDefault) Then
MyState = Defaulted
UserControl.Enabled = True
Else
MyState = Normal
UserControl.Enabled = True
End If
Draw
StandardButton.Enabled = Val
End Property

Public Property Set Font(Val As Font)
Set s_Font = Val
Set UserControl.Font = Val
Set StandardButton.Font = Val
Draw
End Property

Private Sub StandardButton_Click()
RaiseEvent Click
End Sub

Private Sub StandardButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub StandardButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub StandardButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()    ' this provides for a MouseOut event.
If Not MouseIsDown Then
    Dim Mouse As POINT_API
    GetCursorPos Mouse
    ScreenToClient hWnd, Mouse
    If (Mouse.X < UserControl.ScaleLeft) Or (Mouse.Y < UserControl.ScaleTop) Or (Mouse.X > (UserControl.ScaleLeft + UserControl.ScaleWidth)) Or Mouse.Y > ((UserControl.ScaleTop + UserControl.ScaleHeight)) Then
        Timer1.Enabled = False
        If UserControl.Ambient.DisplayAsDefault Then
        MyState = Defaulted
        Else
        MyState = Normal
        End If
        Draw
    End If
End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If s_Enabled Then
    RaiseEvent Click
    Else
    Beep
    End If
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If Not ThemesSupported Then                                         ' this event usually happens when another control on the form gets focus.
StandardButton.Default = UserControl.Ambient.DisplayAsDefault
ElseIf PropertyName = "DisplayAsDefault" Then
If (MyState = Normal) And UserControl.Ambient.DisplayAsDefault Then
MyState = Defaulted
ElseIf (MyState = Defaulted) And Not UserControl.Ambient.DisplayAsDefault Then
MyState = Normal
End If
Draw
End If
End Sub

Private Sub UserControl_GotFocus()
If Not s_Enabled Then Beep
End Sub

Private Sub UserControl_Initialize()    ' check if the system supports themes.
If ThemesSupported Then                 ' if so, then get ready to use them.
StandardButton.Visible = False          ' if not, then show the standard command button.
MyState = Normal
Else
With StandardButton
.Visible = True
.Width = Width
.Height = Height
End With
End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 32) And s_Enabled Then
IsKeyDown = True
MyState = Pressed
Draw
End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = 32) And s_Enabled Then
IsKeyDown = False
MyState = Defaulted
RaiseEvent Click
Draw
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not IsKeyDown Then
If (MyState <> Disabled) And (Button = 1) Then
MyState = Pressed
Draw
End If
MouseIsDown = True
If s_Enabled Then RaiseEvent MouseDown(Button, Shift, X, Y)
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not IsKeyDown Then
If Button = 1 And Not MouseIsDown Then
MyState = MouseOver
Draw
End If
If (MyState <> Disabled) And (MyState <> Pressed) And (Button = 0) Then
MyState = MouseOver
Draw
ElseIf (MyState <> Disabled) And ((X > UserControl.ScaleWidth) Or (Y > UserControl.ScaleHeight) Or (X < 0) Or (Y < 0)) Then
MyState = MouseOver
Draw
ElseIf (MyState <> Disabled) And (Button = 1) Then
MyState = Pressed
Draw
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
If s_Enabled And ThemesSupported Then Timer1.Enabled = True
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not IsKeyDown Then
If MyState = MouseOver Then
MyState = Defaulted
Draw
ElseIf MyState = Pressed Then
MyState = MouseOver
Draw
End If
If (X >= 0) And (Y >= 0) And (X <= UserControl.ScaleWidth) And (Y <= UserControl.ScaleHeight) And (Button = 1) Then RaiseEvent Click
MouseIsDown = False
If s_Enabled Then RaiseEvent MouseUp(Button, Shift, X, Y)
End If
End Sub

Private Sub UserControl_Paint()
Draw True
End Sub

Private Sub UserControl_Resize()
If ThemesSupported Then
UserControl.ScaleMode = 3
Draw
Else
UserControl.ScaleMode = 1
With StandardButton
.Width = UserControl.Width
.Height = UserControl.Height
End With
End If
End Sub

Private Sub Draw(Optional Force As Boolean)
If ThemesSupported Then
Dim MyRECT As RECT
MyRECT.Top = 0
MyRECT.Left = 0
MyRECT.Right = UserControl.ScaleWidth
MyRECT.bottom = UserControl.ScaleHeight
If MyState = Normal And UserControl.Ambient.DisplayAsDefault Then MyState = Defaulted
If (MyState <> LastState) Or Force Then  ' if we check this first, it can prevent ALOT of flickering.
DrawButton UserControl.hWnd, UserControl.hDC, MyRECT, s_Caption, MyState
LastState = MyState
End If
Else
StandardButton.Visible = True
End If
End Sub




Private Sub UserControl_InitProperties()
s_Caption = "XPCommandButton"
s_Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Me.Caption = PropBag.ReadProperty("Caption", "XPCommandButton")
Me.Enabled = PropBag.ReadProperty("Enabled", True)
Set Me.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", s_Caption, "XPCommandButton"
PropBag.WriteProperty "Enabled", s_Enabled, True
PropBag.WriteProperty "Font", s_Font, UserControl.Ambient.Font
End Sub

