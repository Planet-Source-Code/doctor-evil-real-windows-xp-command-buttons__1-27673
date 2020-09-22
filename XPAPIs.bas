Attribute VB_Name = "XPAPIs"

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Type POINT_API
    X As Long
    Y As Long
End Type

Public Enum BtnState
Defaulted = PBS_DEFAULTED
Disabled = PBS_DISABLED
MouseOver = PBS_HOT
Normal = PBS_NORMAL
Pressed = PBS_PRESSED
End Enum

Public Sub DrawButton(hWnd As Long, hDC As Long, DestRect As RECT, Caption As String, State As BtnState)
'This draws out the button, then button the text over it.
    Dim hTheme As Long
    hTheme = OpenThemeData(hWnd, "BUTTON")
    DrawThemeBackground hTheme, hDC, BP_PUSHBUTTON, CLng(State), DestRect, ByVal 0&
    DrawThemeText hTheme, hDC, BP_PUSHBUTTON, CLng(State), Caption, -1, DT_CENTER Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_SINGLELINE, 0, DestRect
    CloseThemeData hTheme
End Sub

Public Function ThemesSupported() As Boolean
' First, we make sure that the UXTHEME.DLL file exsists.
' Then, we call 2 APIs that make sure that the current app is supposed to use themes.
If CheckForDLL Then
If CheckForThemes Then ThemesSupported = True
End If
End Function

Private Function CheckForDLL() As Boolean
' This sees if the UXTHEME.DLL exsists, meaning that it is XP or greater.
    Dim hLib As Long
    hLib = LoadLibrary("uxtheme.dll")
    If hLib <> 0 Then FreeLibrary hLib
    CheckForDLL = Not (hLib = 0)
End Function

Private Function CheckForThemes() As Boolean
' If UXTHEME.DLL exsists, this function checks if we should really use themes.
' There are 2 cases we wouldn't:
'   (1) The user set the apperance back to Windows Classic Style.
'   (2) This program is running in compatibility mode with visual themes disabled.
    If CBool(IsAppThemed) And CBool(IsThemeActive) Then CheckForThemes = True
End Function
