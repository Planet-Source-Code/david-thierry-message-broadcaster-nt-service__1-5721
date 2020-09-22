Attribute VB_Name = "modMisc"
Option Explicit
Public Function PlayWav(ByVal fName As String)
  Dim PlayIt
    On Error Resume Next
     fName = App.Path & "\" & fName
     PlayIt = sndPlaySound(fName, 1)
End Function


Function CenterInWorkArea(frm As Form)
    Dim lNewTop As Long
    Dim lNewLeft As Long
    Dim WA As RECT
    Dim lReturn As Long
    'Get the work area in a RECTangle structure from the '
    'SystemParametersInfo API Call
    lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, WA, 0&)
    
    'Convert the virtual coordinates to scale coordinates
    WA.Left = WA.Left * Screen.TwipsPerPixelX
    WA.Right = WA.Right * Screen.TwipsPerPixelX
    WA.Top = WA.Top * Screen.TwipsPerPixelY
    WA.Bottom = WA.Bottom * Screen.TwipsPerPixelY
    
    
    'WA.Bottom-WA.Top = Work Area Height
    lNewTop = ((WA.Bottom - WA.Top - frm.Height) / 2) + WA.Top
    
    'Top is off screen or hidden because form is taller than workspac
    '     e; adjust
    If lNewTop < WA.Top Then lNewTop = WA.Top
    
    'WA.Right - WA.Left = Work Area Width
    lNewLeft = ((WA.Right - WA.Left - frm.Width) / 2) + WA.Left
    
    'Left is off screen or hidden because form is too wide for worksp
    '     ace; adjust
    If lNewLeft < WA.Left Then lNewLeft = WA.Left
    
    'Perfect Centering!
    frm.Move lNewLeft, lNewTop
End Function
Function AlignBottom(frm As Form)
    Dim lNewTop As Long
    Dim lNewLeft As Long
    Dim WA As RECT
    Dim lReturn As Long
    'Get the work area in a RECTangle structure from the '
    'SystemParametersInfo API Call
    lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, WA, 0&)
    
    'Convert the virtual coordinates to scale coordinates

    
    'make form the same width as the desktop
    frm.ScaleWidth = (WA.Right - WA.Left)
    
    'set the top position of the form
    
    lNewTop = WA.Bottom - frm.ScaleHeight
    
    'Top is off screen or hidden because form is taller than workspac
    '     e; adjust
    If lNewTop < WA.Top Then lNewTop = WA.Top
    
    'WA.Right - WA.Left = Work Area Width
    'lNewLeft = ((WA.Right - WA.Left - frm.Width) / 2) + WA.Left
    lNewLeft = 0
    
    'Left is off screen or hidden because form is too wide for worksp
    '     ace; adjust
    If lNewLeft < WA.Left Then lNewLeft = WA.Left
    
    'Perfect Centering!
    'frm.Move lNewLeft, lNewTop
    MoveWindow frm.hwnd, lNewLeft, lNewTop, frm.ScaleWidth, frm.ScaleHeight, 0
End Function
Function AlignTop(frm As Form)
    Dim lNewTop As Long, lNewLeft As Long
    Dim WA As RECT, lReturn As Long
    'Get the work area in a RECTangle structure from the '
    'SystemParametersInfo API Call
    lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, WA, 0&)
    

    'make form the same width as the desktop
    frm.ScaleWidth = (WA.Right - WA.Left)
    
    MoveWindow frm.hwnd, 0, 0, frm.ScaleWidth, frm.ScaleHeight, 0
End Function
Public Sub SetTopMost(frm As Form, TopMost As Boolean)
Dim Success As Long
If TopMost Then
  Success = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
Else
  Success = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End If
End Sub



