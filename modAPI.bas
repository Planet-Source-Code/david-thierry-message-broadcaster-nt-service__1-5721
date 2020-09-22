Attribute VB_Name = "modAPI"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
      
Public Const SPI_GETWORKAREA = 48
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
