VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmController 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmController.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin NTService.NTService NTService1 
      Left            =   1440
      Top             =   2640
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "SCI Broadcasting Service"
      ServiceName     =   "SCI Broadcasting Service"
      StartMode       =   3
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   2640
   End
   Begin VB.Menu PopUp 
      Caption         =   "PopUp"
      Begin VB.Menu SendMessage 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuSendToSelf 
         Caption         =   "Send To Self"
         Checked         =   -1  'True
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlignBottom 
         Caption         =   "Align Bottom"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAlignTop 
         Caption         =   "Align Top"
         Enabled         =   0   'False
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoopMessages 
         Caption         =   "Loop Messages"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IconObject As Object
Dim Counter As Integer



Private Sub Initialize()
      boolLoopMessages = True
      boolSendToSelf = True
     'set up a tray icon and a message hook back to this form
      Set IconObject = Me.Icon
      AddIcon Me, IconObject.Handle, IconObject, "Animated TrayIcon"
     
     'load the broadcast form but don't show it
      Load frmBroadcast
      
     'load the messages form but don't show it
      Load frmMessages
End Sub
Private Sub Form_Load()
On Error GoTo Err_Load
    Dim strDisplayName As String
    Dim bStarted As Boolean
    
    strDisplayName = NTService1.DisplayName
    If Command = "-install" Then
        ' enable interaction with desktop
        NTService1.Interactive = True
        Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, " Installing Services")
        If NTService1.Install Then
            Call NTService1.SaveSetting("Parameters", "TimerInterval", "1000")
            MsgBox strDisplayName & " installed successfully"
        Else
            MsgBox strDisplayName & " failed to install"
        End If
        End
    ElseIf Command = "-uninstall" Then
        Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, " Un-installing Services")
        If NTService1.Uninstall Then
            MsgBox strDisplayName & " uninstalled successfully"
        Else
            MsgBox strDisplayName & " failed to uninstall"
        End If
        End
    ElseIf Command = "-standalone" Then
        mnuExit.Enabled = True
        Initialize
        Exit Sub
    End If

    ' enable Pause/Continue. Must be set before StartService
    ' is called or in design mode
    NTService1.ControlsAccepted = svcCtrlPauseContinue
    
  
    ' connect service to Windows NT services controller
    Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, " Starting Services")
    NTService1.StartService
    
    Exit Sub
Err_Load:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Message As Long
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    Case WM_RBUTTONUP:
        Me.PopupMenu PopUp
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'unload the messages form
     Unload frmMessages
     
    'unload the broadcast form
     Unload frmBroadcast
     
    'remove the app icon from the system tray
     delIcon IconObject.Handle
     delIcon Me.Icon.Handle
End Sub

Private Sub mnuAlignBottom_Click()
  If mnuAlignBottom.Checked Then Exit Sub
  AlignBottom frmMessages
  mnuAlignBottom.Checked = True
  mnuAlignTop.Checked = False
End Sub

Private Sub mnuAlignTop_Click()
  If mnuAlignTop.Checked Then Exit Sub
  AlignTop frmMessages
  mnuAlignBottom.Checked = False
  mnuAlignTop.Checked = True
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuLoopMessages_Click()
  mnuLoopMessages.Checked = Not mnuLoopMessages.Checked
  boolLoopMessages = mnuLoopMessages.Checked
End Sub

Private Sub mnuSendToSelf_Click()
  mnuSendToSelf.Checked = Not mnuSendToSelf.Checked
  boolSendToSelf = mnuSendToSelf.Checked
End Sub

Private Sub SendMessage_Click()
  frmBroadcast.Show
End Sub

Private Sub Timer1_Timer()
'    Counter = Counter + 1
'    Form1.Icon = ImageList1.ListImages(Counter).Picture
'    If Counter > ImageList1.ListImages.Count - 1 Then Counter = 0
'    modIcon Form1, IconObject.Handle, Form1.Icon, "Animated TrayIcon"
End Sub

'------------------------------------------------------------------------------------
'Continue Event - Fires off when the Continue button is clicked in the Service
'Control Manager.
'------------------------------------------------------------------------------------
Private Sub NTService1_Continue(Success As Boolean)
On Error GoTo Err_Continue
    Success = True
    Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "Service continued")
    Exit Sub
Err_Continue:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub

'------------------------------------------------------------------------------------
'Pause Event - Fires off when the Pause button is clicked in the Service
'Control Manager.
'------------------------------------------------------------------------------------
Private Sub NTService1_Pause(Success As Boolean)
On Error GoTo Err_Pause
    Call NTService1.LogEvent(svcEventError, svcMessageError, "Service paused")
    Success = True
    Exit Sub
Err_Pause:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub
'------------------------------------------------------------------------------------
'Start Event - Fires off when the Start button is clicked in the Service
'Control Manager.
'------------------------------------------------------------------------------------
Private Sub NTService1_Start(Success As Boolean)
On Error GoTo Err_Start
    Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "[Start Event] " & "Starting Service")
    Initialize
    Success = True
    Exit Sub
Err_Start:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    Success = False
End Sub
'------------------------------------------------------------------------------------
'Stop Event - Fires off when the Stop button is clicked in the Service
'Control Manager.
'------------------------------------------------------------------------------------
Private Sub NTService1_Stop()
On Error GoTo Err_Stop
  Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "[Stop Event] " & "Stopping Service")
  Unload Me
  Exit Sub
Err_Stop:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


