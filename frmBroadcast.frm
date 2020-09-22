VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBroadcast 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Broadcast Message"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameComputers 
      Caption         =   "Broadcast to Whom?"
      Height          =   4215
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtFriendlyName 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtComputerName 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
         Width           =   2055
      End
      Begin MSComctlLib.ListView lbComputerNames 
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4471
         View            =   2
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Friendly Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Host Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   855
      End
   End
   Begin VB.Frame framMessage 
      Caption         =   "Message to Broadcast"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox chkCloseForm 
         Caption         =   "Close Form on Send"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtMessage 
         Height          =   2655
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Erase"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   3720
         Width           =   1095
      End
   End
   Begin MSWinsockLib.Winsock WSServer 
      Left            =   7680
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmBroadcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim AreChanges As Boolean
Dim strMyName As String

Function ComputerExists(ByVal aComputer As String) As Boolean
Dim li As ListItem
  aComputer = Trim(LCase(aComputer))
  For Each li In lbComputerNames.ListItems
    If LCase(li.Tag) = aComputer Then
      ComputerExists = True
      Exit Function
    End If
  Next li
  ComputerExists = False
End Function
Private Sub cmdAdd_Click()
'Dim li As ListItem
'  'Load frmNewComputer
'   Me.Hide
'   frmNewComputer.Show vbModal
'   If Not frmNewComputer.Cancelled Then
'     Set li = lbComputerNames.ListItems.Add
'     li.Text = frmNewComputer.txtFriendlyName.Text
'     li.Tag = frmNewComputer.txtComputerName
'     AreChanges = True
'   End If
'  Unload frmNewComputer
'  Me.Show
Dim li As ListItem
 If txtComputerName.Text <> "" Then
  If Not ComputerExists(txtComputerName.Text) Then
    Set li = lbComputerNames.ListItems.Add
    li.Text = txtFriendlyName.Text
    li.Tag = Trim(LCase(txtComputerName.Text))
    AreChanges = True
    txtComputerName.Text = ""
    txtFriendlyName.Text = ""
  Else
    MsgBox "The host name already exists."
  End If
 End If
End Sub

Private Sub cmdCancel_Click()
Dim i As Integer
  'get the computers
  If AreChanges Then
    ComputerCount = lbComputerNames.ListItems.Count
    If ComputerCount <> 0 Then
      ReDim arrComputers(ComputerCount) As CompRec
      
      For i = 0 To (ComputerCount - 1)
        arrComputers(i + 1).netName = lbComputerNames.ListItems(i + 1).Tag
        arrComputers(i + 1).friendName = lbComputerNames.ListItems(i + 1).Text
      Next i
    End If
    SaveSettings
    AreChanges = False
  End If
  Me.Hide
End Sub

Private Sub cmdClear_Click()
  txtMessage.Text = ""
  txtMessage.SetFocus
End Sub
Private Sub DoSend()
Dim li As ListItem
  On Error GoTo err_handler
   If boolSendToSelf Then
       WSServer.RemoteHost = strMyName 'name of remote computer or IP Address
      'send the data
       WSServer.SendData txtMessage.Text
   End If
   For Each li In lbComputerNames.ListItems
     If li.Checked Then
      'get the nameof the computer to send the data to
       WSServer.RemoteHost = li.Tag 'name of remote computer or IP Address
      'send the data
       WSServer.SendData txtMessage.Text
     End If
   Next li
   Exit Sub
err_handler:
  Select Case Err.Number
    Case 10014:
      'MsgBox "User is not listening"
      Resume Next
    Case Else
      MsgBox Err.Number
  End Select

End Sub

Private Sub cmdDelete_Click()
Dim CurrItem As ListItem
  Set CurrItem = lbComputerNames.SelectedItem
  If Not CurrItem Is Nothing Then
    lbComputerNames.ListItems.Remove (CurrItem.Index)
    AreChanges = True
  End If
End Sub

Private Sub cmdEdit_Click()
Dim CurrItem As ListItem
  Set CurrItem = lbComputerNames.SelectedItem
  If Not CurrItem Is Nothing Then
    txtFriendlyName.Text = CurrItem.Text
    txtComputerName.Text = CurrItem.Tag

    If Not frmNewComputer.Cancelled Then
      CurrItem.Text = frmNewComputer.txtFriendlyName
      CurrItem.Tag = frmNewComputer.txtComputerName
      AreChanges = True
    End If
    Unload frmNewComputer
    
  End If
End Sub

Private Sub CloseForm()
Dim i As Integer
  'get the computers
  If AreChanges Then
    ComputerCount = lbComputerNames.ListItems.Count
    If ComputerCount <> 0 Then
      ReDim arrComputers(ComputerCount) As CompRec
      
      For i = 0 To (ComputerCount - 1)
        arrComputers(i + 1).netName = lbComputerNames.ListItems(i + 1).Tag
        arrComputers(i + 1).friendName = lbComputerNames.ListItems(i + 1).Text
      Next i
    End If
    SaveSettings
    AreChanges = False
  End If
  Me.Hide
End Sub
Private Sub cmdSend_Click()
  
  DoSend
  
  If chkCloseForm.Value <> 0 Then
    CloseForm
  Else
    txtMessage.Text = ""
    txtMessage.SetFocus
  End If
  
End Sub

Private Sub Form_Activate()
  
  'txtMessage.Text = ""
  'txtMessage.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Integer
  Dim li As ListItem
  AreChanges = False
  boolSendToSelf = True
  chkCloseForm.Value = 1
  WSServer.RemotePort = 1007
  strMyName = WSServer.LocalHostName
  LoadSettings
  
   'add companies to list
   If ComputerCount <> 0 Then
     For i = LBound(arrComputers) To UBound(arrComputers)
       Set li = lbComputerNames.ListItems.Add
         li.Text = arrComputers(i).friendName
         li.Tag = arrComputers(i).netName
     Next i
   End If
   
   cmdDelete.Enabled = Not (lbComputerNames.SelectedItem Is Nothing)
   cmdAdd.Enabled = False


  
End Sub
Private Sub Form_Unload(Cancel As Integer)
 'Cancel = True
 'Me.Hide
End Sub
Private Sub lbComputerNames_AfterLabelEdit(Cancel As Integer, NewString As String)
  If NewString <> lbComputerNames.SelectedItem.Text Then
    AreChanges = True
  End If
End Sub

Private Sub lbComputerNames_Click()
    cmdDelete.Enabled = Not (lbComputerNames.SelectedItem Is Nothing)
End Sub

Private Sub txtComputerName_Change()
 cmdAdd.Enabled = (txtComputerName.Text <> "")
End Sub
