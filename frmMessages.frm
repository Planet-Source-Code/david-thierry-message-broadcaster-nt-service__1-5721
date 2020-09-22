VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMessages 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Message Board"
   ClientHeight    =   1020
   ClientLeft      =   4440
   ClientTop       =   7965
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock WSClient 
      Left            =   7440
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   120
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   0
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim Tempstring As String
Dim lngBitBltRslt As Long
Dim strText As String
Dim isTopMost As Boolean
Dim boolDisplayingMsg As Boolean

Dim colMessages As Collection
Dim iNextMessage As Integer

Function GetMessage() As Boolean
  
  If colMessages.Count = 0 Then
    GetMessage = False
    Exit Function
  End If
  
  On Error GoTo err_handler
  strText = colMessages.Item(iNextMessage)
  If Not boolLoopMessages Then
     colMessages.Remove 1
  Else
    If iNextMessage = colMessages.Count Then
      iNextMessage = 1
    Else
      iNextMessage = iNextMessage + 1
    End If
  End If

  P1.Width = P1.TextWidth(strText) + 25
  PrintText strText
  theleft = ScaleWidth
  p1wid = P1.ScaleWidth  'p1wid = P1.Width
  GetMessage = True
  Exit Function
err_handler:
  iNextMessage = 1
  GetMessage = False
End Function
Sub Form_Load()

        Set colMessages = New Collection
        
        WSClient.Bind 1007
        isTopMost = True
        iNextMessage = 1
        boolDisplayingMsg = False
        'CenterInWorkArea Me
        AlignBottom Me
        'AlignTop Me
        P1.AutoRedraw = True
        P1.Visible = False
        P1.FontSize = 12

        P1.ForeColor = &HFF0000
        P1.BackColor = BackColor
        P1.ScaleMode = 3
        ScaleMode = 3
        P1.Height = P1.TextHeight("Test Height") + 4
        P1.Top = (Me.ScaleHeight / 2) - (P1.ScaleHeight / 2)
        P1.Left = 0
        thetop = P1.Top
        p1hgt = P1.ScaleHeight
        SetTopMost Me, isTopMost
        
        Timer1.Enabled = True
        Timer1.Interval = 10
        

End Sub
Sub Form_Load_Save()
 
        Set colMessages = New Collection
        
        WSClient.Bind 1007
        isTopMost = True
       
        'CenterInWorkArea Me
        AlignBottom Me
        'AlignTop Me
        P1.AutoRedraw = True
        P1.Visible = False
        P1.FontSize = 12

        P1.ForeColor = &HFF0000
        P1.BackColor = BackColor
        P1.ScaleMode = 3
        ScaleMode = 3
        Open (App.Path & "\credits.txt") For Input As #1
        Line Input #1, Tempstring
        'P1.Height = (Val(Tempstring) * P1.TextHeight("Test Height")) + 200
        'P1.Height = (1 * P1.TextHeight("Test Height")) + 200
        P1.Height = P1.TextHeight("Test Height") + 4

        Do Until EOF(1)
            Line Input #1, Tempstring
            strText = strText & " " & Tempstring
            'PrintText Tempstring
        Loop
        Close #1
        P1.Width = P1.TextWidth(strText) + 25
        P1.Top = (Me.ScaleHeight / 2) - (P1.ScaleHeight / 2)
        P1.Left = 0
        PrintText strText
        theleft = ScaleWidth
        thetop = P1.Top
        p1hgt = P1.ScaleHeight
        p1wid = P1.Width
        SetTopMost Me, isTopMost
        Timer1.Enabled = True
        Timer1.Interval = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
  While colMessages.Count <> 0
    colMessages.Remove 1
  Wend
  Set colMessages = Nothing
End Sub
Sub ProcessMessages()
  If boolDisplayingMsg Then
     lngBitBltRslt = BitBlt(hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
     theleft = theleft - 1
     If theleft < -p1wid Then 'we've reach the end of the current message
       boolDisplayingMsg = False
     End If
  Else
     If GetMessage Then
       boolDisplayingMsg = True
       If Not Me.Visible Then Me.Show
     Else
       Me.Hide
     End If
  End If
End Sub
Sub Timer1_Timer()
'       lngBitBltRslt = BitBlt(hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
'        theleft = theleft - 1
'        If theleft < -p1wid Then
'        Timer1.Enabled = False
'        'Txt$ = "Credits Completed"
'        CurrentY = ScaleHeight / 2
'        CurrentX = (ScaleWidth - TextWidth("Credits Completed")) / 2
'        Print "Credits Completed"
'        End If
         ProcessMessages
End Sub

Sub PrintText(Text As String)
  P1.Cls
  'P1.CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
  P1.CurrentX = 0
  P1.ForeColor = &HFFFF&
  P1.Print Text
  'If frmBroadcast.Visible Then
  '  frmBroadcast.SetFocus
  'End If
End Sub

Private Sub WSClient_DataArrival(ByVal bytesTotal As Long)
Dim inData As String
  PlayWav ("Reminder.wav")
  WSClient.GetData inData, vbString
  colMessages.Add inData

End Sub

