VERSION 5.00
Begin VB.Form frmNewComputer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Computer"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtFriendlyName 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox txtComputerName 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Friendly Name"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmNewComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private boolCancelled As Boolean

Private Sub cmdCancel_Click()
  boolCancelled = True
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  boolCancelled = False
  Me.Hide
End Sub

Public Property Get Cancelled() As Variant
  Cancelled = boolCancelled
End Property



Private Sub Form_Load()
  boolCancelled = True
End Sub
