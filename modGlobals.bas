Attribute VB_Name = "modGlobals"


Public Type CompRec
  netName As String
  friendName As String
End Type

Public ComputerCount
Public arrComputers() As CompRec

Public boolSendToSelf As Boolean
Public boolLoopMessages As Boolean


