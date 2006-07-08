VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function GetName() As String
    GetName = "Pts Counter Plugin  Version " & App.Major & "." & App.Minor
End Function
Public Function GetCommands() As String
    GetCommands = ",Wipe,3,Inc,3,Points,1,ZoneAdd,3,Zonelist,1,"
End Function

Public Function main(dadata As String) As String
main = "testing.. data-" & dadata
End Function

