VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public pConn As ADODB.Connection
Public openConnection As Boolean


Private Sub Class_Initialize()
On Error GoTo err

    Set pConn = New ADODB.Connection

    pConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\file.mdb;Jet OLEDB:database password=Scope4"
    openConnection = True
    Exit Sub
err:
    openConnection = False
End Sub

Private Sub Class_Terminate()
On Error GoTo err
    Set pConn = Nothing
err:
Debug.Print "Class_Terminate: " & err.Number & " " & err.Description: Resume Next
End Sub




