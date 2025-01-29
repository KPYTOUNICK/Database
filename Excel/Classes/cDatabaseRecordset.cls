VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatabaseRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pRecordset As ADODB.Recordset

Public Sub Class_Initialize()
    Set pRecordset = New ADODB.Recordset
End Sub

Public Sub CreateWithOpen(Query As String, Connection As String)
    pRecordset.Open Query, Connection
End Sub

Public Sub CreateWithExecute(Query As String, Connection As String)
    Set pRecordset = Connection.Execute(Query)
End Sub
