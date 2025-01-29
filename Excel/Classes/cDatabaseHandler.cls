VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatabaseHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Connection As cDatabaseConnection
Public Recordset As cDatabaseRecordset

Public Sub Class_Initialize()
    
    Set Connection = New cDatabaseConnection
    Set Recordset = New cDatabaseRecordset
    
End Sub
