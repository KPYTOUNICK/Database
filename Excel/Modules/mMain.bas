Attribute VB_Name = "mMain"
Global Database As cDatabaseHandler

Global WorksheetSettings As Worksheet
Global WorksheetWorkspace As Worksheet

Public Sub DatabaseManager_Initialize()
    
End Sub

Public Sub OpenConnectionWindow()
    frmConnection.Show
End Sub

Public Sub OpenCreateTableWindow()
    frmCreateTable.Show
End Sub
