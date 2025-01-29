VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xloDBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not DebugMode Then
        MsgBox "You can't save this workbook!"
        Cancel = True
    End If
End Sub

Private Sub Workbook_NewSheet(ByVal Sheet As Object)
    If Not DebugMode Then
        With Application
            
            .DisplayAlerts = False
            Sheet.Delete
            .DisplayAlerts = True
            
            MsgBox "You can't create new worksheet!"
            
        End With
    End If
End Sub

Private Sub Workbook_Open()
    
    DebugMode = False
    
    Call mMain.DatabaseManager_Initialize
    Call mMain.OpenConnectionWindow
    
End Sub
