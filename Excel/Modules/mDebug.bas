Attribute VB_Name = "mDebug"

Global DebugMode As Boolean

Public Sub ShowApplicationWindow()
    ThisWorkbook.Windows(1).Visible = True
End Sub

Public Sub SwitchMode()
    Let DebugMode = Not DebugMode
    Beep
End Sub
