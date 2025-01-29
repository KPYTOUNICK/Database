Attribute VB_Name = "mDebug"
Global DebugMode As Boolean

Public Sub ShowApplicationWindow()
    ThisWorkbook.Windows(1).Visible = True
End Sub

' Запускать функцию всегда, при редактировании файла с макросом
Public Sub SwitchMode()
    Let DebugMode = Not DebugMode
    Beep
End Sub
