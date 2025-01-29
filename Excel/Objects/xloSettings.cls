VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xloSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private ws As Worksheet

Public Sub Initialize()
    Set ws = ThisWorkbook.Worksheets("SETTINGS")
End Sub

Public Function RowsCount(ByVal ColumnIndex As Long) As Long
    RowsCount = ws.Range(ColumnIndex).End(xlDown).Row
End Function
