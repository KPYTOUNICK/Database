VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatabaseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private pConnection As ADODB.Connection

Private pConnectionString As String

Private pIsConnected As Boolean

Public Sub Class_Initialize()
    Set pConnection = New ADODB.Connection
    pIsConnected = False
End Sub

Private Property Let ConnectionExists(Status As Boolean)
    Let pIsConnected = Status
    Let frmConnection.btnSingin.Enabled = Not Status
    Let frmConnection.btnSingout.Enabled = Status
End Property

Public Property Get Exists() As Boolean
    Exists = pIsConnected
End Property

Public Property Get Instance() As ADODB.Connection
    Set Instance = pConnection
End Property

Public Sub Create(Connection As String)
    
    Let pConnectionString = Connection
    
    On Error GoTo ErrorHandler_Connection
    pConnection.Open pConnectionString
    
    Let ConnectionExists = True
    
    Exit Sub
    
ErrorHandler_Connection:
    MsgBox "Couldn't connect to the server. Try again."
    
End Sub

Public Sub Destroy()
    
    If pIsConnected Then
        On Error GoTo ErrorHandler_Disconnection
        pConnection.Close
        Let ConnectionExists = False
    Else
        MsgBox "Disconnecting from the server is not possible because there is no connection."
    End If
    
    Exit Sub
    
ErrorHandler_Disconnection:
    MsgBox "Couldn't disconnect from the server correctly."
    
End Sub
