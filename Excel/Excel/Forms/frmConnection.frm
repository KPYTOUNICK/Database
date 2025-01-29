VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConnection 
   Caption         =   "Sing in"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7125
   OleObjectBlob   =   "frmConnection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    If Database Is Nothing Then
        Set Database = New cDatabaseHandler
    End If
    
    With Me
        If Not Database.Connection.Exists Then
            btnSingin.Enabled = True
            btnSingout.Enabled = False
        Else
            btnSingin.Enabled = False
            btnSingout.Enabled = True
        End If
    End With
    
    frmConnection.txtLogin.SetFocus
    
End Sub

Private Sub btnSingin_Enter()
    
    Dim Caption As String
    
    With frmConnection
        Caption = "DSN=" & txtDriver & ";" & _
                  "UID=" & txtLogin & ";" & _
                  "PWD=" & txtPassword & ";" & _
                  "Database=" & txtDatabase
    End With
    
    Database.Connection.Create Caption
    
    If Database.Connection.Exists Then
        frmConnection.Hide
        ActiveWorkbook.Windows(1).Visible = True
    End If
    
    frmConnection.txtLogin.SetFocus
    
End Sub

Private Sub btnSingout_Enter()
    
    Database.Connection.Destroy
    
    frmConnection.txtLogin.SetFocus
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Me.Hide
End Sub
