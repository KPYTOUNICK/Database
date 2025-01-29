VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateTable 
   Caption         =   "Create Table"
   ClientHeight    =   8385.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9570.001
   OleObjectBlob   =   "frmCreateTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pDataTypesRange As Range
Private pNamesException As Variant
Private pFieldsCount As Long
Private pIsReseting As Boolean

Private Sub Reset(Flag As Boolean)
    
    Dim ctrl As Control
    
    btnAddField1.Visible = True
    
    For Each ctrl In Me.Controls
        
        If TypeName(ctrl) = "ComboBox" Then
            For Each cell In pDataTypesRange
                ctrl.AddItem cell.Value
            Next cell
        End If
        
        If Not mAdditional.IsInArray(ctrl.Name, pNamesException) Then
            ctrl.Visible = False
        End If
        
    Next ctrl
    
End Sub

Private Sub AddField(Number As Long)
    
    Inc pFieldsCount
    
    With Me
        
        With Controls("btnAddField" & Number + 1)
            If Not Number = 6 Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        
        .Controls("btnDeleteField" & Number + 1).Visible = True
        .Controls("cbFieldType" & Number + 1).Visible = True
        .Controls("lblFieldType" & Number + 1).Visible = True
        .Controls("lblFieldName" & Number + 1).Visible = True
        .Controls("txtFieldName" & Number + 1).Visible = True
        
        If Not Number = 1 Then
            .Controls("btnDeleteField" & Number).Visible = False
        End If
        
        .Controls("btnAddField" & Number).Visible = False
        
    End With
    
End Sub

Private Sub DeleteField(Number As Long)
    
    Let pFieldsCount = pFieldsCount - 1
    
    With Me
        
        With Controls("btnDeleteField" & Number - 1)
            If Not Number = 2 Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        
        Let pIsReseting = False
        
        .Controls("cbFieldType" & Number).Visible = False
        .Controls("lblFieldType" & Number).Visible = False
        .Controls("lblFieldName" & Number).Visible = False
        .Controls("txtFieldName" & Number).Visible = False
        .Controls("btnAddField" & Number).Visible = False
        .Controls("btnDeleteField" & Number).Visible = False
        
        .Controls("btnAddField" & Number - 1).Visible = True
        
    End With
    
    Let pIsReseting = True
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("SETTINGS")
    Dim ctrl As Control
    Dim rngDataTypes, cell As Range
    
    Let pFieldsCount = 1
    Let pIsReseting = True
    
    Let pNamesException = Array( _
                         "frameTable", "frameMain", _
                         "cbFieldType1", _
                         "txtFieldName1", "txtLogin", "txtTable", _
                         "btnAddField1", "btnReset", "btnCreate", _
                         "lblSchema", "lblTable", "lblFieldName1", "lblFieldType1" _
                         )
    
    With ws
        Set pDataTypesRange = .Range(.Cells(1, 1), .Cells(1, 1).CurrentRegion)
    End With
    
    Call Reset(True)
     
End Sub

Private Sub btnReset_Enter()
    If pIsReseting Then
        Call Reset(pIsReseting)
    End If
End Sub

Private Sub btnAddField1_Click()
    Call AddField(1)
End Sub

Private Sub btnAddField2_Click()
    Call AddField(2)
End Sub

Private Sub btnAddField3_Click()
    Call AddField(3)
End Sub

Private Sub btnAddField4_Click()
    Call AddField(4)
End Sub

Private Sub btnAddField5_Click()
    Call AddField(5)
End Sub

Private Sub btnAddField6_Click()
    Call AddField(6)
End Sub

Private Sub btnDeleteField2_Click()
    Call DeleteField(2)
End Sub

Private Sub btnDeleteField3_Click()
    Call DeleteField(3)
End Sub

Private Sub btnDeleteField4_Click()
    Call DeleteField(4)
End Sub

Private Sub btnDeleteField5_Click()
    Call DeleteField(5)
End Sub

Private Sub btnDeleteField6_Click()
    Call DeleteField(6)
End Sub

Private Sub btnDeleteField7_Click()
    Call DeleteField(7)
End Sub
