Attribute VB_Name = "mAdditional"

Function Inc(ByRef Value As Variant)
    Let Value = Value + 1
End Function

Function IsInArray(Value As Variant, Arr As Variant) As Boolean
    
    Dim Element As Variant
    
    IsInArray = False
    
    For Each Element In Arr
        If Element = Value Then
            IsInArray = True
            Exit Function
        End If
    Next Element
    
End Function

Function IsIn2DArray(Value As Variant, Arr As Variant) As Boolean
    
    Dim i, j As Long
    
    IsIn2DArray = False
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            If Arr(i, j) = Value Then
                IsIn2DArray = True
                Exit Function
            End If
        Next j
    Next i
    
End Function
