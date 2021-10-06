Attribute VB_Name = "Func"
Option Explicit

Public Function Init(o As IClass, ParamArray x()) As Object
    Dim v()
    
    v = x

    Set Init = o.Constructor(v)
End Function

Public Function InValue(value_1 As Variant, ParamArray args() As Variant) As Boolean
    If IsObject(value_1) Then
        Exit Function
    End If
    
    Dim i As Long
    
    For i = 0 To UBound(args, 1)
        If IsObject(args(i)) Then
            ' Continue
        ElseIf IsArray(args(i)) Then
            ' Continue
        ElseIf value_1 = args(i) Then
            InValue = True
            Exit For
        End If
    Next i
End Function
