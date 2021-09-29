VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "String2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IClass

Private mString As String
Private mReady As Boolean

Public Function Format(ParamArray args()) As String2
    Call ReadyCheck

    Dim s As String2
    Dim i As Long

    Set s = Init(New String2, mString)

    For i = 0 To UBound(args)
        Set s = Init(New String2, s.Replace("\{" & i & "(:.+?|)\}", CStr(args(i))))
    Next i

    Set Format = s
End Function

Public Function Replace(p As String, r As String, Optional i As Boolean = False, Optional g As Boolean = True) As String2
    Call ReadyCheck

    Dim o As Object

    Set o = CreateObject("VBScript.RegExp")

    o.Pattern = p
    o.IgnoreCase = i
    o.Global = g

    Set Replace = Init(New String2, o.Replace(mString, r))
End Function

Public Function Test(p As String, Optional i As Boolean = False) As Boolean
    Call ReadyCheck

    Dim o As Object

    Set o = CreateObject("VBScript.RegExp")

    o.Pattern = p
    o.IgnoreCase = i

    Test = o.Test(mString)
End Function

Property Let Value(x As String)
Attribute Value.VB_UserMemId = 0
    Call ReadyCheck

    mString = x
End Property

Property Get Value() As String
    Call ReadyCheck
    
    Value = mString
End Property

Private Function IClass_Constructor(args() As Variant) As Object
    If mReady Then
        Call ErrRaise(513)
    End If
    
    If UBound(args) >= 0 Then
        If Not IsObject(args(0)) Or TypeName(args(0)) = "String2" Then
            mString = CStr(args(0))
        End If
    End If
    
    mReady = True
    
    Set IClass_Constructor = Me
End Function

' ----- >

Private Sub ErrRaise(e As Long)
    Select Case e
        Case 513:
            Err.Raise e, TypeName(Me), "It has already been initialized."
        Case 514:
            Err.Raise e, TypeName(Me), "Please initialize."
        Case Else:
            Err.Raise e, TypeName(Me)
    End Select
End Sub

Private Sub ReadyCheck()
    If Not mReady Then
        Call ErrRaise(514)
    End If
End Sub

