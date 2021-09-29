VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IClass

Private mCn As Object
Private mRs As Object
Private mReady As Boolean

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    If Not (mRs Is Nothing) Then
        If (mRs.State And 1) = 1 Then
            mRs.Close
        End If
    End If

    If Not (mCn Is Nothing) Then
        If (mCn.State And 1) = 1 Then
            mCn.Close
        End If
    End If
End Sub

Public Sub ConnectServerAuth(data_source As String, initial_catalog As String, user_id As String, user_password As String)
    Call ReadyCheck
    Call Class_Terminate

    Dim s As String2
    
    Set s = Init(New String2, "Provider=SQLOLEDB;Data Source={0};Initial Catalog={1};UID={2};PWD={3};")

    mCn.ConnectionString = s.Format(data_source, initial_catalog, user_id, user_password)
    mCn.Open
End Sub

Public Sub ConnectWindowsAuth(data_source As String, initial_catalog As String)
    Call ReadyCheck
    Call Class_Terminate
    
    Dim s As String2
    
    Set s = Init(New String2, "Provider=SQLOLEDB;Data Source={0};Initial Catalog={1};Trusted_Connection=Yes;")
    
    mCn.ConnectionString = s.Format(data_source, initial_catalog)
    mCn.Open
End Sub

Public Sub Execute(sql As String)
    Call ReadyCheck
    
    If mCn Is Nothing Then
        Call ErrRaise(515)
    End If

    If (mCn.State And 1) <> 1 Then
        Call ErrRaise(515)
    End If

    mCn.Execute sql
End Sub

Private Function IClass_Constructor(args() As Variant) As Object
    If mReady Then
        Call ErrRaise(513)
    End If
        
    Set mCn = CreateObject("ADODB.Connection")
    Set mRs = CreateObject("ADODB.Recordset")
    
    mReady = True
    
    Set IClass_Constructor = Me
End Function

Public Function Query(sql As String) As Object
    Call ReadyCheck
    
    If mCn Is Nothing Then
        Call ErrRaise(515)
    End If
    
    If (mCn.State And 1) <> 1 Then
        Call ErrRaise(515)
    End If

    Set mRs = mCn.Execute(sql)
    Set Query = mRs
End Function

' ----- >

Private Sub ErrRaise(e As Long)
    Select Case e
        Case 513:
            Err.Raise e, TypeName(Me), "It has already been initialized."
        Case 514:
            Err.Raise e, TypeName(Me), "Please initialize."
        Case 515:
            Err.Raise e, TypeName(Me), "Please connect to the database."
        Case Else:
            Err.Raise e, TypeName(Me)
    End Select
End Sub

Private Sub ReadyCheck()
    If Not mReady Then
        Call ErrRaise(514)
    End If
End Sub


