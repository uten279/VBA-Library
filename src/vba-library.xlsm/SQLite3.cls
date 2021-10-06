VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IClass

Private mCn As Object
Private mRs As Object
Private mReady As Boolean

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

Public Sub Execute(sql)
    Call ReadyCheck

    mCn.Execute sql
End Sub

Private Function IClass_Constructor(x() As Variant) As Object
    If mReady Then
        Call ErrRaise(513)
    End If
    
    If Not CanUsed() Then
        Call ErrRaise(516)
    End If
    
    Set mCn = CreateObject("ADODB.Connection")
    Set mRs = CreateObject("ADODB.Recordset")
    
    mCn.Open ConnectionString:="DRIVER=SQLite3 ODBC Driver; DataBase=" & x(0)
    
    mReady = True
    
    Set IClass_Constructor = Me
End Function

Public Function Query(sql As String) As Object
    Call ReadyCheck

    If (mRs.State And 1) = 1 Then
        mRs.Close
    End If

    mRs.CursorLocation = 3  ' adUseClient
    mRs.Open sql, mCn
    
    Set Query = mRs
End Function

Public Function TableExists(name_1 As String) As Boolean
    Call ReadyCheck
    
    Dim rs As Object
    Dim sql As String
    
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "     name"
    sql = sql & " FROM"
    sql = sql & "     sqlite_master"
    sql = sql & " WHERE"
    sql = sql & "     type = 'table'"
    sql = sql & " AND name = '" & name_1 & "'"
    
    mCn.CursorLocation = 3  ' adUseClient
    Set rs = mCn.Execute(sql)
    
    TableExists = rs.RecordCount > 0
    
    rs.Close
    Set rs = Nothing
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
        Case 516:
            Err.Raise e, TypeName(Me), "Please install SQLite ODBC driver(32&64). -> http://www.ch-werner.de/sqliteodbc/"
        Case Else:
            Err.Raise e, TypeName(Me)
    End Select
End Sub

Private Function CanUsed() As Boolean
    Const SQLITE3_ODBC_DISPLAY_NAME_32 As String = "SQLite ODBC Driver (remove only)"
    Const SQLITE3_ODBC_DISPLAY_NAME_64 As String = "SQLite ODBC Driver for Win64 (remove only)"
    
    CanUsed = Win.HasInstalled(SQLITE3_ODBC_DISPLAY_NAME_32) And Win.HasInstalled(SQLITE3_ODBC_DISPLAY_NAME_64)
End Function

Private Sub ReadyCheck()
    If Not mReady Then
        Call ErrRaise(514)
    End If
End Sub

