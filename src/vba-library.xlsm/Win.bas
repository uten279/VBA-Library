Attribute VB_Name = "Win"
Option Explicit

Public Function InstalledApps() As Collection
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const SUB_KEY_NAME = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"

    Dim apps1 As Collection
    Dim obj1 As Object
    Dim keys1 As Variant
    Dim key1 As Variant
    Dim name1 As String
    Dim r As Long
    
    Set apps1 = New Collection
    Set obj1 = CreateObject("WbemScripting.SWbemLocator").ConnectServer(, "root\default").Get("StdRegProv")
    
    obj1.EnumKey HKEY_LOCAL_MACHINE, SUB_KEY_NAME, keys1

    On Error Resume Next

    For Each key1 In keys1
        name1 = ""
        r = obj1.GetStringValue(HKEY_LOCAL_MACHINE, SUB_KEY_NAME & key1, "DisplayName", name1)
        
        If r <> 0 Then
            r = obj1.GetStringValue(HKEY_LOCAL_MACHINE, SUB_KEY_NAME & key1, "QuietDisplayName", name1)
        End If
        
        If r = 0 And Len(Trim(name1)) > 0 Then
            apps1.Add name1
        End If
    Next
    
    On Error GoTo 0
    
    Set InstalledApps = apps1
End Function

Public Function HasInstalled(name_1 As String) As Boolean
    Dim names1 As Collection
    Dim i As Long
    
    Set names1 = Win.InstalledApps()
    
    For i = 1 To names1.Count
        If names1(i) = name_1 Then
            HasInstalled = True
            Exit Function
        End If
    Next i
End Function
