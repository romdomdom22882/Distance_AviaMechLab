Attribute VB_Name = "modGlobals"
Option Compare Database
Option Explicit

' === Настройки проекта ===
Public Const AUTH_TABLE_USERS As String = "tUsers"
Public Const AUTH_TABLE_ROLES As String = "tRoles"
Public Const AUTH_TABLE_LOG As String = "tAuthLog"

Public Const AUTH_LOCK_AFTER_FAILED As Long = 5
Public Const AUTH_LOCK_MINUTES As Long = 10

' === Утилиты ===
Public Function SqlQ(ByVal s As String) As String
    ' Кавычки для SQL: ' -> ''
    SqlQ = "'" & Replace(Nz(s, vbNullString), "'", "''") & "'"
End Function

Public Function ClientInfo() As String
    On Error Resume Next
    ClientInfo = "COMPUTER=" & Environ$("COMPUTERNAME") & "; USER=" & Environ$("USERNAME")
End Function

Public Sub TempVar_Set(ByVal name As String, ByVal value As Variant)
    On Error Resume Next
    TempVars.Remove name
    Err.Clear
    TempVars.Add name, value
End Sub

Public Function TempVar_Get(ByVal name As String, Optional ByVal defaultValue As Variant) As Variant
    On Error GoTo EH
    TempVar_Get = TempVars(name).Value
    Exit Function
EH:
    TempVar_Get = defaultValue
End Function

Public Sub TempVar_Remove(ByVal name As String)
    On Error Resume Next
    TempVars.Remove name
End Sub

