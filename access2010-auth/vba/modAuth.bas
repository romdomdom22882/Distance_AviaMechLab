Attribute VB_Name = "modAuth"
Option Compare Database
Option Explicit

' === Публичный API ===

Public Function Auth_IsLoggedIn() As Boolean
    Dim v As Variant
    v = TempVar_Get("Auth_UserID", Null)
    Auth_IsLoggedIn = Not IsNull(v)
End Function

Public Sub Auth_Logout()
    TempVar_Remove "Auth_UserID"
    TempVar_Remove "Auth_Login"
    TempVar_Remove "Auth_RoleID"
    TempVar_Remove "Auth_IsAdmin"
End Sub

Public Function Auth_Login(ByVal login As String, ByVal password As String, Optional ByVal errorLabel As Object) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Dim userId As Long
    Dim roleId As Long
    Dim isAdmin As Boolean
    Dim isActive As Boolean
    Dim mustChange As Boolean
    Dim saltHex As String
    Dim hashHex As String
    Dim storedHash As String
    Dim failed As Long
    Dim lockedUntil As Variant

    Auth_Login = False
    SetErrorText errorLabel, vbNullString

    login = Trim$(Nz(login, vbNullString))
    password = Nz(password, vbNullString)
    If Len(login) = 0 Or Len(password) = 0 Then
        SetErrorText errorLabel, "Введите логин и пароль."
        Exit Function
    End If

    Set db = CurrentDb
    sql = "SELECT * FROM " & AUTH_TABLE_USERS & " WHERE [Login]=" & SqlQ(login)
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)

    If rs.EOF Then
        Call Auth_LogAttempt(login, Null, False, "Пользователь не найден")
        SetErrorText errorLabel, "Неверный логин или пароль."
        Exit Function
    End If

    userId = rs!UserID
    roleId = rs!RoleID
    isActive = rs!IsActive
    mustChange = rs!MustChangePassword
    failed = Nz(rs!FailedAttempts, 0)
    lockedUntil = rs!LockedUntil

    If Not isActive Then
        Call Auth_LogAttempt(login, userId, False, "Пользователь отключен (IsActive=False)")
        SetErrorText errorLabel, "Учетная запись отключена."
        Exit Function
    End If

    If Not IsNull(lockedUntil) Then
        If lockedUntil > Now() Then
            Call Auth_LogAttempt(login, userId, False, "Учетная запись заблокирована до " & CStr(lockedUntil))
            SetErrorText errorLabel, "Учетная запись заблокирована. Повторите позже."
            Exit Function
        End If
    End If

    saltHex = Nz(rs!PasswordSalt, vbNullString)
    storedHash = Nz(rs!PasswordHash, vbNullString)

    hashHex = Crypto_Sha256Hex(saltHex & password)
    If Len(hashHex) = 0 Then
        Call Auth_LogAttempt(login, userId, False, "Ошибка вычисления хэша (CryptoAPI)")
        SetErrorText errorLabel, "Ошибка безопасности: не удалось проверить пароль."
        Exit Function
    End If

    If StrComp(LCase$(storedHash), LCase$(hashHex), vbTextCompare) <> 0 Then
        failed = failed + 1
        rs.Edit
        rs!FailedAttempts = failed
        If failed >= AUTH_LOCK_AFTER_FAILED Then
            rs!LockedUntil = DateAdd("n", AUTH_LOCK_MINUTES, Now())
        End If
        rs.Update

        Call Auth_LogAttempt(login, userId, False, "Неверный пароль; FailedAttempts=" & CStr(failed))
        SetErrorText errorLabel, "Неверный логин или пароль."
        Exit Function
    End If

    ' Успешный вход
    isAdmin = Auth_IsRoleAdmin(roleId)

    rs.Edit
    rs!FailedAttempts = 0
    rs!LockedUntil = Null
    rs!LastLoginAt = Now()
    rs.Update

    TempVar_Set "Auth_UserID", userId
    TempVar_Set "Auth_Login", login
    TempVar_Set "Auth_RoleID", roleId
    TempVar_Set "Auth_IsAdmin", isAdmin

    Call Auth_LogAttempt(login, userId, True, "OK")

    If mustChange Then
        ' В учебном проекте просто сообщаем — как вариант, открывайте форму смены пароля.
        SetErrorText errorLabel, "Пароль нужно сменить (MustChangePassword=True)."
    End If

    Auth_Login = True
End Function

Public Sub Auth_EnsureSeedData()
    ' Создает роль Admin и пользователя admin/Admin123! (если их нет)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim adminRoleId As Long
    Dim userId As Long

    Set db = CurrentDb

    adminRoleId = EnsureRole("Admin", True)
    Call EnsureRole("User", False)

    Set rs = db.OpenRecordset("SELECT UserID FROM " & AUTH_TABLE_USERS & " WHERE [Login]='admin'", dbOpenSnapshot)
    If rs.EOF Then
        userId = CreateUser("admin", "Администратор", adminRoleId, True, True)
        Call Auth_SetPassword(userId, "Admin123!", True)
    End If
End Sub

Public Sub Auth_SetPassword(ByVal userId As Long, ByVal newPassword As String, Optional ByVal mustChangePassword As Boolean = False)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim saltHex As String
    Dim hashHex As String

    If userId <= 0 Then Exit Sub
    If Len(newPassword) = 0 Then Exit Sub

    saltHex = Crypto_RandomSaltHex(16)
    hashHex = Crypto_Sha256Hex(saltHex & newPassword)
    If Len(saltHex) = 0 Or Len(hashHex) = 0 Then Exit Sub

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & AUTH_TABLE_USERS & " WHERE UserID=" & CStr(userId), dbOpenDynaset, dbSeeChanges)
    If rs.EOF Then Exit Sub

    rs.Edit
    rs!PasswordSalt = saltHex
    rs!PasswordHash = hashHex
    rs!MustChangePassword = mustChangePassword
    rs.Update
End Sub

Public Function Auth_CurrentUserId() As Variant
    Auth_CurrentUserId = TempVar_Get("Auth_UserID", Null)
End Function

Public Function Auth_CurrentLogin() As String
    Auth_CurrentLogin = Nz(TempVar_Get("Auth_Login", vbNullString), vbNullString)
End Function

Public Function Auth_IsAdmin() As Boolean
    Auth_IsAdmin = CBool(Nz(TempVar_Get("Auth_IsAdmin", False), False))
End Function

' === Внутреннее ===

Private Function Auth_IsRoleAdmin(ByVal roleId As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Auth_IsRoleAdmin = False
    If roleId <= 0 Then Exit Function

    Set db = CurrentDb
    sql = "SELECT IsAdmin FROM " & AUTH_TABLE_ROLES & " WHERE RoleID=" & CStr(roleId)
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then Exit Function
    Auth_IsRoleAdmin = CBool(Nz(rs!IsAdmin, False))
End Function

Private Sub Auth_LogAttempt(ByVal login As String, ByVal userId As Variant, ByVal success As Boolean, ByVal message As String)
    On Error GoTo EH
    Dim db As DAO.Database
    Dim sql As String

    Set db = CurrentDb
    sql = "INSERT INTO " & AUTH_TABLE_LOG & " ([Login], UserID, Success, ClientInfo, EventAt, [Message]) VALUES (" & _
          SqlQ(login) & ", " & IIf(IsNull(userId), "NULL", CStr(userId)) & ", " & IIf(success, "True", "False") & ", " & _
          SqlQ(ClientInfo()) & ", #" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "#, " & SqlQ(Left$(message, 255)) & ")"
    db.Execute sql, dbFailOnError
    Exit Sub
EH:
    ' аудит не должен ломать авторизацию
End Sub

Private Sub SetErrorText(Optional ByVal errorLabel As Object, ByVal text As String)
    On Error Resume Next
    If Not errorLabel Is Nothing Then
        errorLabel.Caption = text
        errorLabel.Visible = (Len(text) > 0)
    End If
End Sub

Private Function EnsureRole(ByVal roleName As String, ByVal isAdmin As Boolean) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb
    sql = "SELECT * FROM " & AUTH_TABLE_ROLES & " WHERE RoleName=" & SqlQ(roleName)
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)

    If rs.EOF Then
        rs.AddNew
        rs!RoleName = roleName
        rs!IsAdmin = isAdmin
        rs.Update
        rs.Bookmark = rs.LastModified
    End If

    EnsureRole = rs!RoleID
End Function

Private Function CreateUser(ByVal login As String, ByVal fullName As String, ByVal roleId As Long, ByVal isActive As Boolean, ByVal mustChangePassword As Boolean) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & AUTH_TABLE_USERS & " WHERE 1=0", dbOpenDynaset, dbSeeChanges)

    rs.AddNew
    rs!Login = login
    rs!FullName = fullName
    rs!RoleID = roleId
    rs!IsActive = isActive
    rs!MustChangePassword = mustChangePassword
    rs!FailedAttempts = 0
    rs!LockedUntil = Null
    rs!CreatedAt = Now()
    rs!LastLoginAt = Null
    rs!PasswordSalt = "seed"
    rs!PasswordHash = "seed"
    rs.Update

    rs.Bookmark = rs.LastModified
    CreateUser = rs!UserID
End Function

