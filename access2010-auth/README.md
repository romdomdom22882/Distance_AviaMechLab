## Проект «Авторизация в Microsoft Access 2010» (VBA + таблицы + формы)

Этот мини‑проект показывает, как сделать **вход по логину/паролю** в Access 2010:

- **пользователи и роли** (таблицы `tUsers`, `tRoles`);
- **проверка пароля** по хэшу (SHA‑256 через Windows CryptoAPI) + соль;
- **блокировка** пользователя при множественных ошибках входа;
- **аудит** попыток входа;
- **контекст текущего пользователя** через `TempVars`;
- шаблоны форм: `frmLogin` (вход) и `frmMain` (главная форма).

> Важно: Access не является идеальной платформой для «железобетонной» безопасности. Это учебная реализация для курсового/практики.

---

## Структура (что лежит в репозитории)

- `access2010-auth/schema/access-ddl.sql` — DDL (приближенный к Access SQL) для создания таблиц/индексов.
- `access2010-auth/vba/`
  - `modAuth.bas` — логика входа, блокировки, аудит.
  - `modCryptoApiSha256.bas` — SHA‑256 + соль (Windows CryptoAPI).
  - `modGlobals.bas` — общие константы/утилиты.

---

## Таблицы

### `tRoles`

- `RoleID` (Автонумерация, PK)
- `RoleName` (Короткий текст, уникальный) — например: `Admin`, `User`
- `IsAdmin` (Да/Нет)

### `tUsers`

- `UserID` (Автонумерация, PK)
- `Login` (Короткий текст, уникальный) — логин
- `PasswordSalt` (Короткий текст) — соль (hex)
- `PasswordHash` (Короткий текст) — SHA‑256(salt + password) (hex)
- `FullName` (Короткий текст)
- `RoleID` (Числовой, FK на `tRoles.RoleID`)
- `IsActive` (Да/Нет)
- `MustChangePassword` (Да/Нет)
- `FailedAttempts` (Числовой, по умолчанию 0)
- `LockedUntil` (Дата/время, может быть NULL)
- `CreatedAt` (Дата/время)
- `LastLoginAt` (Дата/время, может быть NULL)

### `tAuthLog` (аудит)

- `AuthLogID` (Автонумерация, PK)
- `Login` (Короткий текст)
- `UserID` (Числовой, может быть NULL)
- `Success` (Да/Нет)
- `ClientInfo` (Короткий текст) — `Environ("COMPUTERNAME")`, `Environ("USERNAME")`
- `EventAt` (Дата/время)
- `Message` (Короткий текст)

---

## Формы (как собрать в Access 2010)

### 1) Создайте базу

1. Access 2010 → **Создать** → пустая база, например `AuthDemo.accdb`.
2. Импортируйте/создайте таблицы из раздела DDL (ниже).

### 2) Добавьте VBA‑модули

1. `ALT+F11` (редактор VBA).
2. **File → Import File…** и импортируйте:
   - `access2010-auth/vba/modGlobals.bas`
   - `access2010-auth/vba/modCryptoApiSha256.bas`
   - `access2010-auth/vba/modAuth.bas`

### 3) Создайте форму входа `frmLogin`

Создайте форму (Конструктор), добавьте:

- Текстовое поле `txtLogin`
- Текстовое поле `txtPassword` (свойство **Input Mask** можно оставить пустым; «скрытие» делаем через `PasswordChar` в событии, либо через свойство `InputMask`/`Format` по вкусу)
- Метку `lblError`
- Кнопку `btnLogin`

Код в обработчик `btnLogin_Click`:

```vb
Private Sub btnLogin_Click()
    Dim ok As Boolean
    ok = Auth_Login(Me.txtLogin.Value, Me.txtPassword.Value, Me.lblError)
    If ok Then
        DoCmd.OpenForm "frmMain"
        DoCmd.Close acForm, Me.Name
    End If
End Sub
```

### 4) Главная форма `frmMain`

Сделайте любую «главную» форму и в `Form_Open`/`Form_Load` проверьте, что пользователь авторизован:

```vb
Private Sub Form_Open(Cancel As Integer)
    If Not Auth_IsLoggedIn() Then
        DoCmd.OpenForm "frmLogin"
        DoCmd.Close acForm, Me.Name
    End If
End Sub
```

---

## Инициализация (создать роли и первого админа)

Один раз выполните в Immediate Window (CTRL+G):

```vb
Call Auth_EnsureSeedData
```

По умолчанию создается:

- роль `Admin`
- пользователь `admin` с паролем `Admin123!` (и флагом `MustChangePassword = True`)

---

## DDL / запросы (Access SQL)

См. файл `access2010-auth/schema/access-ddl.sql`.

---

## Настройки безопасности (рекомендации)

- храните базу в защищенной папке, ограничьте права на файл;
- разделяйте «front‑end/back‑end» (таблицы в отдельном ACCDB на сервере/шаре);
- включайте параметр «Trust access to VBA project object model» только при необходимости;
- не используйте «общий» пароль для всех, включайте блокировку и аудит.

