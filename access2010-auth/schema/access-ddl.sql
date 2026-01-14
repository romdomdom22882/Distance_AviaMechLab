-- Access 2010 SQL (ACE/Jet). Выполняйте запросы по частям в "Создание запроса" → "Режим SQL".
-- Примечание: В некоторых конфигурациях Access может ругаться на CONSTRAINT/FK в DDL.
-- Тогда создайте связи через "Схема данных" вручную.

-- 1) Роли
CREATE TABLE tRoles (
  RoleID AUTOINCREMENT CONSTRAINT PK_tRoles PRIMARY KEY,
  RoleName TEXT(50) NOT NULL,
  IsAdmin YESNO NOT NULL
);

CREATE UNIQUE INDEX UX_tRoles_RoleName ON tRoles (RoleName);

-- 2) Пользователи
CREATE TABLE tUsers (
  UserID AUTOINCREMENT CONSTRAINT PK_tUsers PRIMARY KEY,
  [Login] TEXT(50) NOT NULL,
  PasswordSalt TEXT(128) NOT NULL,
  PasswordHash TEXT(128) NOT NULL,
  FullName TEXT(100),
  RoleID LONG NOT NULL,
  IsActive YESNO NOT NULL,
  MustChangePassword YESNO NOT NULL,
  FailedAttempts LONG NOT NULL,
  LockedUntil DATETIME,
  CreatedAt DATETIME NOT NULL,
  LastLoginAt DATETIME
);

CREATE UNIQUE INDEX UX_tUsers_Login ON tUsers ([Login]);
CREATE INDEX IX_tUsers_RoleID ON tUsers (RoleID);

-- 3) Аудит входов
CREATE TABLE tAuthLog (
  AuthLogID AUTOINCREMENT CONSTRAINT PK_tAuthLog PRIMARY KEY,
  [Login] TEXT(50),
  UserID LONG,
  Success YESNO NOT NULL,
  ClientInfo TEXT(255),
  EventAt DATETIME NOT NULL,
  [Message] TEXT(255)
);

CREATE INDEX IX_tAuthLog_EventAt ON tAuthLog (EventAt);
CREATE INDEX IX_tAuthLog_UserID ON tAuthLog (UserID);

