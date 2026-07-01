/* ============================================================
   Employee sync source table for the EmployeeSyncAgent.
   Run against MSaleData.

   The /api/sync/employees/changes endpoint selects:
       Id, FirstName, LastName, Pin, RoleName, JobCode,
       IsActive, Email, Phone, ModifiedUtc
   and filters on (ModifiedUtc > @since), so ModifiedUtc must be
   maintained whenever a row changes.
   ============================================================ */

IF OBJECT_ID('dbo.Employees', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Employees
    (
        Id           INT            IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Employees PRIMARY KEY,
        FirstName    NVARCHAR(100)  NOT NULL CONSTRAINT DF_Employees_FirstName DEFAULT(''),
        LastName     NVARCHAR(100)  NOT NULL CONSTRAINT DF_Employees_LastName  DEFAULT(''),
        Pin          NVARCHAR(20)   NULL,
        RoleName     NVARCHAR(100)  NULL,
        JobCode      NVARCHAR(50)   NULL,
        IsActive     BIT            NOT NULL CONSTRAINT DF_Employees_IsActive   DEFAULT(1),
        Email        NVARCHAR(256)  NULL,
        Phone        NVARCHAR(50)   NULL,
        ModifiedUtc  DATETIME2(3)   NOT NULL CONSTRAINT DF_Employees_ModifiedUtc DEFAULT(SYSUTCDATETIME())
    );

    -- Incremental pulls filter/sort on ModifiedUtc.
    CREATE INDEX IX_Employees_ModifiedUtc ON dbo.Employees (ModifiedUtc);
END
GO

/* Keep ModifiedUtc current on every UPDATE so incremental sync
   (ModifiedUtc > @since) picks up edits automatically. */
IF OBJECT_ID('dbo.trg_Employees_ModifiedUtc', 'TR') IS NOT NULL
    DROP TRIGGER dbo.trg_Employees_ModifiedUtc;
GO
CREATE TRIGGER dbo.trg_Employees_ModifiedUtc
ON dbo.Employees
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE e
        SET ModifiedUtc = SYSUTCDATETIME()
    FROM dbo.Employees e
    INNER JOIN inserted i ON e.Id = i.Id;
END
GO

/* ------------------------------------------------------------
   Optional: the agent's WriteSyncLog() inserts into a SyncLog
   table (errors are swallowed if it's missing, but creating it
   lets you see the agent's run history).
   ------------------------------------------------------------ */
IF OBJECT_ID('dbo.SyncLog', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.SyncLog
    (
        Id         INT            IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_SyncLog PRIMARY KEY,
        SyncType   NVARCHAR(100)  NULL,
        LogLevel   NVARCHAR(20)   NULL,
        Message    NVARCHAR(MAX)  NULL,
        CreatedUtc DATETIME2(3)   NOT NULL CONSTRAINT DF_SyncLog_CreatedUtc DEFAULT(SYSUTCDATETIME())
    );
END
GO

/* ------------------------------------------------------------
   Sample row so you can confirm the endpoint returns data:
   GET https://localhost:7259/api/sync/employees/changes?storeId=1
   ------------------------------------------------------------ */
-- INSERT INTO dbo.Employees (FirstName, LastName, Pin, RoleName, JobCode, IsActive, Email, Phone)
-- VALUES (N'Test', N'Server', N'1234', N'Server', N'SRV', 1, N'test@store.com', N'813-555-0100');
