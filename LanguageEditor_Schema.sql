/* ============================================================
   Language editor - schema setup
   Target: SQL Server, database [MSaleData]
   Run once before using /MSaleAdmin/LanguageEditor.
   Safe to re-run.
   ============================================================ */

USE [MSaleData];
GO

/* 1. De-duplicate [Language] on [English] so a unique key can be
      created. Column-name-agnostic (the translation column was renamed
      during the Jet -> SQL migration), so it just keeps one row per
      English value. */
;WITH ranked AS (
    SELECT [English],
           rn = ROW_NUMBER() OVER (PARTITION BY [English] ORDER BY (SELECT 1))
    FROM [dbo].[Language]
)
DELETE FROM ranked WHERE rn > 1;
GO

/* 2. Enforce English as the key (the upsert MERGE relies on it). */
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE name = N'UX_Language_English'
      AND object_id = OBJECT_ID(N'[dbo].[Language]'))
BEGIN
    CREATE UNIQUE INDEX [UX_Language_English] ON [dbo].[Language] ([English]);
END
GO

/* 3. Settings table to hold the "Language On" flag that used to live
      in the Windows registry (SaveSetting "Micro$ale","Register Setup",
      "Language On"). Generic name/value for reuse by other settings. */
IF OBJECT_ID(N'[dbo].[AppSetting]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[AppSetting](
        [SettingName]  NVARCHAR(100) NOT NULL PRIMARY KEY,
        [SettingValue] NVARCHAR(400) NULL,
        [ModifiedUtc]  DATETIME2     NOT NULL
            CONSTRAINT DF_AppSetting_Modified DEFAULT (SYSUTCDATETIME())
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[AppSetting] WHERE [SettingName] = N'Language On')
    INSERT INTO [dbo].[AppSetting] ([SettingName], [SettingValue]) VALUES (N'Language On', N'0');
GO
