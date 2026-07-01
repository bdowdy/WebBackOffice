-- =============================================================================
-- Create [Wait List] table in [MSaleData]
-- Run in SSMS. Safe to re-run.
-- =============================================================================

USE [MSaleData];
GO

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_CATALOG = 'MSaleData'
      AND TABLE_SCHEMA  = 'dbo'
      AND TABLE_NAME    = 'Wait List'
)
BEGIN
    CREATE TABLE [dbo].[Wait List] (
        [ID]              INT           IDENTITY(1,1) NOT NULL,
        [Day]             INT           NOT NULL DEFAULT 0,   -- 0 = today; archived rows get a date number
        [Date]            NVARCHAR(20)  NOT NULL DEFAULT '',
        [Seated]          BIT           NOT NULL DEFAULT 0,
        [Name]            NVARCHAR(100) NOT NULL DEFAULT '',
        [Number in Party] INT           NOT NULL DEFAULT 1,
        [Time In]         NVARCHAR(10)  NOT NULL DEFAULT '',
        [Preference]      NVARCHAR(50)  NOT NULL DEFAULT '',
        [Quoted]          INT           NOT NULL DEFAULT 0,   -- quoted wait minutes
        [Time Seated]     NVARCHAR(10)  NOT NULL DEFAULT '',
        [Pager]           NVARCHAR(30)  NOT NULL DEFAULT '',  -- cell phone
        [Group]           NVARCHAR(50)  NOT NULL DEFAULT 'ALL',
        [TextTime]        NVARCHAR(10)  NOT NULL DEFAULT '',  -- time text was sent

        CONSTRAINT [PK_WaitList] PRIMARY KEY CLUSTERED ([ID])
    );

    CREATE NONCLUSTERED INDEX [IX_WaitList_Day]
        ON [dbo].[Wait List] ([Day], [Seated]);

    PRINT 'Table [Wait List] created successfully.';
END
ELSE
    PRINT 'Table [Wait List] already exists — no changes made.';
GO
