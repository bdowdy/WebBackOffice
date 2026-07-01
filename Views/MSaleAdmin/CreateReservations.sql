-- =============================================================================
-- Create [Reservations] table in [MSaleData]
-- Run in SSMS connected to your SQL Server instance.
-- Safe to re-run: uses IF NOT EXISTS check.
-- =============================================================================

USE [MSaleData];
GO

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_CATALOG = 'MSaleData'
      AND TABLE_SCHEMA  = 'dbo'
      AND TABLE_NAME    = 'Reservations'
)
BEGIN
    CREATE TABLE [dbo].[Reservations] (
        [Reservation ID]   INT           IDENTITY(1,1) NOT NULL,
        [Party Name]       NVARCHAR(100) NOT NULL,
        [Party Size]       INT           NOT NULL DEFAULT 1,
        [Reservation Date] DATE          NULL,
        [Reservation Time] VARCHAR(10)   NULL,   -- stored as HH:MM (24-hr)
        [Phone]            NVARCHAR(30)  NULL,
        [Notes]            NVARCHAR(500) NULL,
        [Table Number]     NVARCHAR(50)  NULL,
        [Status]           NVARCHAR(20)  NOT NULL DEFAULT 'Pending',
        [Created]          DATETIME      NOT NULL DEFAULT GETDATE(),

        CONSTRAINT [PK_Reservations] PRIMARY KEY CLUSTERED ([Reservation ID]),
        CONSTRAINT [CK_Reservations_Status]
            CHECK ([Status] IN ('Pending', 'Seated', 'Cancelled', 'No-Show'))
    );

    PRINT 'Table [Reservations] created successfully.';
END
ELSE
BEGIN
    PRINT 'Table [Reservations] already exists — no changes made.';
END
GO

-- Optional: index for fast date lookups
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('[dbo].[Reservations]')
      AND name = 'IX_Reservations_Date'
)
    CREATE NONCLUSTERED INDEX [IX_Reservations_Date]
        ON [dbo].[Reservations] ([Reservation Date], [Status]);
GO

PRINT 'Done.';
GO
