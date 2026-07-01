/* ============================================================
   Receipt print size - clean storage for per-printer big/small
   Target: SQL Server, database [MSaleData]
   Run once. Safe to re-run.

   One row per printer:
     [Item Size] / [Modifier Size] = 'Large' or 'Small'
     [Big Esc] / [Small Esc]       = comma-separated decimal ESC bytes
                                     (e.g. '29,33,17' = GS ! 0x11 double,
                                           '29,33,0'  = GS ! 0x00 normal)

   VB6 read (pseudocode):
     Rs = "Select * from [Receipt Print Format] where [Printer Name] = '" & P & "'"
     BigEsc  = BytesFromCsv(Rs![Big Esc])      ' Chr(29)&Chr(33)&Chr(17)...
     SmallEsc = BytesFromCsv(Rs![Small Esc])
     ItemEsc = IIf(Rs![Item Size] = "Large", BigEsc, SmallEsc)
     ModEsc  = IIf(Rs![Modifier Size] = "Large", BigEsc, SmallEsc)
     ' wrap item lines with ItemEsc, modifier lines with ModEsc,
     ' then reset with SmallEsc (normal) after each line.
   ============================================================ */

USE [MSaleData];
GO

IF OBJECT_ID(N'[dbo].[Receipt Print Format]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Receipt Print Format](
        [Printer Name]  NVARCHAR(100) NOT NULL
            CONSTRAINT PK_ReceiptPrintFormat PRIMARY KEY,
        [Item Size]     NVARCHAR(10)  NOT NULL CONSTRAINT DF_RPF_Item     DEFAULT (N'Small'),
        [Modifier Size] NVARCHAR(10)  NOT NULL CONSTRAINT DF_RPF_Mod      DEFAULT (N'Small'),
        [Big Esc]       NVARCHAR(100) NOT NULL CONSTRAINT DF_RPF_Big      DEFAULT (N'29,33,17'),
        [Small Esc]     NVARCHAR(100) NOT NULL CONSTRAINT DF_RPF_Small    DEFAULT (N'29,33,0'),
        [Modified Utc]  DATETIME2     NOT NULL CONSTRAINT DF_RPF_Modified DEFAULT (SYSUTCDATETIME())
    );
END
GO
