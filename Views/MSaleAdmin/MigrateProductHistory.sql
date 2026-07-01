-- =============================================================================
-- Migrate [Product HistorySQL] tables and data into [MSaleData].[dbo]
-- Run this in SSMS connected to your SQL Server instance.
-- Safe to re-run: uses IF NOT EXISTS checks before creating tables.
-- =============================================================================

USE [MSaleData];
GO

-- =============================================================================
-- STEP 1: Preview — list all tables in [Product HistorySQL]
--         Review this list before running Step 2.
-- =============================================================================
SELECT
    TABLE_NAME,
    TABLE_TYPE
FROM   [Product HistorySQL].INFORMATION_SCHEMA.TABLES
WHERE  TABLE_TYPE = 'BASE TABLE'
ORDER  BY TABLE_NAME;
GO


-- =============================================================================
-- STEP 2: Copy every table from [Product HistorySQL] into [MSaleData].[dbo]
--
--   For each table this script will:
--     a) Skip if the table already exists in [MSaleData].[dbo]
--     b) Create the table with the same columns (via SELECT INTO)
--     c) Copy all rows
--
--   Run the block below as a single batch after reviewing Step 1 output.
-- =============================================================================
DECLARE @tableName  NVARCHAR(256);
DECLARE @sql        NVARCHAR(MAX);

DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
    SELECT TABLE_NAME
    FROM   [Product HistorySQL].INFORMATION_SCHEMA.TABLES
    WHERE  TABLE_TYPE = 'BASE TABLE'
    ORDER  BY TABLE_NAME;

OPEN cur;
FETCH NEXT FROM cur INTO @tableName;

WHILE @@FETCH_STATUS = 0
BEGIN
    -- Only copy if the table does not already exist in MSaleData
    IF NOT EXISTS (
        SELECT 1
        FROM   INFORMATION_SCHEMA.TABLES
        WHERE  TABLE_CATALOG = 'MSaleData'
          AND  TABLE_SCHEMA  = 'dbo'
          AND  TABLE_NAME    = @tableName
    )
    BEGIN
        -- CREATE + populate in one shot using SELECT INTO
        SET @sql = N'SELECT * INTO [MSaleData].[dbo].[' + @tableName + N'] '
                 + N'FROM [Product HistorySQL].[dbo].[' + @tableName + N'];';

        PRINT 'Creating and populating: ' + @tableName;
        EXEC sp_executesql @sql;
    END
    ELSE
    BEGIN
        PRINT 'SKIPPED (already exists): ' + @tableName;
    END

    FETCH NEXT FROM cur INTO @tableName;
END

CLOSE cur;
DEALLOCATE cur;
GO

PRINT '-- Migration complete. Run Step 3 to verify row counts.';
GO


-- =============================================================================
-- STEP 3: Verify — compare row counts between source and destination
-- =============================================================================
DECLARE @tableName  NVARCHAR(256);
DECLARE @sql        NVARCHAR(MAX);

DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
    SELECT TABLE_NAME
    FROM   [Product HistorySQL].INFORMATION_SCHEMA.TABLES
    WHERE  TABLE_TYPE = 'BASE TABLE'
    ORDER  BY TABLE_NAME;

CREATE TABLE #MigrationCheck (
    TableName       NVARCHAR(256),
    SourceRows      INT,
    DestRows        INT,
    Match           BIT
);

OPEN cur;
FETCH NEXT FROM cur INTO @tableName;

WHILE @@FETCH_STATUS = 0
BEGIN
    DECLARE @srcCount  INT = 0;
    DECLARE @dstCount  INT = 0;

    SET @sql = N'SELECT @c = COUNT(*) FROM [Product HistorySQL].[dbo].[' + @tableName + N']';
    EXEC sp_executesql @sql, N'@c INT OUTPUT', @c = @srcCount OUTPUT;

    IF EXISTS (
        SELECT 1 FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_CATALOG = 'MSaleData' AND TABLE_SCHEMA = 'dbo' AND TABLE_NAME = @tableName
    )
    BEGIN
        SET @sql = N'SELECT @c = COUNT(*) FROM [MSaleData].[dbo].[' + @tableName + N']';
        EXEC sp_executesql @sql, N'@c INT OUTPUT', @c = @dstCount OUTPUT;
    END

    INSERT INTO #MigrationCheck VALUES (
        @tableName,
        @srcCount,
        @dstCount,
        CASE WHEN @srcCount = @dstCount THEN 1 ELSE 0 END
    );

    FETCH NEXT FROM cur INTO @tableName;
END

CLOSE cur;
DEALLOCATE cur;

SELECT
    TableName,
    SourceRows,
    DestRows,
    CASE WHEN Match = 1 THEN 'OK' ELSE '*** MISMATCH ***' END AS Status
FROM #MigrationCheck
ORDER BY TableName;

DROP TABLE #MigrationCheck;
GO


-- =============================================================================
-- STEP 4: Fix [HouseCharges] view / object that references [houseaccount].dbo
--
-- First, find exactly what references houseaccount:
-- =============================================================================
SELECT
    o.type_desc  AS ObjectType,
    o.name       AS ObjectName,
    m.definition AS Definition
FROM   sys.sql_modules m
JOIN   sys.objects     o ON o.object_id = m.object_id
WHERE  m.definition LIKE '%houseaccount%'
   OR  m.definition LIKE '%Product HistorySQL%'
ORDER  BY o.type_desc, o.name;
GO

-- =============================================================================
-- After reviewing the output above, uncomment and edit the ALTER VIEW below
-- replacing the old cross-database reference with [MSaleData].[dbo]:
--
-- Example — if HouseCharges was a view pointing at [houseaccount].dbo.HouseCharges:
--
-- ALTER VIEW [dbo].[HouseCharges] AS
--     SELECT * FROM [MSaleData].[dbo].[HouseCharges_Data];   -- adjust as needed
-- GO
-- =============================================================================
