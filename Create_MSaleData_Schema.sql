/* ================================================================
   Create MSaleData on a new server - FULL SCHEMA
   ================================================================

   You cannot reliably hand-write the whole MicroSale schema; generate
   it from your existing MSaleData. Two ways below.

   -------------------------------------------------------------
   OPTION A (recommended) - SSMS "Generate Scripts" wizard
   -------------------------------------------------------------
   On the server that HAS MSaleData:
     1. Object Explorer -> right-click [MSaleData]
        -> Tasks -> Generate Scripts...
     2. Choose "Script entire database and all database objects"
        (or pick just Tables).
     3. Set Scripting Options -> Advanced:
          - "Types of data to script" = Schema only
          - "Script Indexes" = True
          - "Script Primary/Foreign/Unique Keys" = True
          - "Script Triggers" = True (if you use any)
          - "Script USE DATABASE" = False   (so you control the DB)
     4. Save to a .sql file.
     5. On the NEW server, first run the CREATE DATABASE in OPTION B
        step 1 below, then run the generated .sql against MSaleData.
   This captures columns, keys, indexes, defaults, triggers, and
   correct dependency order - the safest full copy of the structure.

   NOTE: If you just want an exact clone (schema + data), the backup/
   restore in Copy_MSaleData_To_New_Server.sql is simpler than
   scripting. Use this schema route when you want structure only.
   ================================================================ */


/* -------------------------------------------------------------
   OPTION B step 1 - create the empty database on the NEW server.
   Edit the file paths to folders that exist on the target.
   ------------------------------------------------------------- */
IF DB_ID(N'MSaleData') IS NULL
BEGIN
    CREATE DATABASE [MSaleData]
    ON PRIMARY
    (
        NAME = N'MSaleData',
        FILENAME = N'C:\SQLData\MSaleData.mdf',
        SIZE = 100MB, FILEGROWTH = 64MB
    )
    LOG ON
    (
        NAME = N'MSaleData_log',
        FILENAME = N'C:\SQLData\MSaleData_log.ldf',
        SIZE = 50MB, FILEGROWTH = 64MB
    );
END
GO

-- Match your source settings as needed:
-- ALTER DATABASE [MSaleData] SET RECOVERY SIMPLE;
-- ALTER DATABASE [MSaleData] SET COMPATIBILITY_LEVEL = 150;
GO


/* =================================================================
   OPTION B step 2 - T-SQL SCHEMA GENERATOR
   Run this ON THE SOURCE server, in the MSaleData database.
   It PRINTS CREATE TABLE + PK + indexes + foreign keys for every
   user table. Copy the Messages output and run it on the new server
   (against MSaleData). Good for automation; the SSMS wizard in
   OPTION A is more complete (triggers, extended properties, etc.).

   Tips before running:
     - Results To Text (Ctrl+T), and raise the output limit:
       Tools > Options > Query Results > SQL Server > Results to Text
       > "Maximum number of characters displayed" = 8192.
   ================================================================= */
USE [MSaleData];
GO
SET NOCOUNT ON;

/* ---- 1. Tables + columns ---- */
DECLARE @tables TABLE (id INT IDENTITY, obj INT, sname SYSNAME, tname SYSNAME);
INSERT INTO @tables (obj, sname, tname)
SELECT t.object_id, s.name, t.name
FROM sys.tables t
JOIN sys.schemas s ON s.schema_id = t.schema_id
WHERE t.is_ms_shipped = 0
ORDER BY s.name, t.name;

DECLARE @i INT = 1, @max INT = (SELECT COUNT(*) FROM @tables);
DECLARE @obj INT, @sn SYSNAME, @tn SYSNAME, @ddl NVARCHAR(MAX);

WHILE @i <= @max
BEGIN
    SELECT @obj = obj, @sn = sname, @tn = tname FROM @tables WHERE id = @i;

    SET @ddl = 'CREATE TABLE [' + @sn + '].[' + @tn + '] (' + CHAR(13);

    ;WITH cols AS (
        SELECT
            c.column_id,
            line =
              '    [' + c.name + '] ' +
              CASE
                WHEN c.is_computed = 1 THEN
                    'AS ' + ISNULL((SELECT definition FROM sys.computed_columns cc
                                    WHERE cc.object_id = c.object_id AND cc.column_id = c.column_id), '')
                ELSE
                    UPPER(ty.name) +
                    CASE
                        WHEN ty.name IN ('varchar','char','varbinary','binary')
                            THEN '(' + IIF(c.max_length = -1, 'MAX', CAST(c.max_length AS VARCHAR(10))) + ')'
                        WHEN ty.name IN ('nvarchar','nchar')
                            THEN '(' + IIF(c.max_length = -1, 'MAX', CAST(c.max_length/2 AS VARCHAR(10))) + ')'
                        WHEN ty.name IN ('decimal','numeric')
                            THEN '(' + CAST(c.precision AS VARCHAR(10)) + ',' + CAST(c.scale AS VARCHAR(10)) + ')'
                        WHEN ty.name IN ('datetime2','datetimeoffset','time')
                            THEN '(' + CAST(c.scale AS VARCHAR(10)) + ')'
                        ELSE ''
                    END +
                    CASE WHEN c.is_identity = 1
                         THEN ' IDENTITY(' +
                              CAST(IDENT_SEED(@sn+'.'+@tn) AS VARCHAR(20)) + ',' +
                              CAST(IDENT_INCR(@sn+'.'+@tn) AS VARCHAR(20)) + ')'
                         ELSE '' END +
                    CASE WHEN c.is_nullable = 0 THEN ' NOT NULL' ELSE ' NULL' END +
                    ISNULL(' CONSTRAINT [' + dc.name + '] DEFAULT ' + dc.definition, '')
              END
        FROM sys.columns c
        JOIN sys.types ty ON ty.user_type_id = c.user_type_id
        LEFT JOIN sys.default_constraints dc ON dc.object_id = c.default_object_id
        WHERE c.object_id = @obj
    )
    SELECT @ddl = @ddl + STRING_AGG(line, ',' + CHAR(13)) WITHIN GROUP (ORDER BY column_id) FROM cols;

    SET @ddl = @ddl + CHAR(13) + ');';
    PRINT @ddl;
    PRINT 'GO';

    SET @i = @i + 1;
END

/* ---- 2. Primary keys & unique constraints ---- */
DECLARE @keys NVARCHAR(MAX) = '';
SELECT @keys = @keys +
    'ALTER TABLE [' + s.name + '].[' + t.name + '] ADD CONSTRAINT [' + kc.name + '] ' +
    CASE WHEN kc.type = 'PK' THEN 'PRIMARY KEY ' ELSE 'UNIQUE ' END +
    CASE WHEN i.type = 1 THEN 'CLUSTERED ' ELSE 'NONCLUSTERED ' END + '(' +
    STUFF((SELECT ', [' + c2.name + ']' +
                  CASE WHEN ic2.is_descending_key = 1 THEN ' DESC' ELSE ' ASC' END
           FROM sys.index_columns ic2
           JOIN sys.columns c2 ON c2.object_id = ic2.object_id AND c2.column_id = ic2.column_id
           WHERE ic2.object_id = i.object_id AND ic2.index_id = i.index_id
           ORDER BY ic2.key_ordinal
           FOR XML PATH('')), 1, 2, '') + ');' + CHAR(13) + 'GO' + CHAR(13)
FROM sys.key_constraints kc
JOIN sys.tables t ON t.object_id = kc.parent_object_id
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.indexes i ON i.object_id = kc.parent_object_id AND i.index_id = kc.unique_index_id
WHERE t.is_ms_shipped = 0;
PRINT @keys;

/* ---- 3. Non-key indexes ---- */
DECLARE @ix NVARCHAR(MAX) = '';
SELECT @ix = @ix +
    'CREATE ' + CASE WHEN i.is_unique = 1 THEN 'UNIQUE ' ELSE '' END +
    CASE WHEN i.type = 1 THEN 'CLUSTERED ' ELSE 'NONCLUSTERED ' END +
    'INDEX [' + i.name + '] ON [' + s.name + '].[' + t.name + '] (' +
    STUFF((SELECT ', [' + c2.name + ']' +
                  CASE WHEN ic2.is_descending_key = 1 THEN ' DESC' ELSE ' ASC' END
           FROM sys.index_columns ic2
           JOIN sys.columns c2 ON c2.object_id = ic2.object_id AND c2.column_id = ic2.column_id
           WHERE ic2.object_id = i.object_id AND ic2.index_id = i.index_id AND ic2.is_included_column = 0
           ORDER BY ic2.key_ordinal
           FOR XML PATH('')), 1, 2, '') + ');' + CHAR(13) + 'GO' + CHAR(13)
FROM sys.indexes i
JOIN sys.tables t ON t.object_id = i.object_id
JOIN sys.schemas s ON s.schema_id = t.schema_id
WHERE t.is_ms_shipped = 0
  AND i.is_primary_key = 0 AND i.is_unique_constraint = 0
  AND i.type IN (1,2) AND i.name IS NOT NULL;
PRINT @ix;

/* ---- 4. Foreign keys (run LAST, after all tables exist) ---- */
DECLARE @fk NVARCHAR(MAX) = '';
SELECT @fk = @fk +
    'ALTER TABLE [' + s.name + '].[' + t.name + '] ADD CONSTRAINT [' + fk.name + '] FOREIGN KEY (' +
    STUFF((SELECT ', [' + pc.name + ']'
           FROM sys.foreign_key_columns fkc
           JOIN sys.columns pc ON pc.object_id = fkc.parent_object_id AND pc.column_id = fkc.parent_column_id
           WHERE fkc.constraint_object_id = fk.object_id ORDER BY fkc.constraint_column_id
           FOR XML PATH('')), 1, 2, '') + ') REFERENCES [' +
    rs.name + '].[' + rt.name + '] (' +
    STUFF((SELECT ', [' + rc.name + ']'
           FROM sys.foreign_key_columns fkc
           JOIN sys.columns rc ON rc.object_id = fkc.referenced_object_id AND rc.column_id = fkc.referenced_column_id
           WHERE fkc.constraint_object_id = fk.object_id ORDER BY fkc.constraint_column_id
           FOR XML PATH('')), 1, 2, '') + ');' + CHAR(13) + 'GO' + CHAR(13)
FROM sys.foreign_keys fk
JOIN sys.tables t  ON t.object_id  = fk.parent_object_id
JOIN sys.schemas s ON s.schema_id  = t.schema_id
JOIN sys.tables rt ON rt.object_id = fk.referenced_object_id
JOIN sys.schemas rs ON rs.schema_id = rt.schema_id;
PRINT @fk;
GO
