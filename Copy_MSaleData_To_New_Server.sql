/* ================================================================
   Copy MSaleData from BENV11\MSale to a NEW server (backup & restore)
   Same database name on the target: MSaleData
   No shared folder -> back up locally, copy the .bak by hand, restore.

   Run PART 1 on the SOURCE (BENV11\MSale).
   Copy the .bak file to the target machine.
   Run PART 2 (step 2a then 2b) on the TARGET server.
   ================================================================ */


/* ----------------------------------------------------------------
   PART 1 - on the SOURCE server  (BENV11\MSale)
   COPY_ONLY = does not disturb your normal backup/log chain.
   Source database stays online the whole time.
   ---------------------------------------------------------------- */
BACKUP DATABASE [MSaleData]
TO DISK = N'C:\Temp\MSaleData.bak'
WITH COPY_ONLY,
     INIT, FORMAT,          -- overwrite any existing file at that path
     CHECKSUM,              -- verify page integrity while backing up
     STATS = 5,
     NAME = N'MSaleData-Full Copy';
     -- , COMPRESSION        -- add this if the instance is Standard/Enterprise
     --                        (SQL Express does NOT support backup compression)
GO

-- Optional but recommended: confirm the backup is readable before copying it.
RESTORE VERIFYONLY FROM DISK = N'C:\Temp\MSaleData.bak' WITH CHECKSUM;
GO

/*  Now copy  C:\Temp\MSaleData.bak  to the target machine
    (USB / RDP drive share / network copy), e.g. to C:\Temp\ there.  */


/* ----------------------------------------------------------------
   PART 2a - on the TARGET server: discover the logical file names
   Look at the "LogicalName" column in the results:
     - the data file  (Type = D)  -> probably  MSaleData
     - the log file   (Type = L)  -> probably  MSaleData_log
   Use those exact names in PART 2b below.
   ---------------------------------------------------------------- */
RESTORE FILELISTONLY FROM DISK = N'C:\Temp\MSaleData.bak';
GO


/* ----------------------------------------------------------------
   PART 2b - on the TARGET server: restore as MSaleData
   EDIT these before running:
     * the two logical names (from 2a) after each MOVE
     * the target folder paths for the .mdf and .ldf
       (use folders that exist on the target, e.g. the target's
        default SQL data/log directories)
   WARNING: WITH REPLACE will OVERWRITE an existing MSaleData on the
            target. Remove REPLACE if you do NOT want that.
   ---------------------------------------------------------------- */
RESTORE DATABASE [MSaleData]
FROM DISK = N'C:\Temp\MSaleData.bak'
WITH MOVE N'MSaleData'     TO N'C:\Program Files\Microsoft SQL Server\...\DATA\MSaleData.mdf',
     MOVE N'MSaleData_log' TO N'C:\Program Files\Microsoft SQL Server\...\DATA\MSaleData_log.ldf',
     REPLACE,
     RECOVERY,             -- bring the database online, ready to use
     CHECKSUM,
     STATS = 5;
GO


/* ----------------------------------------------------------------
   PART 2c - on the TARGET server: post-restore fixups
   ---------------------------------------------------------------- */
USE [MSaleData];
GO

-- If the app connects as 'sa', nothing to do (sa maps by fixed SID).
-- If MSaleData has its OWN SQL logins/users, they can be "orphaned"
-- after moving to a new server. List any orphaned users:
SELECT dp.name AS OrphanedUser, dp.type_desc
FROM sys.database_principals dp
LEFT JOIN sys.server_principals sp ON dp.sid = sp.sid
WHERE dp.authentication_type_desc = 'INSTANCE'   -- SQL logins
  AND sp.sid IS NULL
  AND dp.name NOT IN ('dbo', 'guest', 'sys', 'INFORMATION_SCHEMA');
GO

-- For each orphaned user, create/point it at a login on the target, e.g.:
--   CREATE LOGIN [msaleapp] WITH PASSWORD = 'xxxxx';   -- if the login doesn't exist yet
--   ALTER USER [msaleapp] WITH LOGIN = [msaleapp];      -- re-link user -> login
GO

-- Optional: keep the same compatibility level / recovery model as the source.
-- ALTER DATABASE [MSaleData] SET RECOVERY SIMPLE;   -- (or FULL, to match source)
GO
