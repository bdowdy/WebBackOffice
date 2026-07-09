/* ================================================================
   Copy MSaleData -> new server over the internet using Azure Blob
   (BACKUP TO URL on the source, RESTORE FROM URL on the target).
   No desktop / no SMB needed - only outbound HTTPS (443) + your
   remote SQL connections.

   Requires SQL Server 2016 or newer on BOTH servers (block-blob +
   SAS credential). If either is 2012/2014 see the note at the end.
   ================================================================ */

/* ----------------------------------------------------------------
   ONE-TIME AZURE SETUP (portal, ~2 min):
     1. Create a Storage account (e.g. msalebackups), Standard, LRS.
     2. In it, create a Blob CONTAINER (e.g. dbmove), Private access.
     3. Generate a SAS token for that container with permissions:
          Read, Add, Create, Write, List   (Delete optional)
        Set an expiry a few hours out. Copy the SAS string.
        IMPORTANT: remove the leading '?' from the SAS before using
        it as SECRET below (it must start with 'sv=').

   Fill these in everywhere they appear:
     Container URL: https://msalebackups.blob.core.windows.net/dbmove
     SAS SECRET   : sv=2023-...&ss=b&srt=co&sp=rwcl...&sig=...
   ---------------------------------------------------------------- */


/* ================================================================
   PART 1 - run on the SOURCE server  (connect to BENV11\MSale)
   ================================================================ */

-- Credential NAME must equal the CONTAINER URL exactly (no blob name).
IF NOT EXISTS (SELECT 1 FROM sys.credentials
               WHERE name = N'https://msalebackups.blob.core.windows.net/dbmove')
    CREATE CREDENTIAL [https://msalebackups.blob.core.windows.net/dbmove]
        WITH IDENTITY = N'SHARED ACCESS SIGNATURE',
             SECRET   = N'sv=PUT-YOUR-SAS-HERE-WITHOUT-LEADING-QUESTION-MARK';
GO

BACKUP DATABASE [MSaleData]
TO URL = N'https://msalebackups.blob.core.windows.net/dbmove/MSaleData.bak'
WITH COPY_ONLY, CHECKSUM, FORMAT, STATS = 5,
     NAME = N'MSaleData-Full Copy';
     -- , COMPRESSION   -- add on Standard/Enterprise; NOT on Express
GO


/* ================================================================
   PART 2 - run on the TARGET server  (connect to 47.206.96.203)
   ================================================================ */

-- Same credential (same container URL + same SAS secret) on the target.
IF NOT EXISTS (SELECT 1 FROM sys.credentials
               WHERE name = N'https://msalebackups.blob.core.windows.net/dbmove')
    CREATE CREDENTIAL [https://msalebackups.blob.core.windows.net/dbmove]
        WITH IDENTITY = N'SHARED ACCESS SIGNATURE',
             SECRET   = N'sv=PUT-YOUR-SAS-HERE-WITHOUT-LEADING-QUESTION-MARK';
GO

-- 2a. Discover logical file names (look at the LogicalName column).
RESTORE FILELISTONLY
FROM URL = N'https://msalebackups.blob.core.windows.net/dbmove/MSaleData.bak';
GO

-- 2b. Restore. Edit the logical names (from 2a) and the target's
--     local data/log folders (must already exist on the target).
--     WITH REPLACE overwrites an existing MSaleData on the target.
RESTORE DATABASE [MSaleData]
FROM URL = N'https://msalebackups.blob.core.windows.net/dbmove/MSaleData.bak'
WITH MOVE N'MSaleData'     TO N'C:\SQLData\MSaleData.mdf',
     MOVE N'MSaleData_log' TO N'C:\SQLData\MSaleData_log.ldf',
     REPLACE, RECOVERY, CHECKSUM, STATS = 5;
GO


/* ================================================================
   PART 3 - cleanup (both servers, after the restore succeeds)
   ================================================================ */
-- Remove the SAS credential so the token isn't left stored:
-- DROP CREDENTIAL [https://msalebackups.blob.core.windows.net/dbmove];
-- And delete MSaleData.bak from the container in the Azure portal.
-- If MSaleData has its own SQL logins, fix orphaned users on the
-- target (see Copy_MSaleData_To_New_Server.sql, PART 2c).
GO


/* ----------------------------------------------------------------
   IF EITHER SERVER IS SQL 2012 / 2014 (older BACKUP TO URL):
     Use a storage ACCESS KEY credential instead of SAS, e.g.:
       CREATE CREDENTIAL [AzureBackup]
         WITH IDENTITY = N'msalebackups',            -- storage acct name
              SECRET   = N'<storage-account-access-key>';
       BACKUP DATABASE [MSaleData]
         TO URL = N'https://msalebackups.blob.core.windows.net/dbmove/MSaleData.bak'
         WITH CREDENTIAL = N'AzureBackup', COPY_ONLY, STATS = 5;
     (2012/2014 use page blobs; keep backups under ~1 TB.)
   ---------------------------------------------------------------- */
