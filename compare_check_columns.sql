/* Column lists for the live vs archive check tables.
   Run against the server; paste BOTH result sets back to Claude. */

-- LIVE table (currently-open checks)
SELECT COLUMN_NAME, DATA_TYPE
FROM [chk statsql].INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Open Checks'
ORDER BY ORDINAL_POSITION;

-- ARCHIVE table (for reference / confirming names match)
SELECT COLUMN_NAME, DATA_TYPE
FROM [chk statsql].INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Closed Checks'
ORDER BY ORDINAL_POSITION;
