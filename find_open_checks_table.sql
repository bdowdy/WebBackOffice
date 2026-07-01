/* Find the table that holds currently-open / live checks.
   Run against MSaleData. Paste the list back to Claude. */

SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
  AND (TABLE_NAME LIKE '%Check%'
       OR TABLE_NAME LIKE '%Open%'
       OR TABLE_NAME LIKE '%Order%'
       OR TABLE_NAME LIKE '%Tab%')
ORDER BY TABLE_NAME;
