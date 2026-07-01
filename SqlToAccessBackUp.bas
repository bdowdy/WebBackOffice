Attribute VB_Name = "SqlToAccessBackup"
'==================================================================
' Driver for the improved GeneralLib.Utilities.BackupSqlTableToAccess.
' Loops the tables in the Access backup file and snapshots each one
' from SQL Server. The new routine drops + recreates each table, so
' there is NO Delete step here and table names are passed un-bracketed.
'==================================================================
Option Explicit

Public Sub SqlToAccessBackUp(MDBPathandFile As String, MDBFile As String, _
                             PathandFile As String, SQLFile As String)
    On Error GoTo ErrorHandler

    Dim lDb        As ADODB.Connection
    Dim lRs        As ADODB.Recordset
    Dim SRS        As ADODB.Recordset
    Dim GL         As GeneralLib.Utilities
    Dim TableNames As Collection
    Dim TableName  As Variant
    Dim ServerName As String
    Dim LogResult  As String

    Set GL = New GeneralLib.Utilities

    ' Connect to the Access (.mdb/.accdb) backup file.
    If ConnectADODB(MDBPathandFile, lDb) = False Then
        Trans_Log " Back Up " & MDBPathandFile & " Failed to Connect "
        GoTo CleanUp
    End If

    ServerName = ReturnServerName(PathandFile, SQLFile)

    ' 1) Collect the table names first, so we don't walk the live schema
    '    rowset while BackupSqlTableToAccess drops/recreates each table.
    Set TableNames = New Collection
    Set SRS = lDb.OpenSchema(adSchemaTables)
    Do While Not SRS.EOF
        If SRS.Fields("Table_Type").Value = "TABLE" Then
            TableNames.Add SRS.Fields("Table_Name").Value
        End If
        SRS.MoveNext
    Loop
    SetNothing SRS

    ' 2) Snapshot each table. No Delete needed (the routine recreates it),
    '    and the name is passed RAW - the routine brackets it internally.
    For Each TableName In TableNames
        LogResult = GL.BackupSqlTableToAccess(CStr(TableName), _
                                              ServerName, _
                                              SQLFile, _
                                              MDBPathandFile, _
                                              MDBFile)
        Trans_Log " Back Up " & LogResult & " [" & TableName & "]"
    Next TableName

CleanUp:
    On Error Resume Next
    Call DbClose(lDb, lRs)
    Set GL = Nothing
    Exit Sub

ErrorHandler:
    Trans_Log " Back Up error: " & Err.Description
    Resume CleanUp
End Sub
