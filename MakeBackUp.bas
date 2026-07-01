Attribute VB_Name = "MakeBackup"
'==================================================================
' Copies every table from a source database into a backup database.
'  - Clears each destination table first  -> clean snapshot, no dupes
'  - One transaction + batched insert per table  -> faster, atomic
'  - Per-table error handling that LOGS and CONTINUES (one bad table
'    no longer silently kills the whole backup)
'  - Copies only columns that exist on the source and are writable
'==================================================================
Option Explicit

Public Sub MakeBackUp(FromPath As String, GetDataFrom As String, _
                      SendToPath As String, BackFileName As String)
    On Error GoTo FatalHandler

    Dim EDb        As ADODB.Connection     ' source (live data)
    Dim lDb        As ADODB.Connection     ' destination (backup file)
    Dim SRS        As ADODB.Recordset
    Dim TableNames As Collection
    Dim TableName  As Variant
    Dim copied     As Long
    Dim failed     As Long

    If ConnectADODB(FromPath & GetDataFrom, EDb) = False Then
        Trans_Log "MakeBackUp: cannot connect to source " & FromPath & GetDataFrom
        Exit Sub
    End If
    If ConnectADODB(SendToPath & BackFileName, lDb) = False Then
        Trans_Log "MakeBackUp: cannot connect to backup " & SendToPath & BackFileName
        DbClose EDb, Nothing
        Exit Sub
    End If

    ' Collect names first - don't walk the schema rowset while we mutate tables.
    Set TableNames = New Collection
    Set SRS = lDb.OpenSchema(adSchemaTables)
    Do While Not SRS.EOF
        If SRS.Fields("Table_Type").Value = "TABLE" Then
            TableNames.Add SRS.Fields("Table_Name").Value
        End If
        SRS.MoveNext
    Loop
    SRS.Close: Set SRS = Nothing

    For Each TableName In TableNames
        If CopyTable(EDb, lDb, CStr(TableName)) Then
            copied = copied + 1
        Else
            failed = failed + 1
        End If
    Next TableName

    DbClose lDb, Nothing
    DbClose EDb, Nothing
    Trans_Log "MakeBackUp finished: " & copied & " table(s) copied, " & failed & " failed"
    Exit Sub

FatalHandler:
    Error_Log Err.Number, Err.Description, "Weekly Backup", "None"
    On Error Resume Next
    If Not lDb Is Nothing Then DbClose lDb, Nothing
    If Not EDb Is Nothing Then DbClose EDb, Nothing
End Sub


'==================================================================
' Copy one table. Returns True on success. Isolates failure so a
' single problem table doesn't abort the entire backup.
'==================================================================
Private Function CopyTable(ByRef src As ADODB.Connection, _
                           ByRef dst As ADODB.Connection, _
                           ByVal TableName As String) As Boolean
    Dim sRs     As ADODB.Recordset
    Dim dRs     As ADODB.Recordset
    Dim fld     As ADODB.Field
    Dim inTrans As Boolean
    Dim t       As String

    On Error GoTo Failed

    t = "[" & Replace(TableName, "]", "]]") & "]"

    ' Clean snapshot: empty the destination table first.
    dst.Execute "DELETE FROM " & t

    ' Source: fast read-only forward cursor.
    Set sRs = New ADODB.Recordset
    sRs.CursorLocation = adUseClient
    sRs.Open "SELECT * FROM " & t, src, adOpenForwardOnly, adLockReadOnly

    ' Destination: empty, appendable, batched on a client cursor.
    Set dRs = New ADODB.Recordset
    dRs.CursorLocation = adUseClient
    dRs.Open "SELECT * FROM " & t & " WHERE 1 = 0", dst, adOpenStatic, adLockBatchOptimistic

    dst.BeginTrans
    inTrans = True

    Do While Not sRs.EOF
        dRs.AddNew
        For Each fld In dRs.Fields
            ' Only copy columns that exist on the source AND are writable
            ' (skips identity / computed / read-only columns).
            If (fld.Attributes And adFldUpdatable) <> 0 Then
                If FieldExists(sRs, fld.Name) Then
                    dRs.Fields(fld.Name).Value = sRs.Fields(fld.Name).Value
                End If
            End If
        Next fld
        sRs.MoveNext
    Loop

    dRs.UpdateBatch            ' single batched write -> much faster than per-row
    dst.CommitTrans
    inTrans = False

    sRs.Close: Set sRs = Nothing
    dRs.Close: Set dRs = Nothing
    CopyTable = True
    Exit Function

Failed:
    Trans_Log "MakeBackUp: table [" & TableName & "] failed " & Err.Description
    On Error Resume Next
    If inTrans Then dst.RollbackTrans
    If Not sRs Is Nothing Then sRs.Close
    If Not dRs Is Nothing Then dRs.Close
    CopyTable = False
End Function


'==================================================================
' True if the recordset has a field with the given name.
'==================================================================
Private Function FieldExists(ByRef rs As ADODB.Recordset, ByVal name As String) As Boolean
    On Error GoTo NoField
    Dim s As String
    s = rs.Fields(name).name
    FieldExists = True
    Exit Function
NoField:
    FieldExists = False
End Function
