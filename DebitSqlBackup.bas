Attribute VB_Name = "DebitSqlBackup"
'==================================================================
' DebitSqlBackup.bas
' Backs up the debitsql database to spreadsheet-ready CSV files.
' One .csv file per table (each opens as its own sheet in Excel),
' written into a timestamped folder.
'
' Requires a project reference to:
'   Microsoft ActiveX Data Objects 2.x Library  (ADODB)
'
' Usage:
'   BackupDebitSql                       ' uses defaults below
'   BackupDebitSql "C:\Backups\Debit"    ' custom destination root
'   BackupDebitSqlTable "Transactions", "C:\Backups\Debit"  ' one table
'==================================================================
Option Explicit

' ---- EDIT THESE FOR YOUR ENVIRONMENT --------------------------------
Private Const CONN_STRING As String = _
    "Provider=SQLOLEDB;Data Source=BENV11\MSale;Initial Catalog=debitsql;" & _
    "User ID=sa;Password=MSale1221;"

' Where backup folders get created (a timestamped sub-folder is added).
Private Const DEFAULT_BACKUP_ROOT As String = "C:\Backups\DebitSql"
' ---------------------------------------------------------------------


'==================================================================
' Back up every user table in the database, one CSV per table.
'==================================================================
Public Sub BackupDebitSql(Optional ByVal backupRoot As String = "")
    On Error GoTo ErrorHandler

    Dim cn          As ADODB.Connection
    Dim rsTables    As ADODB.Recordset
    Dim destFolder  As String
    Dim tableName   As String
    Dim tableType   As String
    Dim fileCount   As Long

    If Len(backupRoot) = 0 Then backupRoot = DEFAULT_BACKUP_ROOT

    ' Create  <root>\Backup_YYYYMMDD_HHNNSS
    destFolder = backupRoot & "\Backup_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolder destFolder

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open CONN_STRING

    ' List base tables (skip views and system tables).
    Set rsTables = cn.OpenSchema(adSchemaTables)
    Do While Not rsTables.EOF
        tableType = rsTables.Fields("TABLE_TYPE").Value & ""
        If tableType = "TABLE" Then
            tableName = rsTables.Fields("TABLE_NAME").Value & ""
            ExportTable cn, tableName, destFolder & "\" & SafeFileName(tableName) & ".csv"
            fileCount = fileCount + 1
        End If
        rsTables.MoveNext
    Loop

    rsTables.Close: Set rsTables = Nothing
    cn.Close: Set cn = Nothing

    MsgBox "Backup complete." & vbCrLf & _
           fileCount & " table(s) written to:" & vbCrLf & destFolder, _
           vbInformation, "DebitSql Backup"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rsTables Is Nothing Then If rsTables.State = adStateOpen Then rsTables.Close
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    MsgBox "Backup failed: " & Err.Description, vbCritical, "DebitSql Backup"
End Sub


'==================================================================
' Back up a single named table to one CSV file.
'==================================================================
Public Sub BackupDebitSqlTable(ByVal tableName As String, _
                               Optional ByVal backupRoot As String = "")
    On Error GoTo ErrorHandler

    Dim cn         As ADODB.Connection
    Dim destFolder As String

    If Len(backupRoot) = 0 Then backupRoot = DEFAULT_BACKUP_ROOT
    destFolder = backupRoot & "\Backup_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolder destFolder

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open CONN_STRING

    ExportTable cn, tableName, destFolder & "\" & SafeFileName(tableName) & ".csv"

    cn.Close: Set cn = Nothing
    MsgBox "Backed up [" & tableName & "] to:" & vbCrLf & destFolder, _
           vbInformation, "DebitSql Backup"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    MsgBox "Backup of [" & tableName & "] failed: " & Err.Description, _
           vbCritical, "DebitSql Backup"
End Sub


'==================================================================
' Write one table's rows to a CSV file (header row + data rows).
'==================================================================
Private Sub ExportTable(ByRef cn As ADODB.Connection, _
                        ByVal tableName As String, _
                        ByVal filePath As String)
    Dim rs   As ADODB.Recordset
    Dim ff   As Integer
    Dim line As String
    Dim i    As Long

    Set rs = New ADODB.Recordset
    ' Brackets allow table names containing spaces.
    rs.Open "SELECT * FROM [" & tableName & "]", cn, adOpenForwardOnly, adLockReadOnly

    ff = FreeFile
    Open filePath For Output As #ff

    ' Header row
    line = ""
    For i = 0 To rs.Fields.Count - 1
        If i > 0 Then line = line & ","
        line = line & CsvField(rs.Fields(i).Name)
    Next i
    Print #ff, line

    ' Data rows
    Do While Not rs.EOF
        line = ""
        For i = 0 To rs.Fields.Count - 1
            If i > 0 Then line = line & ","
            line = line & CsvField(rs.Fields(i).Value)
        Next i
        Print #ff, line
        rs.MoveNext
    Loop

    Close #ff
    rs.Close: Set rs = Nothing
End Sub


'==================================================================
' Quote a value for CSV: wrap in quotes, double any internal quotes.
' Nulls become empty strings.
'==================================================================
Private Function CsvField(ByVal v As Variant) As String
    Dim s As String
    If IsNull(v) Then
        s = ""
    Else
        s = CStr(v)
    End If
    s = Replace(s, """", """""")
    CsvField = """" & s & """"
End Function


'==================================================================
' Create a folder path, building intermediate folders as needed.
'==================================================================
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim parts() As String
    Dim built   As String
    Dim i       As Long

    parts = Split(folderPath, "\")
    built = parts(0)                       ' drive, e.g. "C:"
    For i = 1 To UBound(parts)
        built = built & "\" & parts(i)
        If Len(Dir(built, vbDirectory)) = 0 Then MkDir built
    Next i
End Sub


'==================================================================
' Strip characters that are illegal in a Windows file name.
'==================================================================
Private Function SafeFileName(ByVal name As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        name = Replace(name, ch, "_")
    Next ch
    SafeFileName = name
End Function
