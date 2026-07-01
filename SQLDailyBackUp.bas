Attribute VB_Name = "SQLDailyBackup"
'==================================================================
' Hardened daily SQL Server backup for MSaleData.
'  - Native BACKUP DATABASE to a LOCAL, date-stamped .bak
'  - WITH CHECKSUM, then RESTORE VERIFYONLY to prove it's readable
'  - Copies the VERIFIED file offsite (the one actually created)
'  - Prunes local backups older than RETAIN_DAYS
'  - Fails loudly so a broken backup is noticed
'==================================================================
Option Explicit

Private Const RETAIN_DAYS As Integer = 14

Public Sub SQLDailyBackUp(PathToBackUp As String)
    On Error GoTo ERH

    If GetConfigInfo("Misc", "Skip BackUp", "No", "Functions.INI") = "ENABLED" Then Exit Sub

    Dim Db       As ADODB.Connection
    Dim Sql      As String
    Dim Stamp    As String
    Dim FileName As String
    Dim Folder   As String
    Dim LocalBak As String

    See_Msg "Backing up main database", 5, ""

    ' Date-stamped name so each day is retained (no overwrite of history).
    Folder = PathToBackUp
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    Stamp = Format(Now, "yyyymmdd_hhnnss")
    FileName = "MSaleData_" & Stamp & ".bak"
    LocalBak = Folder & FileName

    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db

    ' Native backup to LOCAL disk, fresh file, with page checksums.
    Sql = "BACKUP DATABASE [MSaleData] TO DISK = '" & LocalBak & "' " & _
          "WITH INIT, CHECKSUM, NAME = 'MSaleData Daily'"
    Db.Execute Sql

    ' Prove the backup is restorable before we trust / copy it.
    Sql = "RESTORE VERIFYONLY FROM DISK = '" & LocalBak & "' WITH CHECKSUM"
    Db.Execute Sql

    SetNothing Db

    ' Copy the verified file offsite (the file we actually created).
    ' A failure here shouldn't lose the good local backup, so trap it.
    If Len(RemoteBackup_Path) > 0 Then
        On Error Resume Next
        FileCopy LocalBak, RemoteBackup_Path & FileName
        If Err.Number <> 0 Then
            Trans_Log "Daily Sql Backup: offsite copy failed " & Err.Description
            Err.Clear
        End If
        On Error GoTo ERH
    End If

    ' Retention: remove local backups older than RETAIN_DAYS.
    PurgeOldBackups Folder, "MSaleData_*.bak", RETAIN_DAYS

    See_Msg "Daily backup complete", 5, ""
    Exit Sub

ERH:
    Trans_Log "Daily Sql Backup FAILED " & Err.Description
    See_Msg "Daily SQL Backup FAILED", 5, Err.Description   ' fail loudly
    If Not Db Is Nothing Then SetNothing Db
End Sub


'==================================================================
' Delete files matching <pattern> in <folder> older than keepDays.
' Names are collected first (Dir is stateful and can't be walked
' while Kill is deleting).
'==================================================================
Private Sub PurgeOldBackups(ByVal Folder As String, ByVal Pattern As String, _
                            ByVal keepDays As Integer)
    On Error Resume Next
    Dim names As New Collection
    Dim f     As String
    Dim full  As String
    Dim i     As Integer

    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"

    f = Dir(Folder & Pattern)
    Do While Len(f) > 0
        names.Add f
        f = Dir
    Loop

    For i = 1 To names.Count
        full = Folder & names.Item(i)
        If FileDateTime(full) < (Now - keepDays) Then Kill full
    Next i
End Sub
