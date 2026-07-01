Imports System.Data.OleDb

' Backs up SQL Server table(s) into an Access database as a clean snapshot.
' Each run DROPs and recreates the target table, so there are no duplicates
' and the schema always follows the source.
'
' Requires the matching ACE/Jet OLE DB provider installed and the right
' bitness for your process (x86 app -> 32-bit provider).
Public Module SqlAccessBackup

    ''' <summary>
    ''' Snapshot one SQL Server table into the Access database.
    ''' Returns "Success" or "ERROR: ...".
    ''' </summary>
    Public Function BackupSqlTableToAccess(ByVal tableName As String,
                                           ByVal serverName As String,
                                           ByVal sqlDatabase As String,
                                           ByVal accessPath As String,
                                           ByVal accessDb As String,
                                           Optional ByVal useTrustedConnection As Boolean = True,
                                           Optional ByVal sqlUser As String = "",
                                           Optional ByVal sqlPassword As String = "") As String
        Try
            ' Authentication portion of the ODBC source clause.
            Dim auth As String =
                If(useTrustedConnection,
                   "Trusted_Connection=yes;",
                   "Uid=" & sqlUser & ";Pwd=" & sqlPassword & ";")

            ' The IN '' [ODBC; ...] clause tells Jet/ACE to read from SQL Server.
            Dim odbc As String =
                "[ODBC;Driver={SQL Server};Server=" & serverName &
                ";Database=" & sqlDatabase & ";" & auth & "]"

            ' Bracket + escape the identifier so spaces / keywords are safe.
            Dim t As String = "[" & tableName.Replace("]", "]]") & "]"

            Using cn As New OleDbConnection(ReturnAdoConnection(accessPath, accessDb))
                cn.Open()

                ' 1) Drop any prior copy so the backup is a clean snapshot.
                '    Ignore the error if the table isn't there yet.
                Try
                    Using dropCmd As New OleDbCommand("DROP TABLE " & t, cn)
                        dropCmd.ExecuteNonQuery()
                    End Using
                Catch
                    ' Table did not exist - nothing to drop.
                End Try

                ' 2) Recreate the table and copy all rows in one statement.
                Dim sql As String =
                    "SELECT * INTO " & t & " FROM " & t & " IN '' " & odbc & ";"

                Using copyCmd As New OleDbCommand(sql, cn)
                    copyCmd.CommandTimeout = 0   ' don't time out on large tables
                    copyCmd.ExecuteNonQuery()
                End Using
            End Using

            Return "Success"

        Catch ex As Exception
            LogTxt("BackupSqlTableToAccess [" & tableName & "] " & ex.Message)
            Return "ERROR: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Snapshot several tables. Returns a per-table summary string.
    ''' </summary>
    Public Function BackupSqlTablesToAccess(ByVal tableNames As String(),
                                            ByVal serverName As String,
                                            ByVal sqlDatabase As String,
                                            ByVal accessPath As String,
                                            ByVal accessDb As String,
                                            Optional ByVal useTrustedConnection As Boolean = True,
                                            Optional ByVal sqlUser As String = "",
                                            Optional ByVal sqlPassword As String = "") As String
        Dim results As New System.Text.StringBuilder()
        For Each tbl As String In tableNames
            Dim r As String = BackupSqlTableToAccess(tbl, serverName, sqlDatabase,
                                                     accessPath, accessDb,
                                                     useTrustedConnection, sqlUser, sqlPassword)
            results.AppendLine(tbl & ": " & r)
        Next
        Return results.ToString()
    End Function

End Module
