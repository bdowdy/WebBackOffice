Imports System.Data.SqlClient

Public Module SqlBulkBackup

    ''' <summary>
    ''' Bulk-copies the result of a SELECT from a source SQL Server into a
    ''' destination table. Clears the destination first (clean snapshot),
    ''' maps columns by NAME, preserves identity values, and runs the
    ''' clear + load in one transaction so a failure leaves the prior
    ''' backup intact. Returns "Success" or "ERROR: ...".
    ''' </summary>
    Public Function BulkCopyData(ByVal selectSql As String,
                                 ByVal tableName As String,
                                 ByVal desSqlDbConStr As String,
                                 ByVal srcSqlDbConStr As String,
                                 Optional ByVal clearFirst As Boolean = True,
                                 Optional ByVal keepIdentity As Boolean = True) As String
        Dim dest As String = Qualify(tableName)

        Try
            Using d_Db As New SqlConnection(desSqlDbConStr)      ' destination
                d_Db.Open()

                Using tran As SqlTransaction = d_Db.BeginTransaction()
                    Try
                        ' Clean snapshot: clear the destination table first.
                        If clearFirst Then
                            Using clr As New SqlCommand("DELETE FROM " & dest, d_Db, tran)
                                clr.CommandTimeout = 0
                                clr.ExecuteNonQuery()
                            End Using
                        End If

                        Dim options As SqlBulkCopyOptions = SqlBulkCopyOptions.TableLock
                        If keepIdentity Then options = options Or SqlBulkCopyOptions.KeepIdentity

                        Using s_Db As New SqlConnection(srcSqlDbConStr)   ' source
                            s_Db.Open()

                            Using cmd As New SqlCommand(selectSql, s_Db)
                                cmd.CommandTimeout = 0
                                Using reader As SqlDataReader = cmd.ExecuteReader()
                                    Using bulk As New SqlBulkCopy(d_Db, options, tran)
                                        bulk.DestinationTableName = dest
                                        bulk.BulkCopyTimeout = 0      ' no timeout on large tables
                                        bulk.BatchSize = 5000

                                        ' Map by NAME so differing column order can't scramble data.
                                        For i As Integer = 0 To reader.FieldCount - 1
                                            Dim col As String = reader.GetName(i)
                                            bulk.ColumnMappings.Add(col, col)
                                        Next

                                        bulk.WriteToServer(reader)
                                    End Using
                                End Using
                            End Using
                        End Using

                        tran.Commit()
                    Catch
                        tran.Rollback()       ' leave the existing backup untouched
                        Throw
                    End Try
                End Using
            End Using

            Return "Success"

        Catch ex As Exception
            LogTxt("BulkCopyData [" & tableName & "] " & ex.Message)
            Return "ERROR: " & ex.Message
        End Try
    End Function

    ' Bracket-quote a table name unless it's already qualified ([schema].[name]).
    Private Function Qualify(ByVal tableName As String) As String
        If tableName.Contains("[") Then Return tableName
        Return "[" & tableName.Replace("]", "]]") & "]"
    End Function

End Module
