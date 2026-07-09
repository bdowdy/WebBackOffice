Attribute VB_Name = "TerminalUpdates"
' ============================================================
'  Terminal update check - parameterized, server-clock based.
'
'  Replaces:
'    Sql = "Select * from [Terminal Update] where [Name] = '" & _
'           MyRegisterName & "' and [date] <= '" & _
'           Format(Now, "MM/DD/YYYY HH:MM:SS") & "'"
'
'  Improvements:
'    - Parameterized (no string-built SQL / no injection / no
'      date-format ambiguity).
'    - Uses the SQL Server clock (GETDATE()) instead of each
'      terminal's Now, so scheduled updates fire consistently.
'    - Captures a single cutoff so the rows we CLEAR are exactly
'      the rows we ACTED on (no race with newly-inserted rows).
'
'  Typical logon use:
'    Call ApplyTerminalUpdatesOnLogon(MyRegisterName)
'
'  Requires an ADODB reference (Microsoft ActiveX Data Objects).
' ============================================================
Option Explicit

' Returns True if this terminal has pending updates due now or earlier.
' Cutoff is filled with the server time used, so the same value can be
' passed to ClearTerminalUpdates for a consistent delete.
Public Function TerminalUpdatePending(ByVal RegisterName As String, _
                                      ByRef Cutoff As Date) As Boolean
    On Error GoTo ERH
    Dim Db  As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim RS  As ADODB.Recordset

    TerminalUpdatePending = False
    Cutoff = 0

    If ConnectADODB("MSaleData", Db) = False Then Exit Function

    ' Pin the cutoff to the server clock once.
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = Db
    Cmd.CommandType = adCmdText
    Cmd.CommandText = "SELECT GETDATE()"
    Set RS = Cmd.Execute
    If Not RS.EOF Then Cutoff = RS(0)
    RS.Close: Set RS = Nothing

    ' Count pending rows for this terminal at/under the cutoff.
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = Db
    Cmd.CommandType = adCmdText
    Cmd.CommandText = _
        "SELECT COUNT(*) FROM [Terminal Update] " & _
        "WHERE [Name] = ? AND [date] <= ?"
    Cmd.Parameters.Append Cmd.CreateParameter("@Name", adVarWChar, adParamInput, 100, RegisterName)
    Cmd.Parameters.Append Cmd.CreateParameter("@Cut", adDBTimeStamp, adParamInput, , Cutoff)
    Set RS = Cmd.Execute
    If Not RS.EOF Then TerminalUpdatePending = (Val(RS(0) & "") > 0)

    RS.Close: Set RS = Nothing
    Db.Close: Set Db = Nothing
    Exit Function
ERH:
    On Error Resume Next
    If Not RS Is Nothing Then RS.Close
    If Not Db Is Nothing Then Db.Close
    Set RS = Nothing: Set Db = Nothing
    ' Fail closed: no update rather than a bad partial load.
    TerminalUpdatePending = False
End Function

' Clears the rows we just acted on (same terminal, up to the same
' cutoff). Future-dated rows are left in place. Call ONLY after the
' data load succeeded.
Public Sub ClearTerminalUpdates(ByVal RegisterName As String, ByVal Cutoff As Date)
    On Error GoTo ERH
    Dim Db  As ADODB.Connection
    Dim Cmd As ADODB.Command

    If ConnectADODB("MSaleData", Db) = False Then Exit Sub

    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = Db
    Cmd.CommandType = adCmdText
    Cmd.CommandText = _
        "DELETE FROM [Terminal Update] " & _
        "WHERE [Name] = ? AND [date] <= ?"
    Cmd.Parameters.Append Cmd.CreateParameter("@Name", adVarWChar, adParamInput, 100, RegisterName)
    Cmd.Parameters.Append Cmd.CreateParameter("@Cut", adDBTimeStamp, adParamInput, , Cutoff)
    Cmd.Execute

    Db.Close: Set Db = Nothing
    Exit Sub
ERH:
    On Error Resume Next
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing
End Sub

' Logon entry point: if updates are pending, load the data, then clear
' the rows we handled. Marking-done happens ONLY after a successful load,
' so a failed load re-fires on the next logon instead of being lost.
Public Sub ApplyTerminalUpdatesOnLogon(ByVal RegisterName As String)
    Dim Cutoff As Date
    If TerminalUpdatePending(RegisterName, Cutoff) Then
        '--- your existing reload routine goes here, e.g.:
        '    LoadMenuData
        '    LoadEmployees
        '    LoadPrintDrivers
        '--------------------------------------------------
        ClearTerminalUpdates RegisterName, Cutoff
    End If
End Sub
