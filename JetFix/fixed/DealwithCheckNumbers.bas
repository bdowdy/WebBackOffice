Attribute VB_Name = "DealwithCheckNumbers"

' =========================================================
' Small helper that executes a parameterized scalar query:
' returns the first column of the first row or Null.
' =========================================================
Public Sub StartServerAudit(NewServerName As String, OkToStart As Boolean, WorkAreaName As String, StartBank As Currency)
On Error GoTo ErrorHandler

    Dim NextNoAv       As Long
    Dim getnewCknum    As ADODB.Connection
    Dim Records        As ADODB.Recordset
    Dim ShiftNo        As Long
    Dim GroupAuditNo   As Long
    Dim TryCount       As Integer
    Dim SqlDate        As String

    If NewServerName = "" Then
        See_Msg "Please Select a Server", 5, "Cannot Start Audit"
        Exit Sub
    End If

    If ClosingProgess Then
        See_Msg "Day Close Out In Progress", 0, "Cannot Start Audit Until Close Out Complete."
        Exit Sub
    End If

    If OkToStart = False Then
        If Check_Option(82) = 1 Then
         If Punched_In(NewServerName) = False Then
            See_Msg "Punch In Required To Start Audit", 5, "Please Punch In"
            Exit Sub
         End If
        End If
    End If

    If Check_Option(122) = 1 Then
        StartBank = Val(ShowNumberPad("Enter Starting Bank", False, 0, 0))
    End If

    GetSelected = ""
    LoggedOnAs(0).MyRegisterGroup = GetRevenueCenter("", WorkAreaName, "")

    ConnectADODB SvrPath.TerminalPath & "MSaleData", getnewCknum

    '===========================
    ' Group audit setup
    '===========================
    If TerminalIsInGroupAudit(LoggedOnAs(0).MyRegisterGroup) Then

        Sql = "SELECT * FROM [GroupAudit] WHERE [AuditName] = '" & SqlSafe(WorkAreaName) & "'"
        ConnectADORS Sql, getnewCknum, Records

        If Not Records.EOF Then
            GroupAuditNo = Val(Records.fields("CheckSequence").Value & "")
            If GroupAuditNo < 80000 Then
                Trans_Log "reset group audit number"
            Else
                See_Msg "Audit Already Open for This Group", 5, WorkAreaName
                DbClose getnewCknum, Records
                Exit Sub
            End If
        Else
            Records.AddNew
            Records.fields("AuditName").Value = WorkAreaName
            GroupAuditNo = Next_Check_No(WorkAreaName)
            Records.fields("CheckSequence").Value = GroupAuditNo
            Records.fields("GroupName").Value = LoggedOnAs(0).MyRegisterGroup
            Records.fields("InUse").Value = 0
            Records.Update
        End If

        SetNothing Records
        NewServerName = WorkAreaName
    End If

    '===========================
    ' Work area selection
    '===========================
    If Check_Option(89) = 1 And SvrPath.GroupAudit = False Then
        If WorkAreaName = "Driver" Then
            GV.WorkArea = WorkAreaName
        Else
            GetSelected = GetConfigInfo("Dining Room", "Default", "", "FUNCTIONS.INI")
            If GetSelected > "" Then
                GV.WorkArea = GetSelected
            Else
                Load_What = "Work Area"
                Find_What = "Select Room"
                RequestListForm.Show 1
                If GetSelected = "" Then
                    DbClose getnewCknum, Records
                    Exit Sub
                End If
                GV.WorkArea = GetSelected
            End If
        End If
    Else
        GV.WorkArea = WorkAreaName
    End If

    SqlDate = Format$(Date, "mm/dd/yyyy")

    '===========================
    ' Already open audit?
    '===========================
    Sql = "SELECT TOP 1 * FROM [Check Numbers] " & _
          "WHERE [Servers Name] = '" & SqlSafe(NewServerName) & "' " & _
          "AND [Audit Status] = 0 "

        If SvrPath.GroupAudit Then
            Sql = Sql & "AND CONVERT(varchar(10), [Start Date], 101) = '" & SqlDate & "' "
        End If

    ConnectADORS Sql, getnewCknum, Records

    If Not Records.EOF Then
        If Val(Records![Audit Status] & "") = 0 Then
            If SvrPath.GroupAudit Then
                See_Msg NewServerName & " has an open audit", 5, ""
            Else
                See_Msg NewServerName & " has an open audit", 5, "Cannot start audit"
            End If
            DbClose getnewCknum, Records
            Exit Sub
        End If
    End If

    SetNothing Records

    '===========================
    ' Find unique check number
    '===========================
    For TryCount = 1 To 5

        If SvrPath.GroupAudit Then
            NextNoAv = GroupAuditNo
        Else
            NextNoAv = Next_Check_No(NewServerName)
        End If
        If TryCount > 1 Then NextNoAv = NextNoAv + (TryCount * 1000)
        If NextNoAv = 0 Then
            See_Msg "No Check Number assigned", 5, "Cannot Start Audit"
            DbClose getnewCknum, Records
            Exit Sub
        End If

        If Not CheckNoExistsInCheckNumbers(NextNoAv, getnewCknum) Then
            If Not CheckNoExistsInCheckDetail(NextNoAv) Then

                Sql = "SELECT * FROM [Check Numbers] WHERE 1 = 0"
                ConnectADORS Sql, getnewCknum, Records

                Records.AddNew
                Records.fields("Servers Name").Value = NewServerName
                Records.fields("Check No").Value = NextNoAv
                ShiftNo = NextNoAv
                Records.fields("Time").Value = Format$(Time$, "hh:mm AM/PM")
                Records.fields("Start Date").Value = Format$(Date, "mm/dd/yy")
                Records.fields("Audit Status").Value = 0
                Records.fields("In Use").Value = 0

                If Check_Option(0) = 1 Or (Check_Option(34) = 1 And Input_Option(10) = "Open/Close") Then
                    Records.fields("Register Name").Value = MyRegisterName
                    Records.fields("Drawer No").Value = Drwauditno
                Else
                    Records.fields("Register Name").Value = ""
                    Records.fields("Drawer No").Value = 0
                End If

                If TerminalIsInGroupAudit(LoggedOnAs(0).MyRegisterGroup) Then
                    Records.fields("Register Group").Value = LoggedOnAs(0).MyRegisterGroup
                Else
                    Records.fields("Register Group").Value = SvrPath.TerminalGroupName
                End If

                If SvrPath.GroupAudit Then
                    Records.fields("Work Area").Value = WorkAreaName
                Else
                    Records.fields("Work Area").Value = GV.WorkArea & ""
                End If

                Records.fields("Starting Bank").Value = StartBank
                Records.fields("Deposits Paid In").Value = 0
                Records.Update

                DbClose getnewCknum, Records

                See_Msg "Starting Audit For " & NewServerName, 5, ""
                Trans_Log "Starting Audit for " & NewServerName & " Check Number " & NextNoAv
                RecordShiftNumber NewServerName, ShiftNo

                If Check_Option(0) = 1 Or _
                   (Check_Option(34) = 1 And Input_Option(10) = "Open/Close") Or _
                   (Check_Option(213) = 1 And NewServerName = MyRegisterName) Then

                    WriteRegSales NewServerName, StartBank, NextNoAv
                    Print_Audit_Start NewServerName, StartBank
                    Drawer_Pop Drwauditno
                Else
                    If Check_Option(207) = 0 And Check_Option(213) = 0 Then
                        If Check_Option(34) = 0 Then
                            If DrawerPort(0) <> "None" Then
                                If See_Msg("Assign Cash Drawer", 4, "Assign To Drawer ?") = 6 Then
                                    StartBarDrawer NewServerName
                                End If
                            End If
                        End If
                    End If
                End If

                Exit Sub
            End If
        End If
    Next TryCount

    DbClose getnewCknum, Records
    See_Msg "Unable to get unique check number", 5, "Cannot Start Audit"
    Exit Sub

ErrorHandler:
    ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Start Server Audit ")
    Select Case ErrorRespose
        Case "Abort"
            Exit Sub
        Case "Resume"
            Resume
        Case "Resume Next"
            Resume Next
    End Select
End Sub
'//
Public Function Next_Check_No(ByVal GetLastCheckFrom As String) As Long
    On Error GoTo ErrorHandler

    Dim conn     As ADODB.Connection
    Dim cmd      As ADODB.Command
    Dim Sql      As String
    Dim lastNo   As Variant
    Dim nextBlockSeed As Long
    Dim retry    As Integer

    Next_Check_No = 0

    ' ------------------------------------------------------------
    ' 1) Try to get the last open [Check Number] for this server
    '    from [Check Detail], then return +1 if found
    ' ------------------------------------------------------------
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, conn

    Sql = "SELECT TOP 1 [Check Number] " & _
          "FROM [Check Detail] " & _
          "WHERE [Close out Day] = 0 " & _
          "  AND [Servers Name] = ? " & _
          "  AND [Check Number] > 999 " & _
          "  AND [Cashier Name] <> 'Closed'"

    ' Optional scoping (your existing options)
    If Check_Option(0) = 1 Then
        If Check_Option(17) = 1 Then
            Sql = Sql & " AND [Order Type] = '" & SvrPath.TerminalGroupName & "'"
        Else
            Sql = Sql & " AND [Register Name] = '" & MyRegisterName & "'"
        End If
    End If

    Sql = Sql & " ORDER BY [Check Number] DESC"

    lastNo = NCExecScalar(conn, Sql, GetLastCheckFrom)
    If Not IsNull(lastNo) Then
        Next_Check_No = CLng(lastNo) + 1
        GoTo cleanexit
    End If

    ' ------------------------------------------------------------
    ' 2) If “drive-thru number” mode is enabled, try that
    ' ------------------------------------------------------------
    If Check_Option(17) = 1 Then
        Next_Check_No = GetDriveThruNo(0)
        If Next_Check_No > 999 Then
            GoTo cleanexit
        End If
        ' If = 999, fall through and look in audit start table.
    End If

    ' ------------------------------------------------------------
    ' 3) Look for a per-server starting number in MSaleData.[check numbers]
    ' ------------------------------------------------------------
    DbClose conn, Nothing
    ConnectADODB SvrPath.TerminalPath & "MSaleData", conn

    Sql = "SELECT [Check No] " & _
          "FROM [check numbers] " & _
          "WHERE [Servers Name] = ? " & _
          "  AND [Audit Status] = 0"

    lastNo = NCExecScalar(conn, Sql, GetLastCheckFrom)
    If Not IsNull(lastNo) Then
        Next_Check_No = CLng(lastNo)
        GoTo cleanexit
    End If

    ' ------------------------------------------------------------
    ' 4) No per-server row. Pick the next block based on the
    '    last [check numbers] row, split by GroupAudit.
    ' ------------------------------------------------------------
    If SvrPath.GroupAudit = False Then
        ' “Retail” range (< 80000): bump by 1000s
        Sql = "SELECT TOP 1 [Check No] " & _
              "FROM [check numbers] " & _
              "WHERE [Check No] < 80000 " & _
              "ORDER BY [Check No] DESC"

        lastNo = NCExecScalar(conn, Sql)
        If Not IsNull(lastNo) Then
            nextBlockSeed = ((CLng(lastNo) \ 1000) + 1) * 1000
            Next_Check_No = nextBlockSeed
        Else
            ' Nothing at all in table: pick a sensible seed
            Next_Check_No = 1000
            If Check_Option(17) = 0 Then
              Next_Check_No = CheckNumberIsAvailable(GetTerminalNumber(Next_Check_No))
            End If
            '// if no numbers then use terminal number for QSR other wise use check 1000
        End If

        ' Optional: if you want to ensure “no duplicates among open checks”
        ' and align to your existing allocator, uncomment the next line:
        

    Else
        ' “Group audit” range (>= 80000): bump by 10000s
        Sql = "SELECT TOP 1 [Check No] " & _
              "FROM [check numbers] " & _
              "WHERE [Check No] >= 80000 " & _
              "ORDER BY [Check No] DESC"

        lastNo = NCExecScalar(conn, Sql)
        If Not IsNull(lastNo) Then
            nextBlockSeed = ((CLng(lastNo) \ 10000) + 1) * 10000
            Next_Check_No = nextBlockSeed
        Else
            ' Nothing yet, start the group range at 80000
            Next_Check_No = 80000
        End If
    End If

cleanexit:
    DbClose conn, Nothing
    Exit Function

ErrorHandler:
    ' Centralized error policy consistent with your code
    ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Next_Check_No")
    Select Case ErrorRespose
        Case "Abort"
            ' leave Next_Check_No as 0
        Case "Resume"
            Resume
        Case "Resume Next"
            Resume Next
        Case Else
            ' best-effort close
            DbClose conn, Nothing
    End Select
End Function

' =========================================================
' Small helper that executes a parameterized scalar query:
' returns the first column of the first row or Null.
' =========================================================
Private Function NCExecScalar(ByVal conn As ADODB.Connection, _
                            ByVal Sql As String, _
                            Optional ByVal param1 As Variant) As Variant
    On Error GoTo oops

    Dim cmd As ADODB.Command
    Dim RS  As ADODB.Recordset

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = Sql
        .CommandType = adCmdText
        If Not IsMissing(param1) Then
            .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, 50, param1)
        End If
    End With

    Set RS = cmd.Execute
    If Not (RS.BOF And RS.EOF) Then
        NCExecScalar = RS.fields(0).Value
    Else
        NCExecScalar = Null
    End If

    If RS.State <> 0 Then RS.Close
    Set RS = Nothing
    Set cmd = Nothing
    Exit Function

oops:
    NCExecScalar = Null
    On Error Resume Next
    If Not RS Is Nothing Then If RS.State <> 0 Then RS.Close
    Set RS = Nothing
    Set cmd = Nothing
End Function
'//
Private Function SqlSafe(ByVal Txt As String) As String
    SqlSafe = Replace(Txt, "'", "''")
End Function
Private Function CheckNoExistsInCheckNumbers(ByVal CheckNo As Long, ByVal Db As ADODB.Connection) As Boolean
On Error GoTo ErrHandler

    Dim RS As ADODB.Recordset
    Dim Sql As String

    Sql = "SELECT TOP 1 [Check No] " & _
          "FROM [Check Numbers] " & _
          "WHERE [Check No] = " & CheckNo & " " & _
          "AND [Audit Status] = 0"

    ConnectADORS Sql, Db, RS
    CheckNoExistsInCheckNumbers = Not RS.EOF

Cleanup:
    DbClose Nothing, RS
    Exit Function

ErrHandler:
    CheckNoExistsInCheckNumbers = True
    GoTo Cleanup
End Function
Private Function CheckNoExistsInCheckDetail(ByVal CheckNo As Long) As Boolean
On Error GoTo ErrHandler
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim Sql As String

    Sql = "SELECT TOP 1 [Check Number] " & _
          "FROM [Check Detail] " & _
          "WHERE [Check Number] = " & CheckNo & " " & _
          "AND [Close Out Day] = 0"
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    ConnectADORS Sql, Db, RS
    CheckNoExistsInCheckDetail = Not RS.EOF

Cleanup:
    DbClose Db, RS
    Exit Function

ErrHandler:
    CheckNoExistsInCheckDetail = True
    GoTo Cleanup
End Function
'// end new section

Public Function CallInNumberIsGood(MyCheckNo As Long) As Long
On Error GoTo ErrorHandler
Dim NumberInUse As Long
NumberInUse = MyCheckNo
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Rst     As String
Dim Tb      As ADODB.Recordset
GetAnother:
If GRS.IsPhoneOrderCheckNoNotInUse(MyCheckNo, AllOrders.Item(1).CallerPhone) = False Then
  MyCheckNo = GRS.NextCheckNumber("PhoneOrders")  ' "PhoneOrders"
  Sql = "Update [Debit] set [Check Number] = " & MyCheckNo & " Where [close out day] = 0 and [Check Number] = " & NumberInUse
  Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & DebitDb))
  DbClose Nothing, Tb
  GoTo GetAnother
Else
  CallInNumberIsGood = MyCheckNo
End If
Set GRS = Nothing
Exit Function
ErrorHandler:
CallInNumberIsGood = MyCheckNo
End Function
Public Function CheckNumberIsAvailable(ByVal NewNumber As Long) As Long
    On Error GoTo ErrorHandler

    Dim Tb   As ADODB.Recordset
    Dim GRS  As GeneralLib.Utilities
    Dim Sql  As String
    Dim conn As String
    Dim Block As Long
    Dim MaxNo As Long
    Dim retry As Integer

    Set GRS = New GeneralLib.Utilities
    conn = ReturnAdoConnection(SvrPath.TerminalPath & CheckStatDb)

    ' Default: assume the passed number is fine
    CheckNumberIsAvailable = NewNumber

    ' Work in 1000-number blocks
    Block = NewNumber \ 1000

    ' 1) Is there ANY open check already in this same 1000-block?
    Sql = "SELECT TOP 1 1 AS HasRow " & _
          "FROM [Check Detail] " & _
          "WHERE [Close out Day] = 0 " & _
          "  AND ([Check Number] \ 1000) = " & CStr(Block)

    Set Tb = GRS.ReturnRecordset(Sql, conn, "Check Detail")
    If Tb Is Nothing Then GoTo cleanexit

    ' If no rows, the requested number is available
    If Tb.BOF And Tb.EOF Then
        GoTo cleanexit
    End If

    ' 2) Block is taken; choose the next 1000-block based on the current MAX
    Tb.Close
    Set Tb = Nothing

    Sql = "SELECT MAX([Check Number]) AS MaxNo " & _
          "FROM [Check Detail] " & _
          "WHERE [Close out Day] = 0"

    Set Tb = GRS.ReturnRecordset(Sql, conn, "Check Detail")
    If Not (Tb.BOF And Tb.EOF) Then
        If Not IsNull(Tb.fields("MaxNo").Value) Then
            MaxNo = CLng(Tb.fields("MaxNo").Value)
            ' IMPORTANT: ((max \ 1000) + 1) * 1000   (note the parentheses)
            CheckNumberIsAvailable = ((MaxNo \ 1000) + 1) * 1000
        Else
            ' No open checks but prior EXISTS found a row? Defensive fallback
            CheckNumberIsAvailable = 0
        End If
    Else
        ' No rows at all
        CheckNumberIsAvailable = 0
    End If

cleanexit:
    On Error Resume Next
    If Not Tb Is Nothing Then
        If Tb.State <> 0 Then Tb.Close
        Set Tb = Nothing
    End If
    Set GRS = Nothing
    Exit Function

ErrorHandler:
    retry = retry + 1
    If retry < 5 Then
        Sleep 500
        Resume
    End If
    CheckNumberIsAvailable = 0
    Resume cleanexit
End Function
Public Function CheckNumberInUse(ByVal CheckNumber As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Dim cmd  As ADODB.Command
    Dim RS   As ADODB.Recordset
    Dim retry As Integer

    CheckNumberInUse = False

    ' Open the CheckStat connection
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, conn

    ' Use TOP 1 1 to short-circuit as soon as a match exists
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = "SELECT TOP 1 1 FROM [Open Checks] WHERE [Check Number] = ?"
        .Parameters.Append .CreateParameter("pCheckNo", adInteger, adParamInput, , CheckNumber)
    End With

    Set RS = cmd.Execute
    If Not (RS.BOF And RS.EOF) Then
        CheckNumberInUse = True
    End If

cleanexit:
    On Error Resume Next
    If Not RS Is Nothing Then If RS.State <> 0 Then RS.Close
    Set RS = Nothing
    Set cmd = Nothing
    DbClose conn, Nothing
    Exit Function

ErrorHandler:
    retry = retry + 1
    If retry < 5 Then
        Sleep 500
        Resume
    End If
    ' On persistent error, report "not in use" conservatively or flip to True if you prefer fail-closed
    CheckNumberInUse = False
    Resume cleanexit
End Function

Public Function IsCallInCheckNumberInUse(CallInNumber As Long) As Boolean
On Error GoTo ErrorHandler
Dim Tb  As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Rst As String
 IsCallInCheckNumberInUse = False
 Sql = "Select [Check Number] from [Open Checks] "
 Sql = Sql & " where [Check Number] = " & CallInNumber
 Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
 If Not Tb.BOF Then Tb.MoveFirst
 If Not Tb.EOF Then
    IsCallInCheckNumberInUse = True
    SetNothing Tb
  End If
  DbClose Nothing, Tb
 Set GRS = Nothing
Exit Function
ErrorHandler:
retry = retry + 1
If retry < 5 Then Sleep 500: Resume
CallInNumber = 0
End Function
Public Function CallInNumber() As Long
On Error GoTo ErrorHandler
If AbortSale = True Then CallInNumber = Check_no: Exit Function
CallInNumber = 0
Dim Tb  As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Rst As String
GetAnother:
  CallInNumber = GRS.NextCheckNumber("PhoneOrders")
  Sql = "Select [Check Number] from [Open Checks] "
  Sql = Sql & " where [Check Number] = " & CallInNumber
  Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
  If Not Tb.BOF Then Tb.MoveFirst
  If Not Tb.EOF Then
    SetNothing Tb
    Trans_Log "Call In Number In Use " & CallInNumber
    GoTo GetAnother
  End If
  Tb.Close
  Sql = "Select [PhoneOrderNumber] from [PrePaid] "
  Sql = Sql & " where [PhoneOrderNumber] = " & CallInNumber
  Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "PrePaid")
  If Not Tb.BOF Then Tb.MoveFirst
  If Not Tb.EOF Then
    SetNothing Tb
    Trans_Log "Call In Payment Number In Use " & CallInNumber
    GoTo GetAnother
  End If '// order comming in from on line will have payment listed in the table local phone orders need to check as well
  '// make sure that a payment is not picked up on close.
  '// make sure that a payment from future phone 0rder is not open in the system
  DbClose Nothing, Tb
 Set GRS = Nothing
Exit Function
ErrorHandler:
     retry = retry + 1
     If retry < 5 Then Sleep 500: Resume
     CallInNumber = 0
End Function
Public Function Advance_Check(StartAudit As Boolean) As Long
On Error GoTo ErrorHandler
'// quick Service
If AbortSale = True Then Advance_Check = Check_no: Exit Function
'Added for call in order numbers only resets at end of day
TAgain:
If Check_Option(140) = 0 Then Sold_To = "" 'need to reset if not printed

Dim GetTopCheckNo   As Long
Dim Rst             As String
Dim Tb              As ADODB.Recordset
If Check_Option(81) = 0 Then
    If Check_Option(17) = 1 Then
      If OrderIs = GetOrderType(DTSale) Then
         OrderConfBoard StartNewDTOrder, 0
         OrderConfBoard UpdateDTOrderTotal, 0
      End If
    End If
End If
WhosLoggedOn
getreg:

    Advance_Check = 0:  ModIdNo = 0: Selection_Number = 0
    If GetActiveCashier = "" Then GetActiveCashier = LoggedOnAs(0).EMPName
    If Check_Option(17) = 1 And StartAudit = False Then
      Advance_Check = GetDriveThruNo(0)
      If CheckNumberInUse(Advance_Check) Then GoTo getreg
      If Advance_Check = 0 Then GoTo DoRegCheckNo
    Else
      Dim GRS As GeneralLib.Utilities
      Set GRS = New GeneralLib.Utilities

DoRegCheckNo:
      Sql = "Select top 1 [check number] from [Check Detail] "
      Sql = Sql & " Where [close out day] = 0 "
       If Check_Option(0) = 1 Then
        If Check_Option(60) = 0 Then
         Sql = Sql & " and [Servers Name] = '" & GetActiveCashier & "'"
        End If
       End If
       If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" Then
         Sql = Sql & " and [Servers Name] = '" & GetActiveCashier & "'"
         Sql = Sql & " and [Cashier Name] = '" & GetActiveCashier & "'"
         Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
         'must use double condition make sure that cashier check no not using
         'servers check No.
       Else
         Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
       End If
       Sql = Sql & " order by [check number] desc"
       Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(CheckStatDb), "Check Detail")
         If Not Tb Is Nothing Then
           If Not Tb.BOF Then Tb.MoveFirst
         End If
         If Not Tb.EOF Then
           Advance_Check = Val(Tb![Check Number]) + 1
         Else
           If StartAudit = False And Check_Option(34) = 1 Then
              Advance_Check = Assign_Check_No(GetActiveCashier, True)
              If Advance_Check = 0 Then Advance_Check = 1
           Else
            Tb.Close
             If Check_Option(60) = 1 Then
              Sql = "Select * from [check numbers] "
              Sql = Sql & " Where [Register Name] = '" & MyRegisterName & "'"
             Else
              Sql = "Select * from  [check numbers] "
              Sql = Sql & " Where [Register Name] = '" & MyRegisterName & "'"
              Sql = Sql & " and [Servers Name] = '" & GetActiveCashier & "'"
             End If
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection("MSaleData"), "Check Numbers")
             If Not Tb Is Nothing Then
               If Not Tb.BOF Then Tb.MoveFirst
             End If
             If Not Tb.EOF Then
               Advance_Check = Val(Tb![Check No]) + 1
             Else
               Advance_Check = Assign_Check_No(GetActiveCashier, True)
             End If
           End If
         End If
         DbClose Nothing, Tb
         Set GRS = Nothing
    End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Advance Check")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function Assign_Check_No(ByVal AssignServer As String, ByVal ResetSelect As Boolean) As Long
On Error GoTo ErrorHandler
retry = 0
tryagain:
If Trim$(AssignServer) = "" Then
 If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" Then
   AssignServer = GetActiveCashier
 End If
End If

If Trim$(AssignServer) = "" Then
  AssignServer = LoggedOnAs(0).EMPName
  Trans_Log "No Server Name Found Assigning Check To " & AssignServer & " " & LoggedOnAs(0).EMPName
End If
If ResetSelect Then Selection_Number = 0: ModIdNo = 0

Dim GRS         As GeneralLib.Utilities
Dim Rst         As String
Dim getnewCknum As ADODB.Connection
Dim Records     As ADODB.Recordset
Set GRS = New GeneralLib.Utilities

'Need flag for different options.
'Make sure items will not be connected together
If Check_Status = MyOrderStatus(SplitCheck) Then
 'MsgBox "Split Check"
 '// make sure that check number will not increment when spliting a check or split paying a check
Else
 If SvrPath.GroupAudit Then
    Assign_Check_No = GRS.GroupAuditCheckNumber(AssignServer) 'LoggedOnAs(0).EmpAssignedTo
    Set GRS = Nothing
    Exit Function
 End If
End If
If Input_Option(10) = "Open/Close" Then
   Sql = "Select top 1 * from [Audits] where [Name] = '" & AssignServer & "'"
   Set Records = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & FinancialDb), "Audits")
   If Not Records.BOF Then Records.MoveFirst
   If Not Records.EOF Then
      If Records.fields("Register Name").Value <> MyRegisterName Then
         Records.Close
         Sql = "Select Top 1 * from [Audits] where [Register Name] = '" & MyRegisterName & "'"
         Set Records = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & FinancialDb), "Audits")
         If Not Records.BOF Then Records.MoveFirst
         If Not Records.EOF Then
           AssignServer = Records.fields("Name").Value & ""
           Trans_Log "Switched Cashier To " & AssignServer & " Order Is " & OrderIs
         End If
      End If
   End If
   Records.Close
End If

Assign_Check_No = 0: Order_Qty = 0
If UCase(AssignServer) = "TRAINING MODE" Then Assign_Check_No = 1: Exit Function
Sql = "Select top 1 * from [Check Numbers] where [Servers Name] = '" & AssignServer & "'"
If Input_Option(10) = "Open/Close" Then
  '// is server audit is close by cashier audit is open will recycle check numbers
  '// Louis customer 11/29/2022
Else
  Sql = Sql & " and [Audit Status] = 0 "
End If
Set Records = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & "MSaleData"), "Check Numbers")
If Not Records.BOF Then Records.MoveFirst
If Records.EOF Then
            If Input_Option(10) = "Open/Close" Then
             DbClose Nothing, Records
             Sql = "Select top 1 * from [Audits] where [Name] = '" & AssignServer & "'"
             Set Records = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & FinancialDb), "Audits")
             If Not Records.EOF Then
                Assign_Check_No = Records.fields("Check No").Value + 1
             Else
                Assign_Check_No = 0
             End If
            End If
            DbClose Nothing, Records
            Set GRS = Nothing
            Exit Function
Else
            
            'sql = "Select top 1 [check number] from [check detail] where [close out day] = 0 "
            'sql = sql & " and [Servers Name] = '" & AssignServer & "'"
            'sql = sql & " and [Check Number] > 999"
            ' sql = sql & " and TRY_CONVERT(int, [Check Number] / 1000) = " & Int(Val(Records.fields("Check No").Value & "") / 1000)
            'sql = sql & " order by [check number] desc"
            Dim MyBaseNo As Long
            Dim BlockStart As Long
            Dim BlockEnd As Long
            MyBaseNo = Val(Records.fields("Check No").Value & "")
            BlockStart = (MyBaseNo \ 1000) * 1000
            BlockEnd = BlockStart + 999

            Sql = "SELECT TOP 1 [Check Number] " & _
                    "FROM [Check Detail] " & _
                    "WHERE [Close Out Day] = 0 " & _
                    "AND [Servers Name] = '" & Replace(AssignServer, "'", "''") & "' " & _
                    "AND [Check Number] BETWEEN " & BlockStart & " AND " & BlockEnd & " " & _
                    "ORDER BY [Check Number] DESC"
            Dim NextCheck As ADODB.Recordset
            Set NextCheck = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CheckStatDb), "Check Numbers")
            If Not NextCheck.BOF Then NextCheck.MoveFirst
                If NextCheck.EOF Then
                    Assign_Check_No = MyBaseNo
                    If Assign_Check_No = 0 Then
                       See_Msg "No Check Number Assigned for This Server", 0, "Audit Will Be Removed Please Restart Audit"
                       Sql = "Delete from [Check Numbers] where [Servers Name] = '" & AssignServer & "'"
                       getnewCknum.Execute Sql
                       Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & "MSaleData"))
                       DbClose Nothing, NextCheck
                       DbClose Nothing, Records
                       Set GRS = Nothing
                       Exit Function
                    End If
                Else
                    Dim NextCheckNoInUse As Long
                    NextCheckNoInUse = Val(NextCheck(0) & "") + 1
                    If IsCheckNumberInRange(MyBaseNo, NextCheckNoInUse) Then
                       Assign_Check_No = NextCheckNoInUse
                    Else
                       See_Msg "Your Check Number is Out of range", 5, ""
                       Assign_Check_No = NextCheckNoInUse
                    End If
                End If
            DbClose Nothing, NextCheck
            DbClose Nothing, Records
End If
Set GRS = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Assign check No " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case "Resume"
       Resume
     Case "Resume Next"
       Resume Next
   End Select
End Function
Public Function IsCheckNumberInRange(ByVal BaseNumber As Long, nextCheckNo As Long) As Boolean
If nextCheckNo >= BaseNumber And nextCheckNo < BaseNumber + 1000 Then
    IsCheckNumberInRange = True
Else
    IsCheckNumberInRange = False
    Trans_Log "Check Number Is Out of Range " & nextCheckNo & " Base No " & BaseNumber
End If
End Function
Public Function GetDriveThruNo(SetDTNo As Long) As Long
On Error GoTo ErrorHandler
GetDriveThruNo = 0
Dim MYName  As String
Dim Rst     As String
Dim Tb      As ADODB.Recordset
Dim GRS     As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities

If GetConfigInfo("DriveThru", "Numbers by Register", "No", "Functions.INI") = "ENABLED" Then
  MYName = "DriveThru" & ReturnKVSTerminalId
Else
  MYName = "DriveThru"
End If
GetDriveThruNo = GRS.NextCheckNumber(MYName)

If SetDTNo > 0 Then
 If SetDTNo < GetDriveThruNo And GetDriveThruNo > 999 Then
   Exit Function
 End If '// this will make sure the check number cannot get rolled back
 If GetDriveThruNo < 1000 Or GetDriveThruNo <> SetDTNo Then
   GetDriveThruNo = SetDTNo
   Sql = "Delete [SharedCheckNumbers] Where [Name] = '" & MYName & "'"
   Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection("MSaleData"))
   Sql = "Insert Into [SharedCheckNumbers] ([Name],[Check Number]) Values ('" & MYName & "'," & GetDriveThruNo & ")"
   Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection("MSaleData"))
   DbClose Nothing, Tb
 End If
End If
Set GRS = Nothing
Exit Function
ErrorHandler:
retry = retry + 1
If retry < 2 Then Sleep 500: Resume
Error_Log Err.Number, Err.Description, "Get Drive Check", "None"
End Function
Public Function CommonCheckNo(SharedName As String) As Long
On Error GoTo ErrorHandler
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
CommonCheckNo = GL.NextCheckNumber(SharedName)
'Send 'ShareName & " Check Number Starts at=xxx
Set GL = Nothing
Exit Function
ErrorHandler:
 CommonCheckNo = 0
End Function


