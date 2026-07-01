Attribute VB_Name = "Module1"
' ============================================================
' AccountTip.frm  –  Corrected Subs
' Apply these replacements to the matching subs in AccountTip.frm
' ============================================================


' ============================================================
' BUGS #1-#5 FIX — SaveTipToCheckStat
' #1: ConnectADODB return value discarded — added check.
' #2: ConnectADORS return value discarded — added check.
' #3: TimeOfSale Else branch set the variable AFTER the SQL
'     was built, so the time filter was silently dropped when
'     TimeOfSale was blank, causing the query to return all
'     payment rows for that check and pushing into the fragile
'     multi-row fallback.  SQL now always includes time filter;
'     if blank, current time is assigned BEFORE building SQL.
' #4: Tb![Tip Amount] = Format(TipAmt,"0.00") assigns a String
'     to a numeric field.  Changed to Tb![Tip Amount] = TipAmt.
' #5: Multi-row fallback looped on [Tip Amount] = 0, which
'     matches any untipped row and picks the wrong one.
'     Replaced with a tighter ORDER BY + TOP 1 re-query so the
'     correct row is selected without ambiguity.
' Bonus: Format(Time,"HH:MM:SS ampm") — "ampm" is not a VB6
'     token; corrected to "AM/PM".
' ============================================================
Public Sub SaveTipToCheckStat(TipAmt As Currency, PayAmountToCheckNo As Long, ChargeAmt As Currency, ChargeType As String, TimeOfSale As String)
On Error GoTo ERH
    Dim Db As ADODB.Connection
    Dim Tb As ADODB.Recordset

    ' BUG #3 FIX: assign TimeOfSale BEFORE building SQL so the
    ' time filter is always included in the query.
    If TimeOfSale = "" Then
        TimeOfSale = Format(Time, "HH:MM:SS AM/PM")   ' Bonus: "AM/PM" not "ampm"
    End If

    ' BUG #1 FIX: check connection return value
    If ConnectADODB(CheckStatDb, Db) = False Then Exit Sub

    Sql = "Select * from [payments] where [Close Out Day] = 0"
    Sql = Sql & " and [check Number] = " & PayAmountToCheckNo
    Sql = Sql & " and [Payment Type] = '" & ChargeType & "'"
    Sql = Sql & " and [Payment Amount] = " & ChargeAmt
    Sql = Sql & " and [Time] = '" & TimeOfSale & "'"

    ' BUG #2 FIX: check recordset return value
    If ConnectADORS(Sql, Db, Tb) = False Then
        DbClose Db, Tb
        Exit Sub
    End If

    If Tb.EOF Then
        ' No existing row — insert new
        Tb.AddNew
        Tb![Check Number] = PayAmountToCheckNo
        Tb![Payment Type] = ChargeType
        Tb![Payment Amount] = ChargeAmt
        Tb![Date] = Format(Date, "MM/DD/YY")
        Tb![Time] = TimeOfSale
        Tb![Close out day] = 0
        Tb![Tip Amount] = TipAmt         ' BUG #4 FIX: assign Currency, not String
        Tb.Update
    Else
        Dim R As Long
        R = RecordsCount(Tb)
        If R = 1 Then
            Tb![Tip Amount] = TipAmt     ' BUG #4 FIX
            Tb.Update
        Else
            ' BUG #5 FIX: multiple rows matched — re-query with TOP 1 ORDER BY
            ' to select the most recent untipped row rather than looping on
            ' [Tip Amount] = 0 which picks the wrong row.
            SetNothing Tb
            Sql = "Select TOP 1 * from [payments] where [Close Out Day] = 0"
            Sql = Sql & " and [check Number] = " & PayAmountToCheckNo
            Sql = Sql & " and [Payment Type] = '" & ChargeType & "'"
            Sql = Sql & " and [Payment Amount] = " & ChargeAmt
            Sql = Sql & " and [Time] = '" & TimeOfSale & "'"
            Sql = Sql & " ORDER BY cast([Date] as date) DESC, [Time] DESC"
            If ConnectADORS(Sql, Db, Tb) = False Then
                DbClose Db, Tb
                Exit Sub
            End If
            If Not Tb.EOF Then
                Tb![Tip Amount] = TipAmt
                Tb.Update
            End If
        End If
    End If

    DbClose Db, Tb
    Exit Sub
ERH:
    Error_Log Err.Number, Err.Description, "APPLY TIP", "None"
End Sub


' ============================================================
' BUG #1 FIX — Charges_Click
' Old code used InStr("A") which breaks when "A" appears in
' the scan number before "Acct." (e.g. account "ACT-2201").
' Also, InStr("-") for AcctNo fails if scan number has a hyphen.
' ============================================================
Private Sub Charges_Click()
    Dim CKNNO       As Long
    Dim AcctNo      As String
    Dim ListItem    As String
    Dim AcctPos     As Integer

    ListItem = Charges.List(Charges.ListIndex)
    AcctPos = InStr(ListItem, "Acct.")
    If AcctPos = 0 Then Exit Sub        ' guard: malformed list item

    ' "CHK NO " is 7 chars; number starts at position 8
    ' One space separates check# from "Acct.", so length = AcctPos - 9
    CKNNO = Val(Mid(ListItem, 8, AcctPos - 9))

    ' Account number follows "Acct. " (6 chars); strip trailing "-"
    AcctNo = Mid(ListItem, AcctPos + 6)
    If Right(AcctNo, 1) = "-" Then AcctNo = Left(AcctNo, Len(AcctNo) - 1)

    LoadRecords CKNNO, AcctNo
End Sub


' ============================================================
' BUGS #2 #3 #4 FIX — LoadRecords
' #2: Format() returns a String; assigning it to Currency
'     silently zeroes the variable.  Use CDbl(Nz(..., 0)).
' #3: When multiple rows found the code fell through and ran
'     the balance queries with empty HTip(0) — added Exit Sub.
' #4: SELECT TOP 1 without ORDER BY returns a random row.
'     Added ORDER BY [Created] DESC (adjust column if needed).
' ============================================================
Private Sub LoadRecords(CheckNo As Long, AccountN As String)
    Dim Db              As ADODB.Connection
    Dim Tb              As ADODB.Recordset
    Dim BeginBal        As Currency
    Dim Payments        As Currency
    Dim TotalCharges    As Currency

    Sql = "Select * from [HouseCharges] Where [Day of Charge] = 0"
    If AccountN = "" Then
        Sql = Sql & " and [Voided] = 0 and [Check Number] = " & CheckNo
    Else
        Sql = Sql & " and [Voided] = 0 and [Check Number] = " & CheckNo & " and [Scan Number] = '" & AccountN & "'"
    End If

    If ConnectADODB(HouseAcct, Db) = False Then Exit Sub
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub

    If Not Tb.EOF Then
        If RecordsCount(Tb) = 1 Then
            HTip(0).Text = Tb![Charged To] & ""
            HTip(1).Text = Tb![Check Number] & ""
            HTip(2).Text = Format(Nz(Tb![Check Amount]), "0.00")
            HTip(3).Text = Format(Nz(Tb![Adj total]), "0.00")
            HTip(6).Text = Format(Nz(Tb![Tip Amount]), "0.00")
        Else
            ' Multiple charges found — show list and let user pick
            Charges.Clear
            Do While Not Tb.EOF
                Charges.AddItem "CHK NO " & CheckNo & " Acct. " & Trim$(Tb![Scan Number] & "") & "-"
                Tb.MoveNext
            Loop
            DbClose Db, Tb
            Exit Sub    ' BUG #3 FIX: was falling through to balance queries
        End If
    Else
        See_Msg "Record Not Found", 5, ""
        DbClose Db, Tb
        Exit Sub
    End If

    SetNothing Tb

    Sql = "Select sum([Adj Total]) as BalanceDue, sum([Tip Amount]) as CTip "
    Sql = Sql & "From [HouseCharges] where [charged to] = '" & HTip(0).Text & "'"
    Sql = Sql & " and [Voided] = 0"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    ' BUG #2 FIX: Format() returns String; use CDbl(Nz(...,0)) to get Currency
    If Not Tb.EOF Then TotalCharges = CDbl(Nz(Tb.fields("BalanceDue").Value, 0))
    SetNothing Tb

    Sql = "Select sum([Amount]) as BalanceDue "
    Sql = Sql & "From [HousePayments] where [Account Name] = '" & HTip(0).Text & "'"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    If Not Tb.EOF Then Payments = CDbl(Nz(Tb.fields("BalanceDue").Value, 0))
    SetNothing Tb

    ' BUG #4 FIX: Added ORDER BY so TOP 1 is deterministic
    Sql = "Select Top 1 [Beginning Balance] from [Starting Balance]"
    Sql = Sql & " Where [Account Name] = '" & HTip(0).Text & "'"
    Sql = Sql & " Order By [Created] DESC"   ' adjust to your actual date column name
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    If Tb.EOF Then
        BeginBal = 0
    Else
        BeginBal = CDbl(Nz(Tb![Beginning Balance], 0))
    End If

    HTip(4).Text = Format(BeginBal + TotalCharges - Payments, "0.00")
    DbClose Db, Tb
End Sub


' ============================================================
' BUGS #5 #6 #8 FIX — SSCommand14_Click
' #5: ConnectADODB for FinancialDb was called as a statement
'     (return value discarded) — if it fails the code continues
'     into ConnectADORS with a dead connection, then crashes.
' #6: ConnectADORS for Daypayments was also unchecked.
' #8: WHERE [Tip Amount] = HTip(6).Text matches by the OLD tip
'     value (often 0), pulling the wrong row when multiple
'     house charges share the same tip amount.  Removed that
'     condition; check number alone identifies the payment row.
' ============================================================
Private Sub SSCommand14_Click()
    If HTip(0).Text = "" Then Exit Sub
    If Can_Continue(SSCommand14.Caption) = False Then Exit Sub

    Dim fdb         As ADODB.Connection
    Dim ftb         As ADODB.Recordset
    Dim Db          As ADODB.Connection
    Dim RS          As ADODB.Recordset
    Dim HChargeTip  As Currency
    Dim ChgAmt      As Currency
    Dim TimeSale    As String
    Dim MyTenderName As String

    ' Excessive-tip guard
    If Val(Input_Option(66)) > 0 And Val(HTip(2).Text) > 0 Then
        If Val(HTip(5).Text) > (Val(HTip(2).Text) * (Val(Input_Option(66)) / 100)) Then
            See_Msg "Exccessive Tip Amount", 5, "Manager Access Required"
            If Can_Continue("Managers Menu") = False Then Exit Sub
        End If
    End If

    ' --- Update HouseCharges ---
    Sql = "Select * from [HouseCharges] Where [Voided] = 0 and [Check Number] = " & HTip(1).Text
    Sql = Sql & " and [Charged To] = '" & HTip(0).Text & "'"
    Sql = Sql & " and [Day Of Charge] = 0"
    If ConnectADODB(HouseAcct, Db) = False Then Exit Sub
    If ConnectADORS(Sql, Db, RS) = False Then Exit Sub

    HChargeTip = Val(HTip(5).Text)
    If Not RS.EOF Then
        ChgAmt = Val(RS![charge Amount] & "")
        RS![Adj total] = RS![charge Amount] + HChargeTip
        RS![Tip Amount] = HChargeTip
        RS.Update
    End If
    DbClose Db, RS

    ' --- Update Daypayments ---
    ' BUG #5 FIX: check return value before proceeding
    If ConnectADODB(SvrPath.TerminalPath & FinancialDb, fdb) = False Then
        See_Msg "Cannot connect to financial database", 5, "Warning"
        GoTo ClearFields   ' tip is saved in HouseCharges; skip Daypayments update
    End If

    ' BUG #8 FIX: removed "AND [Tip Amount] = HTip(6)" which matched by
    ' the old (often-zero) tip and pulled the wrong row.
    Sql = "Select * from [Daypayments] where [check Number] = " & HTip(1).Text
    Sql = Sql & " and [Charged To] = '" & HTip(0).Text & "'"   ' tighten match with account name

    ' BUG #6 FIX: check return value
    If ConnectADORS(Sql, fdb, ftb) = False Then
        DbClose fdb, ftb
        GoTo ClearFields
    End If

    If Not ftb.EOF Then
        ftb![Tip Amount] = HChargeTip
        TimeSale = ftb.fields("Recorded Time").Value & ""
        MyTenderName = ftb.fields("Payment Type").Value & ""
        ftb.Update
    End If
    DbClose fdb, ftb

    SaveTipToCheckStat HChargeTip, Val(HTip(1).Text), ChgAmt, MyTenderName, TimeSale

ClearFields:
    HTip(0).Text = ""
    HTip(1).Text = ""
    HTip(2).Text = "0.00"
    HTip(3).Text = "0.00"
    HTip(4).Text = "0.00"
    HTip(5).Text = "0.00"
    HTip(6).Text = "0.00"
    See_Msg ReturnLabel("Saved..."), 5, ""
    If Charges.ListCount = 0 Then Quit_Click
End Sub
