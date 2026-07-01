Attribute VB_Name = "functions"
Option Explicit
Public Type SalesTaxAmt
       CashValue       As Currency
       ChkValue        As Currency
       RegularSalesTax As Currency
       SpecialSalesTax As Currency
       LiquorTax       As Currency
End Type
Public Type SeatName
       SeatName As String
       AddName  As String
       
End Type

Public Type TerminalId
   TerminalNo As Long
   TerminalGuid As String
   TerminalName As String
   MachineName As String
   WindowsUser As String
   IpAddress As String
End Type

Public Type TipOutPosition
       JobCode As String
       EMPName As String
End Type
Public Type depositpaidto
      IsPaid As Boolean
      Payment As Currency
      Media As String
End Type
Public Type GiftLocalResponse
      Name As String
      CardValue As Double
      VoucherCode As String
      Message As String
      Status As String
      Code As String
      PublicToken As String
      VoucherRedemptionId As String
      isGiftType As Integer
      isPromoType As Integer
      isRedeemable As Integer
End Type
Public Enum MTGiftOptions
       FTGetBalance = 1
       FTaddfunds = 2
       FTnewcard = 3
       FTverifycvv = 5
       FTgethistory = 5
       FTUpdateCustomer = 6
       FTactivecard = 7
       FTsubtractfunds = 8
End Enum
Public Type ReturnTimeValues
    RecordId As String
    ShiftDate As String
    Results   As String
End Type
 Public Function ReturnMTAction(DataIn As MTGiftOptions) As String
        Select Case DataIn
            Case Is = MTGiftOptions.FTactivecard
                ReturnMTAction = "activate-card"
            Case MTGiftOptions.FTaddfunds
                ReturnMTAction = "add-funds"
            Case MTGiftOptions.FTGetBalance
                ReturnMTAction = "get-balance"
            Case MTGiftOptions.FTgethistory
                ReturnMTAction = "get-history"
            Case MTGiftOptions.FTnewcard
                ReturnMTAction = "new-card"
            Case MTGiftOptions.FTUpdateCustomer
                ReturnMTAction = "update-customer"
            Case MTGiftOptions.FTverifycvv
                ReturnMTAction = "verify-cvv"
            Case MTGiftOptions.FTsubtractfunds
                ReturnMTAction = "subtract-funds"
            Case Else
                ReturnMTAction = "get-balance"
        End Select
End Function
'*****************************************************************************************
' Function Name: Is7shiftSetup
' Description: Checks if the 7shift setup is successful by verifying the setup key.
' Returns:
'   - Boolean: True if the setup is successful, False otherwise.
'*****************************************************************************************
Public Function ReadGiftLocalResponse(apiResponseJson As String) As GiftLocalResponse
Dim D As Object
Set D = ParseKV(apiResponseJson)
If D.Exists("success") And D("success") = "1" Then
    With ReadGiftLocalResponse
        .CardValue = CDbl(D("presentValue"))
        .VoucherCode = D("voucherCode")
        .Message = "Success"
        .Name = D("recipient")
        .Status = D("status")
        .PublicToken = D("publicToken")
        .VoucherRedemptionId = D("voucherRedemptionID")
        .isGiftType = D("isGiftType")
        .isPromoType = D("isPromoType")
        .isRedeemable = D("isRedeemable")
    End With
Else
    ReadGiftLocalResponse.Message = D("messages")
    ReadGiftLocalResponse.Code = D("error")
End If
End Function

' Returns a Scripting.Dictionary of key/value pairs
' Example:
'   Set d = ParseKV("success=1|voucherCode=107-100-718-544|presentValue=50")
'   MsgBox d("voucherCode")
Public Function ParseKV(ByVal s As String) As Object
    Dim D As Object
    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = 1 ' vbTextCompare (case-insensitive keys)

    Dim parts() As String
    Dim i As Long
    Dim P As String
    Dim eqPos As Long
    Dim K As String, v As String

    If LenB(s) = 0 Then
        Set ParseKV = D
        Exit Function
    End If

    parts = Split(s, "|")

    For i = LBound(parts) To UBound(parts)
        P = parts(i)

        ' Find first "=" only (value could contain "=" later, if escaped)
        eqPos = InStr(1, P, "=", vbBinaryCompare)
        If eqPos > 0 Then
            K = Left$(P, eqPos - 1)
            v = Mid$(P, eqPos + 1)

            K = Trim$(K)
            v = UnescapeKV(v)

            If LenB(K) > 0 Then
                D(K) = v   ' overwrite if key repeats
            End If
        End If
    Next i

    Set ParseKV = D
End Function

' Reverse of the EscapeKV used on the .NET side:
'   %25 -> %
'   %7C -> |
'   %3D -> =
Private Function UnescapeKV(ByVal s As String) As String
    If LenB(s) = 0 Then
        UnescapeKV = ""
        Exit Function
    End If

    s = Replace$(s, "%7C", "|")
    s = Replace$(s, "%3D", "=")
    s = Replace$(s, "%25", "%")

    UnescapeKV = s
End Function

Public Function Is7shiftSetup() As Boolean
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim SSF As SevenShift
    
    ' Instantiate SevenShift object
    Set SSF = New SevenShift
    
    ' Check if the setup key retrieved from SevenShift object is "success"
    If SSF.GetSetupKey = "success" Then
        ' If setup key is "success", return True
        Is7shiftSetup = True
    Else
        ' If setup key is not "success", return False
        Is7shiftSetup = False
    End If
    
    ' Clean up - Release object reference
    Set SSF = Nothing
    
    ' Exit function
    Exit Function

ErrorHandler:
    ' In case of an error, return False
    Is7shiftSetup = False
End Function
Public Function DiscountPercent() As Double
    ' Handle division by zero
    If Gross_Sales <= 0 Then
        DiscountPercent = 0
    Else
        DiscountPercent = (Abs(Chk_Sales_data(4)) / Gross_Sales) * 100
    End If  '// used to not cash discount the check if the amount of percent is 85 percent or greater
End Function

Public Function IsDataCapInterface(interface As String) As Boolean
    If interface = "DATACAP" Then
        IsDataCapInterface = True
    Else
        IsDataCapInterface = False
    End If
End Function
Public Function IsPAXInterface(interface As String) As Boolean
    If interface = "PAX" Then
        IsPAXInterface = True
    Else
        IsPAXInterface = False
    End If
End Function
Public Function IsDejavooInterface(interface As String) As Boolean
    If interface = "DEJAVOO" Then
        IsDejavooInterface = True
    Else
        IsDejavooInterface = False
    End If
End Function
Public Sub ForceKillDotNetApp(ByVal ExeName As String)
    Shell "taskkill /F /IM " & ExeName, vbHide
End Sub

Public Function IsProgramRunning(ByVal ExeName As String) As Boolean
    Dim objWMI As Object
    Dim colProcesses As Object
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMI.ExecQuery( _
        "Select * From Win32_Process Where Name = '" & ExeName & "'")
    
    IsProgramRunning = (colProcesses.count > 0)
End Function
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
'ReturnAdoConnection (DatabaseName)

' ---------------------------------------------------------------------


'==================================================================
' Back up every user table in the database, one CSV per table.
'==================================================================
Public Sub BackupDebitSql(Optional ByVal backupRoot As String = "")
    On Error GoTo ErrorHandler

    Dim cn          As ADODB.Connection
    Dim rsTables    As ADODB.Recordset
    Dim destFolder  As String
    Dim TableName   As String
    Dim tableType   As String
    Dim fileCount   As Long

    If Len(backupRoot) = 0 Then backupRoot = DEFAULT_BACKUP_ROOT

    ' Create  <root>\Backup_YYYYMMDD_HHNNSS
    destFolder = backupRoot & "\Backup_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolder destFolder

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open = ReturnAdoConnection(DebitDb)

    ' List base tables (skip views and system tables).
    Set rsTables = cn.OpenSchema(adSchemaTables)
    Do While Not rsTables.EOF
        tableType = rsTables.fields("TABLE_TYPE").Value & ""
        If tableType = "TABLE" Then
            TableName = rsTables.fields("TABLE_NAME").Value & ""
            ExportTable cn, TableName, destFolder & "\" & SafeFileName(TableName) & ".csv"
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
Public Sub BackupDebitSqlTable(ByVal TableName As String, _
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

    ExportTable cn, TableName, destFolder & "\" & SafeFileName(TableName) & ".csv"

    cn.Close: Set cn = Nothing
    MsgBox "Backed up [" & TableName & "] to:" & vbCrLf & destFolder, _
           vbInformation, "DebitSql Backup"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    MsgBox "Backup of [" & TableName & "] failed: " & Err.Description, _
           vbCritical, "DebitSql Backup"
End Sub


'==================================================================
' Write one table's rows to a CSV file (header row + data rows).
'==================================================================
Private Sub ExportTable(ByRef cn As ADODB.Connection, _
                        ByVal TableName As String, _
                        ByVal filePath As String)
    Dim RS   As ADODB.Recordset
    Dim ff   As Integer
    Dim line As String
    Dim i    As Long

    Set RS = New ADODB.Recordset
    ' Brackets allow table names containing spaces.
    RS.Open "SELECT * FROM [" & TableName & "]", cn, adOpenForwardOnly, adLockReadOnly

    ff = FreeFile
    Open filePath For Output As #ff

    ' Header row
    line = ""
    For i = 0 To RS.fields.count - 1
        If i > 0 Then line = line & ","
        line = line & CsvField(RS.fields(i).Name)
    Next i
    Print #ff, line

    ' Data rows
    Do While Not RS.EOF
        line = ""
        For i = 0 To RS.fields.count - 1
            If i > 0 Then line = line & ","
            line = line & CsvField(RS.fields(i).Value)
        Next i
        Print #ff, line
        RS.MoveNext
    Loop

    Close #ff
    RS.Close: Set RS = Nothing
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
Private Function SafeFileName(ByVal Name As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        Name = Replace(Name, ch, "_")
    Next ch
    SafeFileName = Name
End Function


Public Function VoidApprovals(ProcessorRespone As String) As String
   Dim ResponseUpper As String
    ResponseUpper = UCase(ProcessorRespone)
    If InStr(ResponseUpper, "REVERSED") > 0 Or _
       InStr(ResponseUpper, "ALREADY REVERSED") > 0 Or _
       InStr(ResponseUpper, "ALREADY VOIDED") > 0 Or _
       ResponseUpper = "APPROVED OFFLINE" Then
       VoidApprovals = "APPROVAL"
       Emv_Back.r_Status = "APPROVAL"
    Else
       VoidApprovals = ResponseUpper
    End If
End Function
Public Function IsApproved(ProcessorResponse As String) As Boolean
    Dim ResponseUpper As String
   
    ResponseUpper = UCase(ProcessorResponse)

    If InStr(ResponseUpper, "AUTH/TKT") > 0 Or _
       InStr(ResponseUpper, "APPROVAL") > 0 Or _
       InStr(ResponseUpper, "APPROVED") > 0 Or _
       ResponseUpper = "APPROVED OFFLINE" Then
       
       Emv_Back.r_Status = "APPROVAL"
       IsApproved = True
    Else
        IsApproved = False
    End If
End Function
Public Function DropModifierFunctions(ByVal DepartmentName As String, ByVal FunctionName As String) As String
   Select Case UCase(FunctionName)
      Case Is = "CLICK PRICE", "NONE", "CHANGE ITEM", "PREFIX MODIFIER", "", "PRESELECT"
       DropModifierFunctions = DepartmentName
      Case Left$(FunctionName, 5) = "UPICK"
       DropModifierFunctions = DepartmentName
      Case Is = Left$(FunctionName, 5) = "CLICK"
       DropModifierFunctions = DepartmentName
      Case Is = "-TARE"
       DropModifierFunctions = DepartmentName & "-TARE"
      Case Else
       DropModifierFunctions = FunctionName
   End Select
End Function
Public Function ReturnAccountNumber(ByVal DataStrIn As String) As String
 On Error GoTo ERH
 '// parse house and giftcard numbers for local accounts
 
 ReturnAccountNumber = ""
 If Trim$(DataStrIn) = "" Then Exit Function
 '// Input_Option(85) can be a list "Filter 2","Drop1Parse2","Parse 2","Track 2","Track1or2","Both"
 
 If Input_Option(85) = "Filter 2" Then
   If InStr(DataStrIn, "=") > 0 Then
      ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "=") - 2)
   Else
      ReturnAccountNumber = Mid(DataStrIn, InStr(DataStrIn, ";") + 1, Len(DataStrIn) - (InStr(DataStrIn, ";") + 1))
   End If
 Else
   Select Case Input_Option(85)
    Case Is = "Drop1Parse2"
      DataStrIn = Mid(DataStrIn, InStr(DataStrIn, ";"), InStr(DataStrIn, "="))
      ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "=") - 2)
    Case Is = "Parse 2"
      ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "=") - 2)
    Case Is = "Track 1"
      ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "?") - 2)
    Case Is = "Track 2"
      ReturnAccountNumber = Mid(DataStrIn, InStr(DataStrIn, ";") + 1, Len(DataStrIn) - (InStr(DataStrIn, ";") + 1))
    Case Is = "Track1or2"
        If Left$(DataStrIn, 1) = "%" Then  'Track 1
           ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "?") - 2)
        Else
           ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "=") - 2)
        End If
    Case Is = "Both"
       If InStr(DataStrIn, ";") >= 0 Then
          ReturnAccountNumber = Mid(DataStrIn, InStr(DataStrIn, ";") + 1, Len(DataStrIn) - (InStr(DataStrIn, ";") + 1))
       ElseIf Left$(DataStrIn, 1) = "%" Then
          ReturnAccountNumber = Mid(DataStrIn, 2, InStr(DataStrIn, "?") - 2)
       Else
         ReturnAccountNumber = DataStrIn
       End If
    Case Else
      ReturnAccountNumber = DataStrIn
  End Select
 End If
 If Len(ReturnAccountNumber) < 4 Then
  See_Msg "Error Reading Card", 5, DataStrIn
  ReturnAccountNumber = ""
 End If
 Exit Function
ERH:
   See_Msg "Error parsing account number", 5, ""
End Function
Public Function DepartmentTimePriceAdjustment(ItemPrice As Currency, EventPrice As Currency) As Currency
If Abs(EventPrice) > ItemPrice Then
  DepartmentTimePriceAdjustment = ItemPrice * (EventPrice / 100)
Else
  DepartmentTimePriceAdjustment = EventPrice
End If
End Function
Public Function CompleteStripe(ByVal DataIn As String) As Boolean
CompleteStripe = False
If InStr(DataIn, ";") > 0 And InStr(DataIn, "=") > 0 Then
   CompleteStripe = True
End If
'// return true value with a complete magnetic card swipe
End Function
Public Function GetBeginPayrollWeek(BeginWeek As Date) As Date
 GetBeginPayrollWeek = BeginWeek - Val(Input_Option(25))
End Function
Public Function GetStartOfWeek() As Date
Dim StartOfWeek As Date
Dim FirstDay    As String
Dim Find        As Integer
 StartOfWeek = Date
 GetStartOfWeek = Date
 For Find = 0 To 6
   FirstDay = Format(CDate(Date) - Find, "dddd")
   If FirstDay = Input_Option(6) Then
      GetStartOfWeek = CDate(StartOfWeek) - Find
      Exit Function
   End If
 Next Find
End Function

Public Function GetTodaysDate() As String
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
 
 GetTodaysDate = Format(Date, "mm/dd/yyyy")
  Sql = "Select Distinct TRY_CAST([Order Date] AS datetime) from [Closed Checks] Where [Close Out Day] = 0"
  Sql = Sql & " Group by [Order Date] Order by TRY_CAST([Order Date] AS datetime) asc"
  ConnectADODB CheckStatDb, Db
  ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
        GetTodaysDate = Format(RS(0), "MM/DD/YYYY")
    End If
DbClose Db, RS
Exit Function
ErrorHandler:
GetTodaysDate = Format(Date, "MM/DD/YYYY")
Err.Clear
End Function

Public Sub TodaysGame(WhichDay As Integer, STime As String, ETime As String, ByContestName As String)
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim DayOrWeek   As Integer

Set SMix = New SalesMix
DayOrWeek = WhichDay
'Added to accept date entry as well as numeric
Dim qty As Single
Dim Reg As Integer
Dim Inds As Integer
Dim DataStr As String
ConnectADODB CheckStatDb, Db
        
        
        Sql = "Select Distinct [menu Item],[Menu Number],[Quantity], count([Menu Item]) as QtySold, Sum([Price]) as TotalSold, Sum([Cost Amount]) as TCost From [Check Detail] Where [Close Out Day] = " & WhichDay & ""
        Sql = Sql & " and [Voided] = 0 and [Deleted] = 0"
        Sql = Sql & " and [Department Report] = 'GAMES'"
        
        If IsDate(STime) And IsDate(ETime) Then
         Sql = Sql & " and TRY_CAST([recorded time] AS datetime) between '" & CDate(STime) & "' and '" & CDate(ETime) & "'"
        End If
        Sql = Sql & " Group By [Menu Item],[Menu Number],[Quantity]"
      ConnectADORS Sql, Db, Tb
        
        Update.Show
        Update.UPdateWhere.Caption = "Reading Sales " & SvrPath.TerminalName
        Update.Refresh
       '******************
        Do While Not Tb.EOF
         'If ByContestName = Tb![Department Report] & "" Then
           With SMix
           If .Item(Tb![Menu Item] & Tb.fields("Menu Number").Value) Is Nothing Then
            .add Tb![Menu Item] & Tb.fields("Menu Number").Value
           End If
           DataStr = Tb![Menu Item] & Tb.fields("Menu Number").Value
           If Tb!Quantity > 0 And Tb!QtySold > 0 Then
             qty = Tb!Quantity * Tb!QtySold
           Else
             qty = Tb!QtySold
           End If
            .Item(DataStr).MenuItemName = Tb![Menu Item]
            .Item(DataStr).MenuCatName = Tb.fields("Menu Number").Value & ""
            .Item(DataStr).QtySold = .Item(DataStr).QtySold + qty
            .Item(DataStr).DollarsSold = Format(.Item(DataStr).DollarsSold + Val(Tb!Totalsold & ""), "0.00")
            .Item(DataStr).SoldCost = Format(Val(Tb!TCost & ""), "0.00") ' Format(.Item(DataStr).SoldCost + Val(Tb!TCost & ""), "0.00")
            .Item(DataStr).MenuDeptName = "Games" ' Tb![Department Report] & ""
            .Item(DataStr).IsDiscount = False
           End With
           If Tb.AbsolutePosition > 0 Then
            Update.Progress.ProgressBarValue = (Tb.AbsolutePosition / RecordsCount(Tb)) * 100
           End If
           Update.Refresh
         
          Tb.MoveNext
        Loop
        DbClose Db, Tb
        Call MemoryCheck(Update)
Leave:
Exit Sub
ErrorHandler:
If Err.Number = 6 Then Resume Next
ErrorRespose = Main_Error_Handler(Err.Description, Err.Number, "IND SALES MIX ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Public Function ReadCard(ByVal CardHolderName As String) As String
On Error GoTo ErrorHand
'****  adds customer name to tabs
Dim CardData  As String
Dim MyTabName As String
Dim MYName()  As String
Dim Nc        As Integer
CardData = CardHolderName
ReadCard = CardData
CardHolderName = Replace(CardHolderName, "_", " ")
CardHolderName = Replace(CardHolderName, "/", " ")
MYName = Split(Trim$(CardHolderName), " ")
If MYName(0) = MYName(1) Then
  MyTabName = MYName(0)
Else
 Nc = UBound(MYName, 1)
 If Nc > 1 Then
  MyTabName = MYName(0) & " " & MYName(1)
 Else
  MyTabName = MYName(0) & " " & MYName(1)
 End If
End If
ReadCard = Replace(MyTabName, "'", "`")
Exit Function
ErrorHand:
ReadCard = CardData
Trans_Log Err.Description & " Reading Card Name " & CardData
End Function
Public Function PingStation(ByVal RegisterName As String) As Boolean
    Dim ttpath As SetPaths
    Set ttpath = New SetPaths
     PingStation = ttpath.getRemoteStation(RegisterName)
    Set ttpath = Nothing
End Function
Public Function PaxCardEntry(ByVal EntryMode As Integer) As String
      Select Case EntryMode
          Case Is = PaxCardEntryType.Keyed
              PaxCardEntry = "Keyed"
          Case Is = PaxCardEntryType.Swiped
              PaxCardEntry = "Swiped"
          Case Is = PaxCardEntryType.ContactLess
              PaxCardEntry = "Contactless"
          Case Is = PaxCardEntryType.Scanner
              PaxCardEntry = "Scanner"
          Case Is = PaxCardEntryType.Chip
              PaxCardEntry = "Chip"
          Case Is = PaxCardEntryType.ChipBackUpSwipe
              PaxCardEntry = "Chip Fail Swipe"
       End Select
End Function '// card present options for pax credit cards
Public Function PaxCardIs(ByVal CardPresent As Integer) As String
        Select Case CardPresent
         Case Is = PaxCardPresent.CardIsNotPresent
             PaxCardIs = "Not Present"
        Case Else
             PaxCardIs = "Present"
        End Select
End Function
Public Function GetPAXCMV(ByVal ValueIn As Integer) As String
       Select Case ValueIn
           Case Is = PAXCMV.FailedCVM
             GetPAXCMV = "Failed CVM"
           Case Is = PAXCMV.EncipOffLinePin
             GetPAXCMV = "Encipherd offline pin"
           Case Is = PAXCMV.EncipOffLinePinsign
             GetPAXCMV = "Encipherd offline pin sign"
           Case Is = PAXCMV.NoCVMRequired
             GetPAXCMV = "No CVM required"
           Case Is = PAXCMV.OfflinePin
             GetPAXCMV = "Plaintext offline pin verified"
           Case Is = PAXCMV.OffLinePinSign
             GetPAXCMV = "Plaintext offline pin sign"
           Case Is = PAXCMV.OnDeviceCVM
             GetPAXCMV = "On device CVM"
           Case Is = PAXCMV.OnLinePin
             GetPAXCMV = "On line pin"
           Case Is = PAXCMV.Signature
             GetPAXCMV = "Signature"
        End Select
End Function
Public Function FindWindowPartial(ByVal TitlePart As String) As Long
On Error Resume Next
Dim hWndTmp As Long, nRet As Integer, TitleTmp As String
'alter the title to compare it case-insensitively.
TitlePart = UCase$(TitlePart)
hWndTmp = FindWindow(vbNullString, vbNullString)
Do Until hWndTmp = 0
    TitleTmp = Space$(256)
    nRet = GetWindowText(hWndTmp, TitleTmp, Len(TitleTmp))
    If nRet Then
        TitleTmp = UCase$(VBA.Left$(TitleTmp, nRet))
           If InStr(TitleTmp, TitlePart) Then
             FindWindowPartial = hWndTmp
             Exit Do
           End If
    End If
    hWndTmp = GetWindow(hWndTmp, GW_HWNDNEXT)
Loop
'TitleTmp1 = TitleTmp
'// search and make sure that no other programs remain open.
End Function
Public Function SendQBexport(ByVal BeginDate As String, ByVal EndDate As String) As String
   On Error GoTo ERH
   Dim GL As GeneralLib.Utilities
   Set GL = New GeneralLib.Utilities
     SendQBexport = GL.BuildQuickBookExport(BeginDate, EndDate)
   Set GL = Nothing
   Exit Function
ERH:
   Error_Log Err.Number, "Error exporting to Quick Books " & Err.Description, "", "Export QB"
   Err.Clear
End Function
Public Function GetModsONOFF(ByVal C_Name As String, ByVal MName As String) As String
Dim Db  As ADODB.Connection
Dim RS  As ADODB.Recordset
Dim FC  As Integer
Dim cmd As ADODB.Command
ConnectADODB SvrPath.LocalTerminalPath & "MSaleData", Db
 Sql = "select * from [Sub Categories] Where [Sub Name] = ?"
  Set cmd = New ADODB.Command
  With cmd
     .ActiveConnection = Db
     .CommandText = Sql
     .CommandType = adCmdText
     .Parameters.Append .CreateParameter("Sub Name", adVarChar, adParamInput, 50, C_Name)
  End With
    ' Execute the query
 Set RS = cmd.Execute
 If Not RS.EOF Then
   C_Name = RS.fields("Category Name").Value & ""
 Else
 SetNothing RS
 Set cmd = New ADODB.Command
  Sql = "select * from [Menu Items] Where [Menu Item Name] = ?"
  With cmd
     .ActiveConnection = Db
     .CommandText = Sql
     .CommandType = adCmdText
     .Parameters.Append .CreateParameter("Menu Item Name", adVarChar, adParamInput, 50, MName)
  End With
    ' Execute the query
 Set RS = cmd.Execute
 If Not RS.EOF Then
   C_Name = RS.fields("Department Name").Value & ""
 End If
End If
Set cmd = Nothing
DbClose Db, RS
For FC = 1 To Cats.count
  If C_Name = Cats.Item(FC).CategoryName Then
    GetModsONOFF = Cats.Item(FC).DisableMods
    Exit Function
  End If
Next FC
'// determines if modifiers need to popup for items -- use by bar station where the bar tender all ready know the mixer.
End Function

'// And for testing the routines above...
Public Function GetPreAuthValues(GetEmvDevice As String) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
EMVSUPPORTPREAUTH = 0
PAXBARTABPREAUTH = 0
INCREMENTAL = ""
If ConnectADODB(DebitDb, Db) = False Then Exit Function
Sql = "Select * from [EMVDeviceList] Where [DeviceName] = ?"
Dim cmd As New ADODB.Command
  With cmd
     .ActiveConnection = Db
     .CommandText = Sql
     .CommandType = adCmdText
     .Parameters.Append .CreateParameter("DeviceName", adVarChar, adParamInput, 50, GetEmvDevice)
  End With
    ' Execute the query
Set RS = cmd.Execute
If Not RS.EOF Then
  EMVSUPPORTPREAUTH = Val(RS.fields("SupportPreAuth").Value & "")
  RAPIDCONNECT = RS.fields("DeviceId").Value & ""
  INCREMENTAL = UCase(RS.fields("DeviceNumberPadType2").Value & "")
  If INCREMENTAL = "NOTUSED" Then INCREMENTAL = ""
  SaveSetting "Micro$ale", "DataCap", "DeviceID", RAPIDCONNECT
  SaveSetting "Micro$ale", "DataCap", "DebitIsOff", Val(RS.fields("DisableDebit").Value & "")
  GetPreAuthValues = "Success"
Else
   See_Msg "", 5, "EMV Device Not Found"
   GetPreAuthValues = "Failed"
End If
Set cmd = Nothing
DbClose Db, RS
End Function
'========================================
' Insert into [BergSaveCheck] (MSaleData)
'========================================
Public Function WriteBergSaveCheck( _
        ByVal ServerName As String, _
        ByVal CheckNo As Long, _
        ByVal TimeStr As String, _
        ByVal StartDateStr As String, _
        ByVal StatusStr As String, _
        ByVal RegisterName As String) As Boolean

    On Error GoTo Fail

    Dim cn  As ADODB.Connection
    Dim cmd As ADODB.Command

    '--- Open connection (use your existing helper if you have one) ---
    Set cn = New ADODB.Connection
    ' If you already have ReturnAdoConnection("MSaleData","MSaleData"), use that:
    ' cn.Open ReturnAdoConnection("MSaleData")
    ' Otherwise, open with your connection string:
    cn.ConnectionString = ReturnAdoConnection("MSaleData")
     
    cn.Open

    '--- Prepare parameterized insert ---
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = _
            "INSERT INTO [BergSaveCheck] " & _
            "([Servers Name],[Check No],[Time],[Start Date],[Status],[Register Name]) " & _
            "VALUES (?,?,?,?,?,?)"

        ' Clear any existing params (safety)
        If .Parameters.count > 0 Then
            Dim i As Long
            For i = .Parameters.count - 1 To 0 Step -1
                .Parameters.delete i
            Next
        End If

        ' Add parameters in the same order as the placeholders
        ' NVARCHAR(50) -> adVarWChar, size 50
        .Parameters.Append .CreateParameter("pServersName", adVarWChar, adParamInput, 50, ServerName)
        .Parameters.Append .CreateParameter("pCheckNo", adInteger, adParamInput, , CheckNo)
        .Parameters.Append .CreateParameter("pTime", adVarWChar, adParamInput, 50, TimeStr)
        .Parameters.Append .CreateParameter("pStartDate", adVarWChar, adParamInput, 50, StartDateStr)
        .Parameters.Append .CreateParameter("pStatus", adVarWChar, adParamInput, 50, StatusStr)
        .Parameters.Append .CreateParameter("pRegisterName", adVarWChar, adParamInput, 50, RegisterName)

        .Execute , , adExecuteNoRecords
    End With

    WriteBergSaveCheck = True
    GoTo Done

Fail:
    WriteBergSaveCheck = False
    ' Optional: log or show the error
    ' Debug.Print "WriteBergSaveCheck error: " & Err.Number & " - " & Err.Description

Done:
    On Error Resume Next
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not cn Is Nothing Then
        If cn.State <> adStateClosed Then cn.Close
        Set cn = Nothing
    End If
End Function

Public Function ReturnEMVProcessorInUse() As String
    Dim SRC As ADODB.Connection
    Dim rsSrc As ADODB.Recordset
    Dim Sql As String
    Dim ProcessorIs As String
    On Error GoTo Fail
    ReturnEMVProcessorInUse = ""
    Set SRC = New ADODB.Connection
    SRC.Open ReturnAdoConnection(DebitDb)

   
    Sql = "SELECT * FROM EMVDeviceList Where DeviceName = '" & MyRegisterName & "'"

    Set rsSrc = New ADODB.Recordset
    rsSrc.CursorLocation = adUseClient
    rsSrc.Open Sql, SRC, adOpenForwardOnly, adLockReadOnly
    ProcessorIs = ""
    If Not rsSrc.EOF Then
      ProcessorIs = rsSrc.fields("NumberPadType").Value & ""
    End If
    
    rsSrc.Close
    SRC.Close
    
    Set rsSrc = Nothing
    Set SRC = Nothing
    ReturnEMVProcessorInUse = ProcessorIs
    Exit Function

Fail:
    On Error Resume Next

End Function
Public Function LoadEMVDevice(GetEmvDevice As String, EmvDevice As EMVProcessLib.EmvDeviceHardWare) As String
On Error GoTo ERH
Dim Db  As ADODB.Connection
Dim RS  As ADODB.Recordset
Dim cmd As New ADODB.Command

Set EmvDevice = New EMVProcessLib.EmvDeviceHardWare
If EmvDevice.Item(GetEmvDevice) Is Nothing Then
   EmvDevice.add (GetEmvDevice)
End If
If ConnectADODB(DebitDb, Db) = False Then Exit Function
 Sql = "Select * from [EMVDeviceList] Where [DeviceName] = ?"
 Set cmd = New ADODB.Command
  With cmd
     .ActiveConnection = Db
     .CommandText = Sql
     .CommandType = adCmdText
     .Parameters.Append .CreateParameter("DeviceName", adVarChar, adParamInput, 50, GetEmvDevice)
  End With
    ' Execute the query
Set RS = cmd.Execute
If RS.EOF Then
 RS.Close
 Sql = "Select * from [EMVDeviceList] Where [DeviceName] = ?"
 Set cmd = New ADODB.Command
  With cmd
     .ActiveConnection = Db
     .CommandText = Sql
     .CommandType = adCmdText
     .Parameters.Append .CreateParameter("DeviceName", adVarChar, adParamInput, 50, "PAT")
  End With
    ' Execute the query
 Set RS = cmd.Execute
End If
If Not RS.EOF Then
  With EmvDevice
    .Item(GetEmvDevice).PinPadMacAddress = RS.fields("PinPadMacAddress").Value & ""
    .Item(GetEmvDevice).PinPadIpAddress = RS.fields("PinPadIpAddress").Value & ""
    .Item(GetEmvDevice).PinPadIpPort = RS.fields("PinPadIpPort").Value & ""
    .Item(GetEmvDevice).deviceId = RS.fields("DeviceId").Value & ""
    .Item(GetEmvDevice).DevicePort = RS.fields("DevicePort").Value & ""
    .Item(GetEmvDevice).ComPortNo = RS.fields("CommPortNo").Value & ""
    If CollectDeposit Then
     .Item(GetEmvDevice).TipLineIs = "On"
    Else
     .Item(GetEmvDevice).TipLineIs = RS.fields("TipLineIs").Value & ""
    End If
    .Item(GetEmvDevice).NumberPadDevice = RS.fields("NumberPadType").Value & ""
    .Item(GetEmvDevice).DisableDebit = Val(RS.fields("DisableDebit").Value & "")
    .Item(GetEmvDevice).SupportPreAuth = Val(RS.fields("SupportPreAuth").Value & "")
    .Item(GetEmvDevice).MaxQtyOffline = Val(RS.fields("MaxQtyOffLine").Value & "")
    .Item(GetEmvDevice).MaxAmtOffLine = Val(RS.fields("MaxAmtOffLine").Value & "")
    .Item(GetEmvDevice).Tip1 = Val(RS.fields("Tip1").Value & "")
    .Item(GetEmvDevice).Tip2 = Val(RS.fields("Tip2").Value & "")
    .Item(GetEmvDevice).Tip3 = Val(RS.fields("Tip3").Value & "")
    .Item(GetEmvDevice).Tip4 = Val(RS.fields("Tip4").Value & "")
    .Item(GetEmvDevice).PaxNoTax = Val(RS.fields("PaxNoTax").Value & "")
    .Item(GetEmvDevice).PaxTimeOut = Val(RS.fields("PaxTimeOut").Value & "")
    .Item(GetEmvDevice).PaxUseSign = Val(RS.fields("PaxUseSignature").Value & "")
    .Item(GetEmvDevice).MerchantToken = RS.fields("MerchantToken").Value & ""
    .Item(GetEmvDevice).MerchantId = RS.fields("MerchantId").Value & ""
    
   End With
   LoadEMVDevice = "Success"
Else
   See_Msg "No Device Found", 5, "Check PinPad in Credit Card Setup"
   LoadEMVDevice = "Failed"
End If
Set cmd = Nothing
DbClose Db, RS
Exit Function
ERH:
See_Msg "Device Error Failed To Load", 5, Err.Description & " " & Err.Number
Trans_Log "Error Loading EMV Device " & Err.Description & " Load EMV Devices"
LoadEMVDevice = "Failed"
End Function
Public Function ReturnDejavooToken(TPN As String) As String
On Error GoTo ERH
Dim Db  As ADODB.Connection
Dim RS  As ADODB.Recordset
Dim cmd As New ADODB.Command
ReturnDejavooToken = ""
If ConnectADODB(DebitDb, Db) = False Then Exit Function
 Sql = "Select * from [EMVDeviceList] Where [MerchantId] = ?"
 Set cmd = New ADODB.Command
  With cmd
     .ActiveConnection = Db
     .CommandText = Sql
     .CommandType = adCmdText
     .Parameters.Append .CreateParameter("MerchantId", adVarChar, adParamInput, 50, TPN)
  End With
    ' Execute the query
Set RS = cmd.Execute
If Not RS.EOF Then
  ReturnDejavooToken = RS.fields("MerchantToken").Value & ""
End If
Set cmd = Nothing
DbClose Db, RS
Exit Function
ERH:
See_Msg "Token Failed to Load", 5, "Return dejavoo token"
ReturnDejavooToken = ""
End Function

Public Function ApplyGiftCardTip(XMLString As String) As String
ApplyGiftCardTip = ""
Dim TipAmt        As Currency
Dim OriginalAmount As Currency

TipAmt = ShowNumberPad("Enter Tip Amount", False, 0, 0)
Trans_Log "Gift Tip Amount Entered " & Format(TipAmt, "0.00")
If TipAmt > 100 Then
   See_Msg "Excessive Tip Amount", 5, "Please See Manager"
   If Can_Continue("Manager Menu") = False Then Exit Function
End If
If TipAmt = 0 Then Exit Function
Dim xmls As MPxmlParse
Set xmls = New MPxmlParse
xmls.xml = XMLString
Dim MGCD As ver11DebitCards.GiftChargeData
Dim MGR As ver11DebitCards.GiftCardResponse
MGCD.AccountNumber = xmls.token("AccountNumber")
'TransTime = xmls.token("Time")
If InStr(MGCD.AccountNumber, "=") > 0 Then
  MGCD.AccountNumber = Left$(MGCD.AccountNumber, InStr(MGCD.AccountNumber, "=") - 1)
End If
MGCD.PathtoFile = SvrPath.TerminalPath
MGCD.OperatorName = LoggedOnAs(0).EMPName
MGCD.TransactionType = GiftSale
MGCD.RegisterName = MyRegisterName
MGCD.ChargeAmount = xmls.token("Purchase")
MGCD.TipAmount = TipAmt
MGCD.AuthCode = xmls.token("AuthCode")
MGCD.CheckNumber = xmls.token("InvoiceNo")
OriginalAmount = MGCD.ChargeAmount
If GiftCardPayments(MGCD, MGR) = False Then
   Exit Function
End If
ApplyFinancialTip MGCD.CheckNumber, OriginalAmount, TipAmt
Set xmls = Nothing
End Function
Public Function ApplyFinancialTip(Check As Long, PurchaseAmount As Currency, TipAmt As Currency) As String
Dim Db As ADODB.Connection
ConnectADODB FinancialDb, Db
 Sql = "Update [payments] Set [Tip Amount] = " & TipAmt & " where [Payment Amount] = " & PurchaseAmount & " and [Payment Type] <> 'Cash' and [check Number] = " & Check
Db.Execute Sql
Db.Close
ConnectADODB CheckStatDb, Db
 Sql = "Update [payments] Set [Tip Amount] = " & TipAmt & " where [Payment Amount] = " & PurchaseAmount & " and [Payment Type] <> 'Cash' and [check Number] = " & Check
Db.Execute Sql
Db.Close
End Function
Public Function GetDcapLoyaltyAuthCode(CheckNo As Long) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetDcapLoyaltyAuthCode = ""
If ConnectADODB("MSaleData", Db) = False Then
   See_Msg "Error connecting to MSaledata", 5, ""
   Exit Function
End If
  Sql = "Select * from [LoyaltyId] Where [CheckNumber] = " & CheckNo
If ConnectADORS(Sql, Db, RS) = False Then
   See_Msg "Error connecting to LoyaltyId", 5, "record set"
   Exit Function
End If
If Not RS.EOF Then
   GetDcapLoyaltyAuthCode = RS.fields("ReturnXML").Value & ""
End If
DbClose Db, RS
End Function
Public Function SaveDcapLoyaltyAuthCode(CheckNo As Long, Data As ver11DebitCards.GiftChargeData) As String
Dim Db As ADODB.Connection
ConnectADODB "MSaleData", Db
  Sql = "Update [LoyaltyId] set [ReturnXML] = '" & Data.XMLString & "',[Authcode] = '" & Data.AuthCode & "' Where [CheckNumber] = " & CheckNo
Db.Execute Sql
Db.Close
End Function
Public Function CancelDcapLoyalty(AccountNumber As String) As String
Dim Db As ADODB.Connection
ConnectADODB "MSaleData", Db
  Sql = "Delete from [LoyaltyId] Where [AccountNumber] = '" & AccountNumber & "'"
Db.Execute Sql
Db.Close
End Function
Public Function DCapLoyalty(PlanId As FrequentDinerData, WorkToDo As DCap_Loyalty) As String
 Dim MGCD   As ver11DebitCards.GiftChargeData
 Dim MGR    As ver11DebitCards.GiftCardResponse
 Dim MG     As ver11DebitCards.MercuryGiftCard
 Set MG = New ver11DebitCards.MercuryGiftCard
 If PlanId.AccountId = "" Then Exit Function
 MGCD.AccountNumber = PlanId.AccountId
 MGCD.AuthCode = PlanId.AuthCode
 MGCD.PathtoFile = SvrPath.TerminalPath
 MGCD.OperatorName = LoggedOnAs(0).EMPName
 If WorkToDo = RemovePoints Then
   MGCD.TransactionType = voidadd
   MGCD.XMLString = GetDcapLoyaltyAuthCode(Check_no)
 Else
   MGCD.TransactionType = add
 End If
 MGCD.RegisterName = MyRegisterName
 MGCD.ChargeAmount = 0
 MGCD.CheckNumber = Check_no
 MGCD.CardType = "Loyalty"
 MGCD.Points = Format(Chk_Sales_data(0), "00")
 MGCD.Units = 1
 MGCD.PriceBal = Format(Chk_Sales_data(0), "0.00")
 MGCD.PromoId = "1"
 If GiftCardPayments(MGCD, MGR) = False Then
   Exit Function
 End If
 Set MG = Nothing
End Function
Public Sub RecordProcessingFeeInfo(ProcessorProgram As String, Percentage As Single)
'this sub records how the processing fees are reported to the hand held units
'CreditPrice or CashPrice
'let the hand held know if it need to increase the price on sync'
On Error GoTo ERH
Dim Db As ADODB.Connection
ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Db
  Sql = "Update [Company Information] Set CreditFeeType = '" & ProcessorProgram & "' , FeePercentage = " & Percentage * 100
Db.Execute Sql
DbClose Db, Nothing
Exit Sub
ERH:
Trans_Log "Error with processor program " & Err.Description
Err.Clear
End Sub
Public Function GiftCardPayments(MGCharge As ver11DebitCards.GiftChargeData, GiftResults As ver11DebitCards.GiftCardResponse) As Boolean
On Error GoTo ERH
                Dim HasValue As Currency
                Dim GCData As ver11DebitCards.MercuryGiftCard
                Set GCData = New ver11DebitCards.MercuryGiftCard
                If MGift Is Nothing Then
                   Set MGift = New AllAccounts
                End If
                Dim ReturnResults As String
                Dim ChargeAmount As Currency
                
                Dim Rsp() As String
                If MGift.Item(MGCharge.AccountNumber) Is Nothing Then MGift.add MGCharge.AccountNumber
                MGift.Item(MGCharge.AccountNumber).AccountNo = MGCharge.AccountNumber
                MGift.Item(MGCharge.AccountNumber).PromoCard = False
DoOver:
                MGCharge.StoreName = StoreName
                MGCharge.StoreAddress = StoreName
                MGCharge.StorePhone = StorePhone
                MGCharge.InOption79 = Val(Check_Option(79)) ' Val(Input_Option(79))
                MGCharge.CheckOp19 = 0
                MGCharge.InOption23 = CardCopies
                ChargeAmount = MGCharge.ChargeAmount
                '// start fran-trak
                If Input_Option(54) = "GiftLocal" Then
                   Dim GiftLocalAccountNo As String
                   If InStr(MGCharge.AccountNumber, "-") > 0 Then
                     GiftLocalAccountNo = Replace(MGCharge.AccountNumber, "-", "")
                   Else
                     GiftLocalAccountNo = MGCharge.AccountNumber
                   End If
                   Dim ReturnValue As GiftLocalResponse
                   
                   Dim GLC As MSaleUtility.Utilities
                   Set GLC = New MSaleUtility.Utilities
                   
                   If MGCharge.TransactionType = VoidSale Then
                     ReturnResults = GLC.VoidGiftLocalVB6(MGCharge.AccountRef)
                     ReturnValue = ReadGiftLocalResponse(ReturnResults)
                     If ReturnValue.Message <> "Success" Then
                        See_Msg ReturnValue.Code, 0, ReturnValue.Message
                        GiftCardPayments = False
                        GiftResults.ApprovalStr = "Declined"
                     Else
                        GiftResults.ApprovalStr = "Approved"
                     End If
                   Else
                     ReturnResults = GLC.ValidateGiftLocalCardVB6(GiftLocalAccountNo, "")
                     ReturnValue = ReadGiftLocalResponse(ReturnResults)
                     
                      
                      If ReturnValue.Message = "Success" Then
                         Select Case ReturnValue.CardValue
                            Case Is = 0
                               HasValue = 0
                               See_Msg "Card has no value", 0, "this card has a 0 balance"
                               GiftCardPayments = False
                               GiftResults.ReturnBalance = 0
                               GiftResults.ApprovalStr = "Declined"
                            Case Is < ChargeAmount
                              HasValue = ReturnValue.CardValue
                              See_Msg "Account Balance Less than Sale", 5, "Sending account balance"
                              MGCharge.ChargeAmount = ReturnValue.CardValue
                              GiftResults.ReturnBalance = ReturnValue.CardValue
                              ReturnResults = GLC.RedeemGiftLocalVB6(ReturnValue.PublicToken, ReturnValue.CardValue)
                            Case Else
                              HasValue = ReturnValue.CardValue
                              If ChargeAmount > 0 Then ReturnResults = GLC.RedeemGiftLocalVB6(ReturnValue.PublicToken, ChargeAmount)
                          End Select
                          If HasValue > 0 Then ReturnValue = ReadGiftLocalResponse(ReturnResults)
                          If ReturnValue.isPromoType = 1 Then
                           If ReturnValue.Message = "Success" Then
                             If HasValue > 0 Then
                                MGift.Item(MGCharge.AccountNumber).PromoCard = True
                                MGift.Item(MGCharge.AccountNumber).CurrentCharge = MGCharge.ChargeAmount
                               Exit Function
                             Else
                               ReturnValue.Message = "No Balance"
                             End If
                           End If
                        '' you are a discount
                          End If
                          If ReturnValue.Message <> "Success" Then
                             See_Msg ReturnValue.Code, 0, ReturnValue.Message
                             GiftCardPayments = False
                             GiftResults.ApprovalStr = "Declined"
                             GiftResults.ReturnBalance = ReturnValue.CardValue
                          Else
                             GiftResults.ApprovalStr = "Approved"
                             GiftResults.ReturnBalance = ReturnValue.CardValue
                             MGift.Item(MGCharge.AccountNumber).AccountId = ReturnValue.VoucherRedemptionId
                          End If
                       Else
                         See_Msg "Card read failed", 0, ReturnValue.Message
                         GiftCardPayments = False
                         GiftResults.ErrorStr = ReturnValue.Message
                         GiftResults.ApprovalStr = "Declined"
                       End If
                  End If
                  Set GLC = Nothing
                Else
                 GiftResults = GCData.ConnectToGiftCard(MGCharge, AllPrinter.Item("Receipt").PrinterName, GetSpoolPrinter(AllPrinter.Item("Receipt").PrinterName))
                End If
                
                If MGCharge.TransactionType = add Then
                 If MGCharge.AuthCode > "" Then
                   SaveDcapLoyaltyAuthCode Check_no, MGCharge
                 End If
                End If
                If GiftResults.ErrorStr > "" Then
                   See_Msg GiftResults.ErrorStr, 5, "Error Processing Gift/Loyalty"
                   Set MGift = Nothing
                   GiftCardPayments = False
                   Exit Function
                End If
                If InStr(GiftResults.ApprovalStr, "Declined") > 0 Then
                  If Not MGift.Item(MGCharge.AccountNumber) Is Nothing Then
                    If GiftResults.ReturnBalance = 0 Then MGift.remove MGCharge.AccountNumber
                  End If
                  If GiftResults.ApprovalStr = "Declined" Then
                     If GiftResults.ReturnBalance = 0 Then
                        See_Msg "No Card Balance", 0, GiftResults.ApprovalStr
                        MGCharge.ChargeAmount = GiftResults.ReturnBalance
                        Exit Function
                     Else
                       If See_Msg("Apply Remaining Balance of " & Format(GiftResults.ReturnBalance, "0.00"), 4, GiftResults.ApprovalStr) = 6 Then
                        MGCharge.ChargeAmount = GiftResults.ReturnBalance
                         GoTo DoOver
                       Else
                         Exit Function
                       End If
                     End If
                  Else
                     See_Msg "Declined ", 0, GiftResults.ApprovalStr & " Card Balance " & Format(GiftResults.ReturnBalance, "0.00")
                     MGCharge.ChargeAmount = 0
                     Exit Function
                  End If
                Else
                  If MGCharge.ChargeAmount = 0 And MGCharge.CardType <> "Loyalty" Then
                    See_Msg "Gift Card Balance", 0, "Card Balance " & Format(GiftResults.ReturnBalance, "0.00")
                    PrintGiftChit MGCharge.AccountNumber, Format(GiftResults.ReturnBalance, "0.00"), "Gift Card Balance"
                    Exit Function
                  ElseIf MGCharge.ChargeAmount < Chk_Sales_data(5) And MGCharge.TransactionType = GiftSale Or MGCharge.ChargeAmount < Chk_Sales_data(5) And MGCharge.TransactionType = GiftPartial Then
                    If MGCharge.ChargeAmount = MGCharge.TipAmount Then
                       See_Msg "Approved", 5, "Tip Charged"
                       GiftCardPayments = True
                       Exit Function
                    End If
                    See_Msg "Partial Approval", 0, "Balance Remaining " & Format(Chk_Sales_data(5) - MGCharge.ChargeAmount, "0.00")
                    Trans_Log "Partial Approval Balance " & Format(Chk_Sales_data(5) - MGCharge.ChargeAmount, "0.00")
                    ChargeAmount = MGCharge.ChargeAmount
                    MGift.Item(MGCharge.AccountNumber).CurrentCharge = MGCharge.ChargeAmount
                  Else
                    MGift.Item(MGCharge.AccountNumber).CurrentCharge = MGCharge.ChargeAmount
                    If GiftResults.ApprovalStr = "Approved" Then
                      See_Msg "Approved", 5, "Account Updated"
                      GiftCardPayments = True
                    End If
                  End If
                End If
                Set GCData = Nothing
Exit Function
ERH:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " GiftResults")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function AddGiftCardBonus(ckNo As Long, Svr_Name As String, NetDiscount As Double, BonusType As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Oc As ADODB.Recordset
ConnectADODB CheckStatDb, Db
Sql = "Select * from [Check Detail] "
Sql = Sql & " Where [close out day] = 0 and [check number] = " & ckNo & " and [servers name] = '" & Svr_Name & "'"
Sql = Sql & " and [voided] = 0 and [Menu Item] = '" & BonusType & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
   DbClose Nothing, RS
   Exit Function
End If

'// customer payment amount (5) (16) is the discount amount
'// deduct cash discount at end of sale
            Sql = "Insert Into [Check Detail] ([Check Number],[Check Date]"
            Sql = Sql & ", [Quantity],[Menu Item],[Price]"
            Sql = Sql & ", [Tax Table],[Special Tax],[Liquor Tax],Voided"
            Sql = Sql & ", [Deleted],[Seat Number],[Sort Level],Printed"
            Sql = Sql & ", [Selection Number],[Disc Type],[Disc Amount]"
            Sql = Sql & ", [Register Name],[Menu Number],[Department Report]"
            Sql = Sql & ", [Recorded Time],[Close out Day],[Record Type]"
            Sql = Sql & ", [Modifier Option],[Voided By],[Void Reason]"
            Sql = Sql & ", [Void Time],[Cateory Name],[Servers Name]"
            Sql = Sql & ", [On Hold],[Check Status],[Cashier Name],[Inclusive Tax],[Inclusive Amount],[Cost Amount],[Order Type],[Pos],[Member Name]) Values"
            Sql = Sql & "(" & ckNo
            Sql = Sql & ",'" & CStr(Format(Date, "MM/DD/YYYY")) & "'"
            Sql = Sql & "," & 1  'qty
            Sql = Sql & ",'" & BonusType & "'" ' name
            Sql = Sql & "," & Format(NetDiscount * -1, "0.00") 'price
            Sql = Sql & "," & 1  'SalesTaxAmt.Tax1   '0   'Reg Tax
            Sql = Sql & "," & 1 'SalesTaxAmt.Tax2   'Spc Tax
            Sql = Sql & "," & 0   'liq tax
            Sql = Sql & "," & 0   'Voided
            Sql = Sql & "," & 0  'non sale revenue or delete field name
            Sql = Sql & "," & 0  'SeatNumber    'seat number
            Sql = Sql & "," & 500 'sort level
            Sql = Sql & "," & 1   'printed
            Sql = Sql & "," & -1000 'sort level
            Sql = Sql & ",'None'"   'Discount name
            Sql = Sql & "," & 0     'Discount Amount
            Sql = Sql & ",'" & MyRegisterName & "'"  'Register Name
            Sql = Sql & "," & 1 'menu number
            Sql = Sql & ",'" & ChargeType & "'"  'department
            Sql = Sql & ",'" & CStr(Format(Time, "HH:MM:SS AmPm")) & "'"
            Sql = Sql & "," & 0   'close out day
            Sql = Sql & "," & 1   'Menu Item Number
            Sql = Sql & "," & 0
            Sql = Sql & ",''"
            Sql = Sql & ",''"
            Sql = Sql & ",''"
            Sql = Sql & ",'" & BonusType & "'"
            Sql = Sql & ",'" & LoggedOnAs(0).EMPName & "'"
            Sql = Sql & ",0"    'on hold
            Sql = Sql & "," & 0
            Sql = Sql & ",'" & LoggedOnAs(0).EMPName & "'"
            Sql = Sql & "," & 0
            Sql = Sql & "," & 0
            Sql = Sql & "," & 0     'AllMenuItems(Grid_array(St, Rw, 1)).ItemCost
            Sql = Sql & ",'" & GetRevenueCenter(SvrPath.TerminalGroupName, LoggedOnAs(0).EMPName, Table_Number) & "'"
            Sql = Sql & "," & 100 & ""
            Sql = Sql & ",'" & BonusType & "')"
    Db.Execute Sql
    Grid_array(SeatNumber, NextRow(SeatNumber), 0) = 1
    Grid_array(SeatNumber, NextRow(SeatNumber), 1) = BonusType
    Grid_array(SeatNumber, NextRow(SeatNumber), 2) = Format(NetDiscount * -1, "0.00")  'price
    Grid_array(SeatNumber, NextRow(SeatNumber), 3) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 4) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 5) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 6) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 7) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 8) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 9) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 10) = 1 'SalesTaxAmt.Tax1
    Grid_array(SeatNumber, NextRow(SeatNumber), 11) = 1 'SalesTaxAmt.Tax2
    Grid_array(SeatNumber, NextRow(SeatNumber), 12) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 13) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 14) = -1000
    Grid_array(SeatNumber, NextRow(SeatNumber), 19) = -1000
    Grid_array(SeatNumber, NextRow(SeatNumber), 15) = 500
    Grid_array(SeatNumber, NextRow(SeatNumber), 16) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 17) = "None"
    Grid_array(SeatNumber, NextRow(SeatNumber), 18) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 20) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 21) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 22) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 23) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 24) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 25) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 26) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 27) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 28) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 29) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 30) = Grid_array(0, NextRow(0), 2)
    Grid_array(SeatNumber, NextRow(SeatNumber), 31) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 32) = 0
    Grid_array(SeatNumber, NextRow(SeatNumber), 33) = ""
    Grid_array(SeatNumber, NextRow(SeatNumber), 34) = ""
    NextRow(SeatNumber) = NextRow(SeatNumber) + 1
    Sales_Totals
    Sql = "Select * from [Open Checks] "
    Sql = Sql & " Where [Close Out Day] = 0 and [check number] = " & ckNo & " and [server name] = '" & Svr_Name & "'"
    ConnectADORS Sql, Db, Oc
        If Not Oc.EOF Then
            Sql = "Update [Open Checks] Set  [Gross Discounts] = [Gross Discounts] + " & Abs(NetDiscount) & ",[Check Total] = " & Format(Chk_Sales_data(5), "0.00")
            Sql = Sql & " ,[Gross Sales Tax] = " & Chk_Sales_data(1) & " ,[Gross Special Tax] = " & Chk_Sales_data(2)
            Sql = Sql & " Where [close out day] = 0 and [check number] = " & ckNo & " and [server name] = '" & Svr_Name & "'"
            Db.Execute Sql
        Else
            Sql = "Update [Closed Checks] Set [Gross Discounts] = [Gross Discounts] + " & Abs(NetDiscount) & ",[Check Total] = " & Format(Chk_Sales_data(5), "0.00")
            Sql = Sql & " ,[Gross Sales Tax] = " & Chk_Sales_data(1) & " ,[Gross Special Tax] = " & Chk_Sales_data(2)
            Sql = Sql & " Where [close out day] = 0 and [check number] = " & ckNo & " and [server name] = '" & Svr_Name & "'"
            Db.Execute Sql
        End If
    Oc.Close
Exit Function
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Add Convenience Fee ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function

Public Function BuildConnectionString(Dbase As String, DbaseName As String) As String
Dim PathtoFile As String
PathtoFile = "Initial Catalog=" & DbaseName
Select Case DbaseName
    'Case Is = "House Accounts", "Frequent DinerSql", "BadChecksSql"
    '      BuildConnectionString = "Data Source=" & CPath.TerminalIpAddress & ";" & PathtoFile & ";Persist Security Info=False;User Id='" & CPath.ServerUser & "';Password='" & CPath.ServerPW & "';MultipleActiveResultSets=True;Asynchronous Processing=true;"
    '   Case Else
          BuildConnectionString = "Data Source=" & SvrPath.TerminalIpAddress & ";" & PathtoFile & ";Persist Security Info=False;User Id='" & SvrPath.ServerUser & "';Password='" & SvrPath.ServerPW & "';MultipleActiveResultSets=True;Asynchronous Processing=true;"
    'End Select
End Function
Public Function LoadPrinters(Db As ADODB.Connection, DepartmentName As String, RMTPT As HoldItemPrinters) As Integer
Dim Tb                  As ADODB.Recordset
LoadPrinters = -1
  Sql = "Select * from [Report Categories] where [Report Name] = '" & DepartmentName & "'"
ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   LoadPrinters = 0
   RMTPT.RMTP1 = Tb![Remote 1] & ""
   If Check_Option(64) = 1 Then
     RMTPT.RMTP1 = GetMonitor(Tb![Remote 1] & "")
   End If
   RMTPT.RMTP2 = Tb![Remote 2] & ""
   If Check_Option(64) = 1 Then
     RMTPT.RMTP2 = GetMonitor(Tb![Remote 2] & "")
   End If
   RMTPT.RMTP3 = Tb![Remote 3] & ""
   If Check_Option(64) = 1 Then
     RMTPT.RMTP3 = GetMonitor(Tb![Remote 3] & "")
   End If
   RMTPT.RMTP4 = Tb![Remote 4] & ""
   If Check_Option(64) = 1 Then
     RMTPT.RMTP4 = GetMonitor(Tb![Remote 4] & "")
   End If
   RMTPT.RMTP5 = Tb![Remote 5] & ""
   If Check_Option(64) = 1 Then
     RMTPT.RMTP5 = GetMonitor(Tb![Remote 5] & "")
   End If
   RMTPT.RMTP6 = Tb![Remote 6] & ""
   If Check_Option(64) = 1 Then
     RMTPT.RMTP6 = GetMonitor(Tb![Remote 6] & "")
   End If
End If
DbClose Nothing, Tb
End Function
Public Function GetCreditCardFee() As Integer
Dim FdFee As Integer
For FdFee = 0 To UBound(CreditCards, 1) - 1
 If CreditCards(FdFee).ConvenienceFee > 0 Then
  If CreditCards(FdFee).ReqAuth = TenderTypes.EVM Then
    GetCreditCardFee = FdFee
    Exit Function
  End If
 End If
Next FdFee
End Function
Public Function GetFeeNames(CollectFees() As String) As Integer
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset

If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Function
    Sql = "Select * from [Credit Cards] where [ConvenienceFeeAmt] > 0"
   ConnectADORS Sql, Db, Tb
        
        GetFeeNames = RecordsCount(Tb)
        ReDim CollectFees(GetFeeNames)
        GetFeeNames = -1
        Do While Not Tb.EOF
          GetFeeNames = GetFeeNames + 1
          CollectFees(GetFeeNames) = Tb.fields("Fee Type").Value & ""
         Tb.MoveNext
        Loop
   DbClose Db, Tb
   Exit Function
ERH:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " Save Alt Description")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function wConvertCurrency(FxData As String) As String
 On Error GoTo ERH
 wConvertCurrency = 0
 Dim yahooHTTP As New WinHttp.WinHttpRequest
 yahooHTTP.Open "GET", "https://free.currencyconverterapi.com/api/v6/convert?q=" & Trim$(FxData) & "_USD&compact=ultra&apiKey=" & GetSetting("Micro$ale", "Register Setup", "Currency API Key")
 yahooHTTP.Send
 If yahooHTTP.responseText = "{}" Then
   wConvertCurrency = 0
   See_Msg "Nothing Returned for " & FxData, 5, ""
   Trans_Log "Updating Currency Conversion No Data Returned"
   Exit Function
 End If
 wConvertCurrency = Replace(yahooHTTP.responseText, "{", "")
 wConvertCurrency = Replace(wConvertCurrency, "}", "")
 Dim GetItems() As String
 GetItems = Split(wConvertCurrency, ":")
 wConvertCurrency = GetItems(2)
 Exit Function
ERH:
 Trans_Log "Error Updating Currency Conversion " & Err.Description
 wConvertCurrency = 0
End Function
Public Function GetAltName(Name As String) As String
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
GetAltName = GL.GetAltName(Name)
If GetAltName = "" Then GetAltName = "Cash Discount"
Set GL = Nothing
End Function
Public Function GetExp_Date(DepartMentNumber As String) As String
 Trans_Log "Plu Number In " & DepartMentNumber
 Dim GL As GeneralLib.Utilities
 Set GL = New GeneralLib.Utilities
 GetExp_Date = GL.GetExpirationTime(DepartMentNumber)
 Set GL = Nothing
 Trans_Log "Exp Returned " & GetExp_Date
End Function
Public Function GetChinesePrint(DataIn As String) As String
 GetChinesePrint = DataIn
 If UseChineseChr = False Then Exit Function
 If DataIn = "" Then Exit Function
 Dim ALTNames() As String
 Dim GCHChr()   As String
 Dim FN         As Integer
 Dim Data       As String
 Dim LCID_PRC   As Long: LCID_PRC = 2052
 Dim stmUTF8    As ADODB.Stream
 Set stmUTF8 = New ADODB.Stream
 With stmUTF8
     .Type = adTypeText
     .Charset = "utf-8"
     .Open
     .LoadFromFile DBFilePath & "INI Files\" & "Chinese.ini"
      Data = StrConv(StrConv(.ReadText(adReadAll), vbFromUnicode, LCID_PRC), vbUnicode)
     .Close
 End With
Set stmUTF8 = Nothing
ALTNames = Split(Data, vbCrLf)
For FN = 0 To UBound(ALTNames, 1)
  If InStr(ALTNames(FN), DataIn) > 0 Then
   GCHChr = Split(ALTNames(FN), "=")
   GetChinesePrint = GCHChr(1)
   Exit Function
  End If
Next FN
End Function
Public Function Enter_Name(NeedLastName As Boolean) As String
On Error GoTo ERH
If FileCheck(SvrPath.TerminalPath & "Customer Names.MDB") Then
 Dim ConnectStr As String
 Dim UInit As Boolean
 If GetConfigInfo("Prompts", "Prompt Last Initial", "No", "Functions.INI") = "ENABLED" Then
    UInit = True
 Else
    UInit = False
 End If
 Dim GtNm As CustomerName.GetCustomerName
 Set GtNm = New CustomerName.GetCustomerName
 ConnectStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SvrPath.TerminalPath & "Customer Names.mdb;Persist Security Info=False"
 Enter_Name = GtNm.GetName(ConnectStr, UInit)
Set GtNm = Nothing
Else
   Enter_Name = CStr(ShowKeyBoard("Enter Customer Name"))
End If
Exit Function
ERH:
Trans_Log "Error with Enter Customer Name " & Err.Description
End Function
Public Function IsSqlDataBase(DbaseName As String) As String
  IsSqlDataBase = DbaseName
End Function
Public Function LoadOnlyCodes(JobLink As String, TipOut As FinancialItems, AuditNetSales As Currency, AuditOpenNetSales As Currency, NonSaleRev As Currency) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim GC As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
OpenTb "Select * from [Tip Out To] Where [Job Code] = '" & JobLink & "'", RS, Db
  Do While Not RS.EOF
    OpenTb "Select * from [Job Titles] Where [Job Title] = '" & RS.fields("Tips Paid To").Value & "' and [Tip Pool Percent] > 0", GC, Db
     If Not GC.EOF Then
      If TipOut.Item(GC.fields("Job Title").Value & "") Is Nothing Then
       TipOut.add GC.fields("Job Title").Value & ""
      End If
      TipOut.Item(GC.fields("Job Title").Value & "").ItemName = GC.fields("Job Title").Value & ""
      TipOut.Item(GC.fields("Job Title").Value & "").Percentage = GC.fields("Tip Pool Percent").Value
      TipOut.Item(GC.fields("Job Title").Value & "").Amount = Format(AuditNetSales - (NonSaleRev + AuditOpenNetSales), "0.00")
      TipOut.Item(GC.fields("Job Title").Value & "").TipShare = 0
     End If
     GC.Close
    RS.MoveNext
   Loop
   DbClose Db, RS

End Function
Public Function LoadJobTipOut(TipOut As FinancialItems, AuditNetSales As Currency, AuditOpenNetSales As Currency, NonSaleRev As Currency) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
OpenTb "Select * from [Job Titles] Where [Tip Pool Percent] > 0", RS, Db
    
    Do While Not RS.EOF
      If TipOut.Item(RS.fields("Job Title").Value & "") Is Nothing Then
       TipOut.add RS.fields("Job Title").Value & ""
      End If
      TipOut.Item(RS.fields("Job Title").Value & "").ItemName = RS.fields("Job Title").Value & ""
      TipOut.Item(RS.fields("Job Title").Value & "").Percentage = RS.fields("Tip Pool Percent").Value
      TipOut.Item(RS.fields("Job Title").Value & "").Amount = Format(AuditNetSales - (NonSaleRev + AuditOpenNetSales), "0.00")
      TipOut.Item(RS.fields("Job Title").Value & "").TipShare = 0
      RS.MoveNext
     Loop
     DbClose Db, RS
End Function
Public Function CollectServerTipShare(ServerName As String, BegCheckNo As Long, AuditStartDate As String, AuditStartTime As String) As Currency
Dim TipOut As FinancialItems
Set TipOut = New FinancialItems
Dim PT     As Integer
LoadJobTipOut TipOut, 0, 0, 0
GetServerTipOut ServerName, BegCheckNo, TipOut
CollectServerTipShare = 0
For PT = 1 To TipOut.count
 If TipOut.Item(PT).Amount > 0 Then
   If JobCodeHasWorked(AuditStartDate, TipOut.Item(PT).ItemName) And TipOutThisJobCode(GetSvrJobCode(ServerName, AuditStartDate, AuditStartTime), TipOut.Item(PT).ItemName) Then
    CollectServerTipShare = CollectServerTipShare + Format(TipOut.Item(PT).Percentage * TipOut.Item(PT).Amount, "0.00")
   End If
 End If
Next PT
Set TipOut = Nothing
End Function
Public Function GetUpdateInfo() As String
On Error GoTo ERH
Dim Db                  As ADODB.Connection
Dim Tb                  As ADODB.Recordset
If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Function
  Sql = "Select * from [Terminal Update] where [Name] = '" & MyRegisterName & "' and [date] <= '" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
If Not Tb.EOF Then
   GetUpdateInfo = Tb.fields("Database").Value & ""
   Tb.delete
End If
DbClose Db, Tb
Exit Function
ERH:
GetUpdateInfo = ""
See_Msg Err.Description, 5, "GetUpdate Information"
Trans_Log Err.Description & " GetUpdateInfo"
End Function
Public Function LoadUpdates() As Boolean
On Error GoTo ERH
Dim CT      As Integer
Dim DBName  As String
LoadUpdates = False
Dim SF As VER11SystemMaintenance.LocationData
Set SF = New VER11SystemMaintenance.LocationData
   CT = 0
   Do
    DBName = GetUpdateInfo
    If DBName = "" Or CT > 10 Then Exit Do
    FileCopy SvrPath.LocalTerminalPath & DBName, SvrPath.LocalTerminalPath & DBName & ".BAK"
    See_Msg "Loading Changes", 5, "To " & DBName
    LoadUpdates = True
    SF.LocationUpdate Nothing, Nothing, SvrPath.LocalTerminalPath, DBName
    CT = CT + 1
   Loop
Set SF = Nothing
Exit Function
ERH:
 See_Msg Err.Description & " Load Updates", 5, Err.Number
 Trans_Log Err.Description & " Load Updates"
End Function
Public Sub ImportDepartments()
On Error GoTo ErrorHandler
Dim iDb As ADODB.Connection
Dim iRs As ADODB.Recordset
Dim EDb As ADODB.Connection
Dim eRs As ADODB.Recordset
Dim SqlStr As String
If ConnectADODB("MSaleData", iDb) = False Then Exit Sub
If ConnectADODB(SvrPath.TerminalPath & "CRMENU.MDB", EDb) = False Then Exit Sub
Sql = "Select * from [report categories]"
If ConnectADORS(Sql, EDb, eRs) = False Then

End If
If Not eRs.EOF Then
  SqlStr = "Delete from [report Categories]"
  iDb.Execute SqlStr
End If
If ConnectADORS(Sql, iDb, iRs) = False Then

End If
Do While Not eRs.EOF
 iRs.AddNew
 iRs.fields("Report Name").Value = eRs.fields("Report Name").Value & ""
 iRs.fields("Flag").Value = Val(eRs.fields("Flag").Value & "")
 iRs.fields("Record Number").Value = Val(eRs.fields("Record Number").Value & "")
 iRs.fields("Remote 1").Value = eRs.fields("Remote 1").Value & ""
 iRs.fields("Remote 2").Value = eRs.fields("Remote 2").Value & ""
 iRs.fields("Remote 3").Value = eRs.fields("Remote 3").Value & ""
 iRs.fields("Remote 4").Value = eRs.fields("Remote 4").Value & ""
 iRs.fields("Remote 5").Value = eRs.fields("Remote 5").Value & ""
 iRs.fields("Remote 6").Value = eRs.fields("Remote 6").Value & ""
 iRs.fields("Master Name").Value = eRs.fields("Master Name").Value & ""
 iRs.fields("Timed Event 1").Value = eRs.fields("Timed Event 1").Value & ""
 iRs.fields("Timed Event 2").Value = eRs.fields("Timed Event 2").Value & ""
 iRs.fields("Timed Event 3").Value = eRs.fields("Timed Event 3").Value & ""
 iRs.fields("Timed Event 4").Value = eRs.fields("Timed Event 4").Value & ""
 iRs.fields("Timed Event 5").Value = eRs.fields("Timed Event 5").Value & ""
 iRs.fields("Department No").Value = Val(eRs.fields("Department No").Value & "")
 iRs.Update
 eRs.MoveNext
Loop
DbClose EDb, eRs
DbClose iDb, iRs
Exit Sub
ErrorHandler:
End Sub

Public Function GetDeptPrinter(DepartmentName As String) As String
Dim Db                  As ADODB.Connection
Dim Tb                  As ADODB.Recordset
ConnectADODB SvrPath.LocalTerminalPath & "CRMENU.MDB", Db
  Sql = "Select * from [Report Categories] where [Report Name] = '" & DepartmentName & "'"
ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   GetDeptPrinter = Tb![Remote 1] & ""
 End If
DbClose Db, Tb
End Function
Public Sub RegUnReg(ByVal inFileSpec As String)
    On Error Resume Next
    Dim inHandle As String
    Dim lLib As Long                 ' Store handle of the control library
    Dim lpDLLEntryPoint As Long      ' Store the address of function called
    Dim lpThreadID As Long           ' Pointer that receives the thread identifier
    Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
    Dim mThread    As Long
    Dim mresult    As Long
    lLib = LoadLibrary(inFileSpec)
    If lLib = 0 Then
       ' See_Msg "Load Lib Failed...", 5, ""
        Exit Sub
    End If
    inHandle = ""
      ' Find and store the DLL entry point, i.e. obtain the address of the
      ' “DllRegisterServer” or "DllUnregisterServer" function (to register
      ' or deregister the server’s components in the registry).
      '
    If inHandle = "" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    ElseIf inHandle = "U" Or inHandle = "u" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
    Else
       See_Msg "UnKnow command Handle", 5, ""
        Exit Sub
    End If
    If lpDLLEntryPoint = vbNull Then
        GoTo earlyExit1
    End If
    
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    If mThread = 0 Then
        GoTo earlyExit1
    End If
    
      ' Use WaitForSingleObject to check the return state (i) when the specified object
      ' is in the signaled state or (ii) when the time-out interval elapses.  This
      ' function can be used to test Process and Thread.
    mresult = WaitForSingleObject(mThread, 10000)
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
      ' We don't call the dangerous TerminateThread(); after the last handle
      ' to an object is closed, the object is removed from the system.
    CloseHandle mThread
    FreeLibrary lLib
    
    Exit Sub
    
    
earlyExit1:
    Screen.MousePointer = vbDefault
    MsgBox "Process failed in obtaining entry point or creating thread."
     ' Decrements the reference count of loaded DLL module before leaving
    FreeLibrary lLib
    Exit Sub
    
earlyExit2:
    MsgBox "Process failed in signaled state or time-out."
    FreeLibrary lLib
     ' Terminate the thread to free up resources that are used by the thread
     ' NB Calling ExitThread for an application's primary thread will cause
     ' the application to terminate
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode
End Sub
Public Function GetEmpId(ByVal FirstName As String, ByVal LastName As String) As String
    On Error GoTo ErrorHandler

    Dim EDb As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim RS  As ADODB.Recordset
    Dim result As String

    result = "0" ' default if not found

    ' open connection
    ConnectADODB SvrPath.LocalTerminalPath & "EMPLOYEE.MDB", EDb

    ' parameterized query
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = EDb
        .CommandText = "SELECT [Company Id] FROM [General Data] WHERE [First Name] = ? AND [Last Name] = ?"
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("FirstName", adVarChar, adParamInput, 50, FirstName)
        .Parameters.Append .CreateParameter("LastName", adVarChar, adParamInput, 50, LastName)
    End With

    Set RS = cmd.Execute
    If Not (RS.EOF And RS.BOF) Then
        If Not IsNull(RS.fields("Company Id").Value) Then
            result = CStr(RS.fields("Company Id").Value)
        End If
    End If

    ' sanity fallback if the "id" isn't numeric
    If Not IsNumeric(result) Then result = "111111"

    GetEmpId = result

cleanexit:
    DbClose EDb, RS
    Set cmd = Nothing
    Exit Function

ErrorHandler:
    ' in case of error, return 0
    GetEmpId = "0"
    Resume cleanexit
End Function

Public Function ReturnTenderName(NameIn As String) As String
On Error Resume Next
      If InStr(NameIn, "MASTER") > 0 Then
           ReturnTenderName = "Master Card"
      ElseIf InStr(NameIn, "AMEX") > 0 Then
           ReturnTenderName = "Amex"
      ElseIf InStr(NameIn, "VISA") > 0 Then
           ReturnTenderName = "Visa"
      ElseIf InStr(NameIn, "DISC") > 0 Then
           ReturnTenderName = "Discover"
      ElseIf InStr(NameIn, "DIN") > 0 Then
           ReturnTenderName = "Diners"
      ElseIf InStr(NameIn, "JC") > 0 Then
           ReturnTenderName = "JCB"
      ElseIf InStr(NameIn, "DCVR") > 0 Then
           ReturnTenderName = "Discover"
      ElseIf InStr(NameIn, "M/C") > 0 Then
           ReturnTenderName = "Master Card"
      Else
           ReturnTenderName = NameIn
      End If
End Function
Public Function IsEMVRapidConnect() As Boolean
If GetConfigInfo("Credit Cards", "Is RapidConnect", "No", "FUNCTIONS.INI") = "ENABLED" Then
   IsEMVRapidConnect = True
   Exit Function
End If

If InStr(UCase(RAPIDCONNECT), "RAPIDCONNECT") > 0 Then
   IsEMVRapidConnect = True
Else
   IsEMVRapidConnect = False
End If
End Function
Public Function GetVoidTotal(CheckNo As Long, SvrName As String) As Currency
On Error GoTo ermsg
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(CheckStatDb, Db) = False Then Exit Function
  Sql = "Select Distinct Sum([Price]) as TVoid from [Check Detail]  Where [Voided] = 1 and [Close Out Day] = 0 and [Check Number] = " & CheckNo & " and [Servers Name] = '" & SvrName & "'"
If ConnectADORS(Sql, Db, Tb) = False Then
     Trans_Log "Error getting void check total"
     Exit Function
  End If
  If Not Tb.EOF Then
      GetVoidTotal = Abs(Val(Tb.fields("TVoid").Value & ""))
  Else
      GetVoidTotal = 0
  End If
  DbClose Db, Tb
  Exit Function
ermsg:
  MsgBox Err.Description, 5, "Get Void Total"
End Function
Public Sub MatchHeader(ServerName As String)
If DualPricing = 0 And ProcessingFeesOn = False Then Exit Sub
If GetConfigInfo("Misc", "Match Check Detail With Header", "Yes", "FUNCTIONS.INI") = "DISABLE" Then Exit Sub
'// added kill switch for header match.
'// this function will make sure that the detail and gross sales match and will insert a closed check header is one is missing
'// bug from big apple bagel
On Error GoTo ermsg
Update.Show
Update.UPdateWhere.Caption = "Reading Sales File"
Update.Refresh
Dim Db      As ADODB.Connection
Dim Cd      As ADODB.Recordset
Dim ch      As ADODB.Recordset
Dim aCD     As ADODB.Recordset
Dim R       As Long
Dim T_Rec   As Long
If ConnectADODB(CheckStatDb, Db) = False Then Exit Sub
  Sql = "SELECT [Check Number],[Servers Name], ABS(SUM([Price])) AS GSales,sum([Inclusive Amount]) as TInclusive FROM [Check Detail] WHERE [Close Out Day] = 0 "
  If ServerName > "" Then
    Sql = Sql & " and [Servers Name] = '" & ServerName & "'"
  End If
  Sql = Sql & "GROUP BY [Check Number], [Servers Name]"
  If ConnectADORS(Sql, Db, Cd) = False Then
     Trans_Log "matching check detail"
     Exit Sub
  End If
  T_Rec = RecordsCount(Cd)
  R = 0
  Do While Not Cd.EOF
    R = R + 1
    Update.Progress.ProgressBarValue = (R / T_Rec) * 100
    Update.Progress.Refresh
    Sql = "Select * from [Closed Checks]  Where [Close Out Day] = 0 and [Check Number] = " & Cd.fields("Check Number").Value
    If ConnectADORS(Sql, Db, ch) = False Then
     Trans_Log "matching check header"
     Exit Sub
    End If
    If Not ch.EOF Then
      If ch.fields("Gross Sales").Value <> (Cd.fields("GSales").Value - Cd.fields("TInclusive").Value) Then
        ch.fields("Gross Sales").Value = Cd.fields("GSales").Value - Cd.fields("TInclusive").Value
        ch.Update
      End If
    Else
      QuickLoadCheck Db, Cd.fields("Check Number").Value, False, Cd.fields("Servers Name").Value & "", False, True, False
      Sales_Totals
      
      Sql = "SELECT * FROM [Check Detail] WHERE [Close Out Day] = 0 and [Check Number] = " & Cd.fields("Check Number").Value
      If ConnectADORS(Sql, Db, aCD) = False Then
       Trans_Log "matching check detail"
       Exit Sub
      End If
      
      ch.AddNew
      ch.fields("Number In Party").Value = 1
      ch.fields("Order Destination").Value = "Dine In"
      ch.fields("Order Date").Value = Format(aCD.fields("Check Date").Value, "mm/dd/yyyy")
      ch.fields("Register Name").Value = MyRegisterName
      ch.fields("Close Out Day").Value = 0
      ch.fields("Transfer From").Value = SvrPath.TerminalGroupName   ' Transfer From <Not used here>
      ch.fields("Server Name").Value = aCD.fields("Servers Name").Value
      ch.fields("Cashier Name").Value = aCD.fields("Servers Name").Value
      ch.fields("Table Number").Value = "Recovered"   'was set to len of 5
      ch.fields("Check Number").Value = aCD.fields("Check Number").Value
      ch.fields("Open Time").Value = Format(aCD.fields("Recorded Time").Value, "HH:MM AmPm")    'Check Opened for open check and open time for closed checks
      ch.fields("Close Time").Value = Format(aCD.fields("Recorded Time").Value, "HH:MM AmPm")   'Check Opened for open check and open time for closed checks
      ch.fields("Gross Sales").Value = Gross_Sales
      ch.fields("Gross Sales Tax").Value = Format(((Chk_Sales_data(1) + Chk_Sales_data(7)) + Chk_Sales_data(12)), "0.00")
      ch.fields("Gross Special Tax").Value = Format(Val(Chk_Sales_data(2)) + (Val(Chk_Sales_data(8))) + Chk_Sales_data(20), "0.00")
      ch.fields("Gross Liquor Tax").Value = Format(Val(Chk_Sales_data(3)) + (Val(Chk_Sales_data(9))), "0.00")
      ch.fields("Gross Discounts").Value = Format((Abs(Val(Chk_Sales_data(4)) - Chk_Sales_data(19)) - Abs(Chk_Sales_data(22))), "0.00")
      ch.fields("Voided Sales").Value = Abs(Val(Chk_Sales_data(6)))
      ch.fields("Voided Sales Tax").Value = 0
      ch.fields("Voided Special Tax").Value = 0
      ch.fields("Voided Liquor Tax").Value = 0
      ch.fields("Voided Discounts").Value = 0
      ch.fields("Actual Cust Count").Value = 1
      ch.fields("Payments").Value = 1
      ch.fields("Check Total").Value = 1
      ch.fields("Seat Number").Value = ""
      ch.fields("Inclusive Total").Value = Format(Abs(Val(Chk_Sales_data(11))) - (Val(Chk_Sales_data(12) + Val(Chk_Sales_data(20)))), "0.00")
      ch.fields("Gross Exempt").Value = Chk_Sales_data(13)
      ch.fields("Voided Exempt").Value = 0
      ch.fields("Member Name").Value = ""
      ch.Update
      aCD.Close
    End If
   Cd.MoveNext
  Loop
  Cd.Close
  DbClose Db, ch
  Unload Update
  Exit Sub
ermsg:
  Trans_Log "Error " & Err.Description & " Match Header"
  DbClose Db, ch
  Err.Clear
End Sub
Public Function SetDualPrice(SalesPrice As Currency) As Currency
If DualPricing > 0 And SalesPrice > 0 Then
  'SetDualPrice = SalesPrice + Int(((SalesPrice * DualPricing) + 0.005) * 100) / 100
  SetDualPrice = SalesPrice + Round(SalesPrice * DualPricing, 2)
  SetDualPrice = RoundPriceTo5or9(CDbl(SetDualPrice))
  '// Return the Credit Card Price
'ElseIf DualPricing < 0 And SalesPrice > 0 Then
'  SetDualPrice = SalesPrice + Int(((SalesPrice * DualPricing) + 0.005) * 100) / 100
  '// Returns then cash value for dual pricing
Else
 SetDualPrice = SalesPrice
End If
End Function
Public Function LoadDualPricesToTab() As String

Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
Dim Results As String
Results = GL.ReloadDualPriceVB6(CDbl(DualPricing))
Set GL = Nothing
Trans_Log "Load Dual Price Table " & Results
End Function
Public Function GetGiftCardSaleAmount(ByVal SalePrice As Currency, ByVal CalculateCash As Boolean) As Currency
Dim GP As Currency
If SalePrice = 0 Then
  GetGiftCardSaleAmount = 0
  Exit Function
End If

If How_Paid(0).PaymentName = "Cash" And CalculateCash = False Then
 GetGiftCardSaleAmount = SalePrice
Else
 GP = DualPricing + 1
 If GP > 0 And SalePrice > 0 Then
  GetGiftCardSaleAmount = Format((SalePrice / GP), "0.00")
 Else
  GetGiftCardSaleAmount = SalePrice
 End If
End If
End Function
' Rounds away-from-zero to 2 decimals, returns Currency (no banker’s rounding)
Private Function Round2(ByVal v As Double) As Currency
    If v >= 0# Then
        Round2 = CCur(Fix(v * 100# + 0.5) / 100#)
    Else
        Round2 = CCur(-Fix(-v * 100# + 0.5) / 100#)
    End If
End Function

Public Function GetCashPrice(ByVal OrderQuantity As Long, _
                             ByVal MenuItemName As String, _
                             ByVal SalesPrice As Currency, ByVal Department As String) As Currency
    Dim GP As Double                 ' 1 + DualPricing (e.g., 1.04 for 4%)
    Dim DP As Currency               ' discount amount per unit (cash-discount path)
    Dim UnitPrice As Currency
    Dim i As Long

    If SalesPrice = 0 Then
        GetCashPrice = 0
        Exit Function
    End If
    If DualPricing = 0 And CashDiscountData.CreditAmount > 0 Then
      If Department > "" And Not AllExemptSales.Item(Department) Is Nothing Then
       'If AllExemptSales.Item(department).ExemptItem > "" Then
        GetCashPrice = SalesPrice
      '  Exit Function
       'End If
      End If
    End If
    MenuItemName = ProcessItemName(MenuItemName, True)
    
    ' If nothing to convert (no dual-pricing and no cash-discount), return as-is
    If DualPricing = 0 And CashDiscountData.CreditAmount <= 0 Then
        GetCashPrice = SalesPrice
        Exit Function
    End If

    If OrderQuantity < 1 Then OrderQuantity = 1         ' guard (modifiers, etc.)
    UnitPrice = Round2(SalesPrice / OrderQuantity)      ' split evenly, numeric rounding

    GP = 1# + CDbl(DualPricing)                         ' dual-pricing multiplier

    ' POSITIVE prices
    If SalesPrice > 0 Then
        GetCashPrice = 0
        For i = 1 To OrderQuantity
            If (CashDiscountData.CreditAmount > 0) And (DualPricing = 0) Then
                ' Cash-discount model: unit minus percentage
                DP = Round2(UnitPrice * (CashDiscountData.CreditAmount / 100#))
                GetCashPrice = GetCashPrice + Round2(UnitPrice - DP)
            Else
                ' Dual-pricing model: deflate by (1 + surcharge) then apply 5/9 rule
                GetCashPrice = GetCashPrice + Round2(RoundPriceTo5or9(UnitPrice / GP))
            End If
        Next

    ' NEGATIVE prices (discount items, comps, etc.)
    Else
        If Left$(MenuItemName, 1) = "%" Then
            ' Percentage-style discount item under dual-pricing: deflate
            GetCashPrice = Round2(SalesPrice / GP)
        ElseIf (CashDiscountData.CreditAmount > 0) And (DualPricing = 0) Then
            DP = Round2(SalesPrice * (CashDiscountData.CreditAmount / 100#))
            GetCashPrice = Round2(SalesPrice - DP)
        Else
            GetCashPrice = SalesPrice
            'GetCashPrice = Round2(RoundPriceTo5or9(unitPrice / GP))
        End If
    End If
End Function

Public Function RollAuthorizationBack(MyCheck As Long, OriginalAuth As Currency) As String
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
  'If Not AllOrders Is Nothing Then AllOrders.Item(1).OrderPaymentAmt = Chk_Sales_data(5)

  ConnectADODB DebitDb, Db
   Sql = "Update [Debit] Set [Amount] = " & OriginalAuth & " Where [Check Number] = " & MyCheck & " and [close out day] = 0"
  Db.Execute Sql
  SetNothing Db
  ConnectADODB CallInOrderDb, Db
   Sql = "Update [Open Checks] Set [Payments] = " & OriginalAuth & ", [Check Total] = " & Chk_Sales_data(5) & " Where [Check Number] = " & MyCheck & " and [close out day] = 0"
  Db.Execute Sql
    Sql = "Update [PrePaid] Set [PaymentAmount] = " & OriginalAuth & " Where [PhoneOrderNumber] = " & MyCheck
  SetNothing Db
  ConnectADODB CheckStatDb, Db
   Sql = "Update [Payments] Set [Payment Amount] = " & OriginalAuth & " Where [Check Number] = " & MyCheck & " and [close out day] = 0"
  Db.Execute Sql
  SetNothing Db
  Trans_Log "EMV roll back to original authorization " & OriginalAuth & " Check No " & MyCheck
  PrintCCSlip MyCheck
  
Exit Function
ErrorHandler:
  Error_Log Err.Description, Err.Description, "Update Delivery", ""
End Function

Public Function UpdateDeliveryValues(MyCheck As Long, UpdateReceipt As Boolean) As String
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim PrePaidAmt As Currency
PrePaidAmt = PhoneOrderPaidAmount(MyCheck)
'// make sure that any prepaid amount gets deducted from the credit card update
'// prepaids dont check only credit card amount can change with preauth
'// example check start 5.00 gift 10.00 credit -- customer all 5.00 additional dollars to the check
'// the new credit card amount will be 15.00. if you dont subtract the prepaid the credit card amount will be over stated.
  If Not AllOrders Is Nothing Then AllOrders.Item(1).OrderPaymentAmt = Chk_Sales_data(5)

  ConnectADODB DebitDb, Db
   Sql = "Update [Debit] Set [Amount] = " & Chk_Sales_data(5) - PrePaidAmt & " Where [Check Number] = " & MyCheck & " and [close out day] = 0"
  Db.Execute Sql
  SetNothing Db
  ConnectADODB CallInOrderDb, Db
   Sql = "Update [Open Checks] Set [Payments] = " & Chk_Sales_data(5) & ", [Check Total] = " & Chk_Sales_data(5) & " Where [Check Number] = " & MyCheck & " and [close out day] = 0"
  Db.Execute Sql
  ConnectADODB CheckStatDb, Db
   Sql = "Update [Payments] Set [Payment Amount] = " & Chk_Sales_data(5) & " Where [Check Number] = " & MyCheck & " and [close out day] = 0"
  Db.Execute Sql
  SetNothing Db
  Trans_Log "EMV delivery update totals " & MyCheck
  If UpdateReceipt Then PrintCCSlip MyCheck
  
Exit Function
ErrorHandler:
  Error_Log Err.Description, Err.Description, "Update Delivery", ""
End Function
Public Sub PrintCCSlip(ckNo As Long)
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim RTB         As ADODB.Recordset
Dim CCD         As CCData
  'Added for captiol cash 4/25/2018
 Sql = "Select * From [Debit] Where [Close Out day] = 0 and [check Number] = '" & CStr(ckNo) & "'"
 ConnectADODB DebitDb, Db
 ConnectADORS Sql, Db, RTB
   If RTB.EOF Then
    See_Msg "No Payment For This Transaction ", 5, "Check Number " & CStr(ckNo)
    DbClose Db, RTB
    Exit Sub
   End If
   Do While Not RTB.EOF
        With CCD
         .HeaderData = "Bar Tab"
         .HeaderReason = "EMV-"
         .AuthNo = RTB.fields("Account Number").Value & ""
         .SaleAmt = Format(RTB.fields("Amount").Value, "0.00")
         .TipAmt = Format(RTB.fields("Tip").Value, "0.00")
         .AccNo = Right(RTB.fields("Card Number").Value & "", 4)
         .CustName = GetTabNameThisCheck(ckNo)
         .CType = RTB.fields("Account Name").Value
         .ServerName = ReturnNickName(RTB.fields("Server Name").Value)
         .TBLName = ""
         .ChkNO = CStr(ckNo)
         .ProcessorResponse = "Approved"
         .Datetime = RTB.fields("Date").Value & RTB.fields("Time").Value
         .ckTotal = Format(RTB.fields("Amount").Value, "0.00")
         End With
         Emv_Back.r_ReturnResult = RTB.fields("XMLString").Value
         PrintSlip CCD
      RTB.MoveNext
     Loop
    DbClose Db, RTB
Exit Sub
ErrorHandler:
Err.Clear
End Sub
'Public Function SafeCCur(v As Variant, Optional ByVal defaultValue As Currency = 0) As Currency
 '   If IsNull(v) Or IsEmpty(v) Then
 '       SafeCCur = defaultValue
 '   ElseIf VarType(v) = vbString Then
 '       Dim s As String: s = Trim$(CStr(v))
 '       If s = "" Then
 '           SafeCCur = defaultValue
 '       ElseIf IsNumeric(s) Then
 '           SafeCCur = CCur(s)
 '       Else
 '           SafeCCur = defaultValue
 '       End If
 '   ElseIf IsNumeric(v) Then
 '       SafeCCur = CCur(v)
 '   Else
 '       SafeCCur = defaultValue
 '   End If
'End Function

Public Function OutOfScope(Employee As String, Tender As String, ChargeAmount As Double, Charge As ChargeTo) As String
    '// routine that handles all EMV Sales
    On Error GoTo ErrorHandler
    OutOfScope = ""
    ResetCardData

    With Emv_Back
       .r_AuthAmt = 0
       .r_CardBalance = 0
       .r_Message = "ABORTED"
       .r_Status = "ABORTED"
    End With
    Dim MyTip As Currency
    If Chk_Sales_data(5) > 0 Then
      If ChargeAmount > Chk_Sales_data(5) Then
       MyTip = ChargeAmount - Chk_Sales_data(5)
       ChargeAmount = Chk_Sales_data(5)
       If See_Msg("Continue ?", 4, "Tip amount is " & Format(MyTip, "0.00")) <> 6 Then Exit Function
      End If
    End If
    Dim PreAuthCard     As Integer
    Dim BalanceDue      As Currency
    Dim CreditFee       As Currency
    Dim FeeAmount       As Currency
    Dim CapSignature    As Boolean
    Dim Results         As String
    Dim Sales           As AllEMVSales
    Dim PMG             As PAXMSG
    Dim LUPRsp          As String
    Dim CA              As ver11DebitCards.Accounting
    Dim EmvEquip        As EMVProcessLib.EmvDeviceHardWare
    Dim HKData          As String
    Dim CheckIsPaid     As Boolean
    Dim TPN             As String
    If Charge = LevelUpTender Then
       'no longer used
    Else
     Dim EMVD As ver11DebitCards.EMVTransaction
     Set EMVD = New ver11DebitCards.EMVTransaction
     Dim MYEMV As EMVOptions
     MYEMV.DisableTip = False
     With MYEMV
          .SaleAmt = ChargeAmount
          .EBTAmt = GetEBT_Total
          .KeySale = False
          .TipAmt = MyTip
     End With
     If LoadEMVDevice(MyRegisterName, EmvEquip) = "Failed" Then
        See_Msg "No Device Found", 5, "Check Device Setup"
        Exit Function
     End If
     Order.Timer2.Enabled = False
     
     CapSignature = False
      If EmvEquip.Item(1).PaxUseSign = 1 Then CapSignature = True
      If IsDejavooInterface(EMVINTERFACE) Then
       TPN = EmvEquip.Item(1).deviceId
      Else
       TPN = ""
      End If
      If (EmvEquip.Item(1).DisableDebit = 0) Or (Check_Option(133) = 1) Or IsDataCapInterface(EMVINTERFACE) Then  'for datacap and dcap direct      '133 = Emv prompt for tip
       If Check_Option(133) = 1 And Check_no <> -9999 Then
        If EmvEquip.Item(1).TipLineIs = "Prompt" Then   '// set so no prompt for tip on order screen if prompt is on the pin pad
         If EmvEquip.Item(1).DisableDebit = 0 Or IsDataCapInterface(EMVINTERFACE) Then
           MYEMV = EMVD.ShowEMV_Menu(SaleType.RegularSale, MYEMV.SaleAmt, MYEMV.TipAmt, MYEMV.EBTAmt)
         Else
           MYEMV.CashBack = 0
           MYEMV.QuitSale = False
           MYEMV.SaleType = "Credit"
         End If
        Else
          MYEMV = EMVD.ShowEMV_Menu(SaleType.CConfirm, MYEMV.SaleAmt, MYEMV.TipAmt, MYEMV.EBTAmt)
        End If
       Else
        MYEMV = EMVD.ShowEMV_Menu(SaleType.RegularSale, MYEMV.SaleAmt, MYEMV.TipAmt, MYEMV.EBTAmt)
       End If
      Else
       MYEMV.CashBack = 0
       MYEMV.QuitSale = False
       MYEMV.SaleType = "Credit"
      End If
     'End If
     If MYEMV.QuitSale Then
       Amount_Tend.CancelConvenienceFee
       With Emv_Back
        .r_AuthAmt = 0
        .r_CardBalance = 0
        .r_Message = "ABORTED"
        .r_Status = "ABORTED"
       End With
       Set EMVD = Nothing
       If IsThisCheckOpen(Check_no) = False Then
         AbortSale = True
       End If
       Order.Timer2.Enabled = True
       
       Exit Function
     End If
     Set Sales = New AllEMVSales
     If Sales.Item(CStr(Check_no)) Is Nothing Then
        Sales.add CStr(Check_no)
     End If
     If MYEMV.SaleType = "EBT" Then
        Sales.Item(CStr(Check_no)).SalesAmt = Format(MYEMV.EBTAmt, "0.00")
        Sales.Item(CStr(Check_no)).TipAmt = 0
     Else
        If MYEMV.SaleType = "Debit" Then
          CreditFee = GetCurrentProcessingFee
          If See_Msg("Apply Tip?", 4, "") = 6 Then
           MYEMV.TipAmt = ShowNumberPad("Enter Amount", False, 0, 0)
          Else
           MYEMV.TipAmt = 0
          End If
          
          If CreditFee > 0 Then
           MYEMV.SaleAmt = MYEMV.SaleAmt - CreditFee
           SetProcessingFeeToZero
           CreditProcessingFee Nothing, MYEMV.SaleAmt
           CreditFee = 0
          End If
          Sales.Item(CStr(Check_no)).SalesAmt = Format(MYEMV.SaleAmt + MYEMV.TipAmt, "0.00")
          Sales.Item(CStr(Check_no)).TipAmt = 0
        Else
          If IsDataCapInterface(EMVINTERFACE) Then  '// do not send tip included in the sale amount
            Sales.Item(CStr(Check_no)).SalesAmt = Format(MYEMV.SaleAmt, "0.00")  '- MYEMV.TipAmt do not deduct tip from sale amount
          Else
            Sales.Item(CStr(Check_no)).SalesAmt = Format(MYEMV.SaleAmt + MYEMV.TipAmt, "0.00")
          End If
          Sales.Item(CStr(Check_no)).TipAmt = MYEMV.TipAmt
        End If
     End If
     If MYEMV.KeySale Then
      Sales.Item(CStr(Check_no)).AccountNo = "Prompt"
     Else
      Sales.Item(CStr(Check_no)).AccountNo = ""
     End If
     Sales.Item(CStr(Check_no)).TaxAmt = NetSalesTax
     Sales.Item(CStr(Check_no)).Cashier = ReturnNickName(Employee) ' LoggedOnAs(0).EMPName
     Sales.Item(CStr(Check_no)).Term = MyRegisterName
     If MYEMV.SaleType = "EBT" Then MYEMV.SaleType = "EBT_FOODSTAMP"
     Sales.Item(CStr(Check_no)).TranType = MYEMV.SaleType
     Sales.Item(CStr(Check_no)).TransNo = CStr(Check_no)
     Sales.Item(CStr(Check_no)).CashBack = MYEMV.CashBack
     Sales.Item(CStr(Check_no)).IpAddress = ""
     Sales.Item(CStr(Check_no)).gateWayIdNo = ""
     Sales.Item(CStr(Check_no)).XMLString = ""
     If MYEMV.SaleType = "EBT_FOODSTAMP" Then
       MYEMV.SaleAmt = MYEMV.EBTAmt
     End If
     If IsDataCapInterface(EMVINTERFACE) Then '   If EMVINTERFACE = "DATACAP" Then
      If EmvEquip.Item(1).MaxQtyOffline = 1 Then 'used for regisiter with processor for processing fees
        If ProcessingFeesOn Then
         CreditFee = GetCurrentProcessingFee + Chk_Sales_data(21)
         '// added the sales tax to fee. if debit check total will be correct otherwise it will be inflated by the
         '// sales tax amount.
         Sales.Item(CStr(Check_no)).PassThru = Format(CreditFee, "0.00")
         Sales.Item(CStr(Check_no)).SalesAmt = Format(MYEMV.SaleAmt - CreditFee, "0.00")
        End If
      End If
      If MYEMV.KeySale Then
       If EmvEquip.Item(1).TipLineIs = "Prompt" Then
         If MYEMV.TipAmt = 0 Then
            MYEMV.TipAmt = Val(ShowNumberPad("Enter Tip Amount?", False, 0, 0))
            If MYEMV.TipAmt = 0 Then
             See_Msg "No tip was entered", 5, ""
            Else
             Sales.Item(CStr(Check_no)).TipAmt = MYEMV.TipAmt
            End If
         End If
       End If
      End If
     Else
      If IsPAXInterface(EMVINTERFACE) Or IsDejavooInterface(EMVINTERFACE) Then
       
       If IsDejavooInterface(EMVINTERFACE) Then
        Dim GL As GeneralLib.Utilities
        Set GL = New GeneralLib.Utilities
        If GL.GetKeyData("Dejavoo Demo") > "" Then
         See_Msg "Running In Demo", 5, "No Live Cards"
        End If  '// only used to check for demo
        Set GL = Nothing
       End If
       '// end DEJAVOO
       If EmvEquip.Item(1).MaxQtyOffline = 1 Then 'used for regisiter with processor for processing fees
        If ProcessingFeesOn Then
         CreditFee = GetCurrentProcessingFee
         Sales.Item(CStr(Check_no)).PassThru = Format(CreditFee, "0.00") * 100
         Sales.Item(CStr(Check_no)).SalesAmt = Format(MYEMV.SaleAmt - CreditFee, "0.00")
        End If
       End If
      End If
     End If
     If MYEMV.KeySale Then
       See_Msg "Please Hand Key ", 5, MYEMV.SaleType & " Amount " & Format(MYEMV.SaleAmt + MYEMV.CashBack, "0.00")
     Else
       See_Msg "Please Swipe Card or Insert Card", 5, MYEMV.SaleType & " Amount " & Format(MYEMV.SaleAmt + MYEMV.CashBack, "0.00")
     End If
     Set PMG = New PAXMSG
     Dim PXFO As EMVProcessLib.Process
     Set PXFO = New EMVProcessLib.Process
     PreAuthCard = 0 '// set value for non preauth as default
     
     If IsDataCapInterface(EMVINTERFACE) Then  'If EMVINTERFACE = "DATACAP" Then
        If EMVSUPPORTPREAUTH = 1 Then
          Results = PXFO.ProcessSale(Sales, DevSaleType_EMVPreAuth, PMG, EmvEquip)
          PreAuthCard = -1
        Else
          Results = PXFO.ProcessSale(Sales, DevSaleType_EMV_Sale, PMG, EmvEquip)
        End If
        If Not PXFO Is Nothing Then PXFO.PadReset EmvEquip
     ElseIf IsDejavooInterface(EMVINTERFACE) Then
          Results = PXFO.ProcessSale(Sales, DevSaleType_djvSale, PMG, EmvEquip)
     Else
          If IsPAXInterface(EMVINTERFACE) Then MYEMV.KeySale = False
          Results = PXFO.ProcessSale(Sales, DevSaleType_XSale, PMG, EmvEquip)
     End If
     Trans_Log EMVINTERFACE & " " & Results
     If UCase(Results) = "TIME OUT" Then
        See_Msg "Time out connecting to pin pad", 5, ""
        With Emv_Back
       .r_AuthAmt = 0
       .r_CardBalance = 0
       .r_Message = "ABORTED"
       .r_Status = "ABORTED"
       End With
       Set EMVD = Nothing
       Order.Timer2.Enabled = True
       Exit Function
     End If
     If Results = "CONNECT ERROR" Then
       Sales.Item(CStr(Check_no)).XMLString = PMG.Item(1).ExtraData
       Call PXFO.ProcessSale(Sales, DevSaleType_XVoidbyRecordNo, PMG, EmvEquip)
       See_Msg Results, 5, ""
       With Emv_Back
       .r_AuthAmt = 0
       .r_CardBalance = 0
       .r_Message = "ABORTED"
       .r_Status = "ABORTED"
       End With
       Set EMVD = Nothing
       Order.Timer2.Enabled = True
       Exit Function
     End If
     If PMG.Item(1) Is Nothing Then
       See_Msg Results, 5, ""
       With Emv_Back
        .r_AuthAmt = 0
        .r_CardBalance = 0
        .r_Message = "ABORTED"
        .r_Status = "ABORTED"
       End With
       Set EMVD = Nothing
       Order.Timer2.Enabled = True
       Exit Function
     End If
     OutOfScope = UCase(PMG.Item(1).ResultTxt)
     Sales.Item(CStr(Check_no)).AccountNo = "xxxx-xxxx-xxxx-" & PMG.Item(1).Last4
     Emv_Back.r_Status = UCase(PMG.Item(1).ResultTxt)
     Emv_Back.r_AuthAmt = Format(SafeCCur(PMG.Item(1).ApprovedAmount), "0.00")
     Emv_Back.r_AuthCode = PMG.Item(1).AuthCode
     Emv_Back.r_Cardtype = PMG.Item(1).CardType
     If Emv_Back.r_Cardtype = "" Then Emv_Back.r_Cardtype = "Other"
     Emv_Back.r_CardBalance = SafeCCur(PMG.Item(1).RemainingBalance)
     Emv_Back.r_TipAmt = SafeCCur(PMG.Item(1).TipAmount)
     If IsDejavooInterface(EMVINTERFACE) Then ' Or IsDataCapInterface(EMVINTERFACE) Then
     
     Else
      If Emv_Back.r_TipAmt > 0 Then
       'If MYEMV.SaleAmt + Emv_Back.r_TipAmt = Emv_Back.r_AuthAmt Then
         Emv_Back.r_AuthAmt = Emv_Back.r_AuthAmt - Emv_Back.r_TipAmt
       'End If
      End If
     End If
     Emv_Back.r_SaleType = MYEMV.SaleType
     Emv_Back.r_Last4 = PMG.Item(1).Last4
     Emv_Back.r_Server = Employee
     If TPN > "" Then
       PMG.Item(1).ExtraData = PMG.Item(1).ExtraData & "<TPN>" & TPN & "</TPN>"
     End If
     If PMG.Item(1).RefNum > "" Then
         Emv_Back.r_ReturnResult = PMG.Item(1).ExtraData & "<Id>" & PMG.Item(1).RefNum & "</Id>"
     Else
        Emv_Back.r_ReturnResult = PMG.Item(1).ExtraData
     End If
     Dim xmls As MPxmlParse
     Set xmls = New MPxmlParse
     xmls.xml = Emv_Back.r_ReturnResult
     Emv_Back.r_CustomerName = Trim$(xmls.token("PLNameOnCard"))
     FeeAmount = Format(SafeCCur(xmls.token("FeeAmount")), "0.00")
     Emv_Back.r_PinStatus = ""
     If MYEMV.SaleType = "Debit" Then
        Emv_Back.r_TipAmt = MYEMV.TipAmt
        If Val(xmls.token("PINStatusNum")) = 1 Then
          Emv_Back.r_PinStatus = "PIN Verified"
        Else
          Emv_Back.r_PinStatus = "PIN Bypassed"
        End If
     End If
     If MYEMV.SaleType = "EBT_FOODSTAMP" Then
        If Val(xmls.token("PINStatusNum")) = 1 Then
          Emv_Back.r_PinStatus = "PIN Verified"
        Else
          'Emv_Back.r_PinStatus = "PIN Bypassed"
          '// no pin by pass with EBT
        End If
     End If
     If Val(xmls.token("AmountDue")) > 0 Then
      Trans_Log "Auth Amount " & Format(Emv_Back.r_AuthAmt, "0.00")
      Trans_Log "Sale Amount " & Format(MYEMV.SaleAmt, "0.00")
      Trans_Log "Amount Due " & xmls.token("AmountDue")
      Emv_Back.r_TipAmt = 0
       ''// must set tip to 0 with partial approval
     End If
     Set xmls = Nothing
     Emv_Back.r_Message = PMG.Item(1).Message
     Emv_Back.r_CashBack = Val(PMG.Item(1).CashBack)
     If PMG.Item(1).Message = "APPROVED OFFLINE" Then
      Trans_Log "Message     " & PMG.Item(1).Message
      Trans_Log "Results     " & PMG.Item(1).ResultTxt
      Trans_Log "Raw Respone " & PMG.Item(1).RawResponse
      'Emv_Back.r_Status = "APPROVAL"
      Emv_Back.r_Message = "OK"
     End If
     If Emv_Back.r_Message = "" Then Emv_Back.r_Message = "Canceled"
     Trans_Log Emv_Back.r_Status & " " & PMG.Item(1).Message & " " & PMG.Item(1).ResultTxt
     Set PMG = Nothing
     Set PXFO = Nothing
     Set Sales = Nothing
     Set EMVD = Nothing
     Set EmvEquip = Nothing
     If Emv_Back.r_Message = "OK" And Emv_Back.r_AuthAmt > 0 Then
         Emv_Back.r_Status = "APPROVAL"
     End If
     
     If IsApproved(Emv_Back.r_Status) Then
      
      If IsPAXInterface(EMVINTERFACE) Or IsDataCapInterface(EMVINTERFACE) Then
       If CreditFee > 0 And Emv_Back.r_AuthAmt < MYEMV.SaleAmt Then
          CheckIsPaid = False
          If (CreditFee + Emv_Back.r_AuthAmt) = MYEMV.SaleAmt And MYEMV.SaleAmt = Chk_Sales_data(5) Then
             CheckIsPaid = True
          Else
           'MYEMV.SaleAmt = Emv_Back.r_AuthAmt
          End If
          '// checkispaid value keeps 1 cent rounding for sales tax not allowing the check to close
          Trans_Log " Credit Card Fee " & CreditFee & " Approval " & Emv_Back.r_AuthAmt & " Sale Amount " & MYEMV.SaleAmt
          SetProcessingFeeToZero
          Trans_Log "New Check Total Is " & Chk_Sales_data(5)
          CreditProcessingFee Nothing, Chk_Sales_data(5)
          If CheckIsPaid Then
            Emv_Back.r_AuthAmt = Chk_Sales_data(5)
          End If
          ChargeAmount = Emv_Back.r_AuthAmt
       End If '// removes the processing fee for bank cards.
      End If
      
      If IsDejavooInterface(EMVINTERFACE) Then
       If Emv_Back.r_AuthAmt + FeeAmount < Chk_Sales_data(5) Then
          Trans_Log " Credit Card Fee " & CreditFee & " Approval " & Emv_Back.r_AuthAmt & " Sale Amount " & MYEMV.SaleAmt
          SetProcessingFeeToZero
          
          CreditProcessingFee Nothing, Chk_Sales_data(5)
          ChargeAmount = Emv_Back.r_AuthAmt
       Else
          Emv_Back.r_AuthAmt = Emv_Back.r_AuthAmt + FeeAmount
       End If '// removes the processing fee for bank cards for dejavoo
      End If
     
      Set GL = New GeneralLib.Utilities
       Dim ChgCrd As EMVSales
       Set ChgCrd = New GeneralLib.EMVSales
       If ChgCrd.Item(CStr(Check_no)) Is Nothing Then
         ChgCrd.add CStr(Check_no)
       End If
       With ChgCrd
          .Item(CStr(Check_no)).Cashier = ReturnMyCashier
          .Item(CStr(Check_no)).Term = MyRegisterName
          .Item(CStr(Check_no)).Media = Emv_Back.r_Cardtype
          .Item(CStr(Check_no)).AccountNo = Emv_Back.r_AuthCode
          .Item(CStr(Check_no)).TimeOfSale = Format(Time, "HH:MM:SS ampm")
          .Item(CStr(Check_no)).SaleType = "EMV"
          If Emv_Back.r_SaleType = "Debit" Then
            .Item(CStr(Check_no)).ChargeType = 1
            .Item(CStr(Check_no)).Media = "Debit"
          ElseIf Emv_Back.r_SaleType = "EBT_FOODSTAMP" Then
           .Item(CStr(Check_no)).Media = "EBT"
           .Item(CStr(Check_no)).ChargeType = 2  '// need a 2 to make sure we can void
          Else
           .Item(CStr(Check_no)).ChargeType = PreAuthCard
          End If
          .Item(CStr(Check_no)).SalesAmt = Emv_Back.r_AuthAmt
          .Item(CStr(Check_no)).CheckTotal = Emv_Back.r_AuthAmt
          .Item(CStr(Check_no)).TransNo = CStr(Check_no)
          .Item(CStr(Check_no)).ServerName = GetWhoOwnsTheCheck(Employee, Check_no)
          .Item(CStr(Check_no)).TipAmt = Emv_Back.r_TipAmt
          .Item(CStr(Check_no)).XMLString = Emv_Back.r_ReturnResult
          .Item(CStr(Check_no)).CardNo = Emv_Back.r_Last4
          If Not allcards Is Nothing Then
           If allcards.Item("-9999") Is Nothing Then
            .Item(CStr(Check_no)).gateWayIdNo = ""
           Else
           .Item(CStr(Check_no)).gateWayIdNo = allcards.Item("-9999").AccountOnFile
           allcards.Item("-9999").AccountOnFile = ""
           End If
          End If
          '// used to save future phone order check number
       End With
       'Set CA = New ver11DebitCards.Accounting
       Results = GL.ChargeAccount(ChgCrd, SvrPath.TerminalPath)
       If Results <> "Success" Then
          See_Msg "Failure to write to debit database", 0, Results
          Trans_Log "Failure to write to debit database"
       End If
       SaveEMVPayments ChgCrd.Item(CStr(Check_no)).TimeOfSale, Check_no, ChgCrd.Item(CStr(Check_no)).ServerName, ChgCrd.Item(CStr(Check_no)).SalesAmt, ChgCrd.Item(CStr(Check_no)).TipAmt, ChgCrd.Item(CStr(Check_no)).Media
       Trans_Log "SaveEMVPayment Update Complete"
       Set ChgCrd = Nothing
       Set GL = Nothing
       Set Sales = Nothing
       If IsDataCapInterface(EMVINTERFACE) And CapSignature Then   'If EMVINTERFACE = "DATACAP" And CapSignature And CapSignature Then
          Call GetSigCapture(Check_no)
       End If
     End If
End If
Sleep 250
Order.Timer2.Enabled = True
Exit Function
ErrorHandler:

Order.Timer2.Enabled = True
If Err.Number = 91 Then
  Error_Log 91, "Error Connecting To PinPad", "Out Of Scope", ""
  OutOfScope = "Error"
  Exit Function
Else
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Out of Scope ")
End If
   Select Case ErrorRespose
     Case "Abort"
       OutOfScope = "Error"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetCheckTotal(CheckNo As Long, CheckTotal As MyCheckTotal) As String
On Error Resume Next
Dim Net_Disc        As Currency
Dim Net_STax        As Currency
Dim Net_SPTax       As Currency
Dim Net_LTax        As Currency
Dim NetSale         As Currency
Dim CheckT          As Currency
Dim Db              As ADODB.Connection
Dim OP              As ADODB.Recordset
GetCheckTotal = "Failed"
ConnectADODB CheckStatDb, Db
  Sql = "Select * from [Open Checks] Where [Close Out Day] = 0 and [Check Number] = " & CheckNo
ConnectADORS Sql, Db, OP
    If Not OP.EOF Then
       Net_Disc = OP.fields("Gross Discounts").Value - OP.fields("Voided Discounts").Value
       Net_STax = OP.fields("Gross Sales Tax").Value - OP.fields("Voided Sales Tax").Value
       Net_SPTax = OP.fields("Gross Special Tax").Value - OP.fields("Voided Special Tax").Value
       Net_LTax = OP.fields("Gross Liquor Tax").Value - OP.fields("Voided Liquor Tax").Value
       NetSale = OP.fields("Gross Sales").Value - OP.fields("Voided Sales").Value
       CheckTotal.ChkTotal = (NetSale + Net_STax + Net_SPTax + Net_LTax) - Net_Disc
       CheckTotal.TaxTotal = ((Net_STax + Net_SPTax) + Net_LTax)
       GetCheckTotal = OP.fields("Table Number").Value & ""
    End If
  DbClose Db, OP
End Function
Public Sub ResetCardData()
With Emv_Back
     .r_AuthAmt = 0
     .r_AuthCode = ""
     .r_CardBalance = 0
     .r_CustomerName = ""
     .r_CashBack = 0
     .r_Last4 = ""
     .r_Message = ""
     .r_ReturnResult = ""
     .r_Status = ""
     .r_TipAmt = 0
End With
End Sub
Public Function AdJustAndCapture(ChargeAmount As Currency, Charge As ChargeTo, CheckInUse As Long) As String
    On Error GoTo ErrorHandler
    ResetCardData
    Dim Results             As String
    Dim GetAccount          As Boolean
    Dim AcctData()          As ver11DebitCards.AccountInfo
    Dim CreditFee           As Currency
    Dim FeeAmount           As Currency
    Dim MyDeposit           As Currency
    Dim OriginalAuth        As Currency
    Dim CardNo              As String
    Dim DoCapture           As Boolean
    Dim TPN                 As String
    Dim Answer              As String
    Dim CheckOwner          As String
    Dim NewAuthCode         As String
    Dim MyTipAtSale         As Currency
    
    DoCapture = False
    Dim DCAC As ver11DebitCards.Accounting
    Set DCAC = New ver11DebitCards.Accounting
    ReDim AcctData(0)
    AcctData(0).CheckNo = CheckInUse
    AcctData(0).ServerName = GetDeliveryDriver()
    AcctData(0).AuthCode = ""
    AcctData(0).CardNumber = ""
    
    GetAccount = DCAC.GetAccountChgData(AcctData(), SvrPath.TerminalPath)
    '// get credit card account information
    If GetAccount = False Then
       If Check_Status = MyOrderStatus(EditCheck) Then
         If Check_Option(156) = 1 Then
           If IsDataCapInterface(EMVINTERFACE) And INCREMENTAL > "" Then
              AdJustAndCapture = ""
              Exit Function
           End If
         End If
       End If
       '// exit if no card is attached
       
       Trans_Log "Credit Card Look Up Failed " & AcctData(0).ServerName & " Check " & AcctData(0).CheckNo
       AcctData(0).CheckNo = CheckInUse
       AcctData(0).ServerName = ""
       AcctData(0).AuthCode = ""
       AcctData(0).CardNumber = ""
       GetAccount = DCAC.GetAccountChgData(AcctData(), SvrPath.TerminalPath)
       If GetAccount = False Then
        See_Msg "Credit Card Not Found", 5, "Check No " & AcctData(0).CheckNo
        Trans_Log "Credit Card Not Found for Check No " & AcctData(0).CheckNo
        AdJustAndCapture = "Exit"
        Set DCAC = Nothing
       End If
    End If
    '// if not card is attached exit function above
    '// end of card lookup
    AcctData(0).ckTotal = AcctData(0).ckTotal + PhoneOrderPaidAmount(AcctData(0).CheckNo)
    CardNo = AcctData(0).CardNumber
    If IsDataCapInterface(EMVINTERFACE) And EMVSUPPORTPREAUTH = 1 Or IsPAXInterface(EMVINTERFACE) And PAXBARTABPREAUTH = 1 Then
      If ChargeAmount = AcctData(0).ckTotal Then
        If Check_Option(219) = 1 And OrderIs = "Delivery" Then
          '// disable credit card skip for delivery close
        Else
         If Check_Option(156) = 0 Then  '// Hold charge card for open tabs if 156 value is set to 1
          PrintCCSlip AcctData(0).CheckNo
         Else
           If Check_Status = MyOrderStatus(CloseCheck) Then
              PrintCCSlip AcctData(0).CheckNo
           End If
           '// print a credit card slip when record is closed
         End If
        End If
        Emv_Back.r_Status = "UPDATED"
        Set DCAC = Nothing
        Exit Function
      End If
      '//
      
      If ChargeAmount > 0 Then
         If Check_Option(151) = 1 And OrderIs = "Delivery" Then
           See_Msg "Collected Cash", 0, "Balance Due " & Format(ChargeAmount - AcctData(0).ckTotal, "0.00")
           Emv_Back.r_Status = "UPDATED"
           Set DCAC = Nothing
           Exit Function
         End If
         
         If Check_Option(242) = 1 Then
           If See_Msg("Use a different tender", 4, "Balance Due") = 6 Then
             Emv_Back.r_Status = "UPDATED"
             Set DCAC = Nothing
             Exit Function
           End If
         End If
         
         If AllOrders Is Nothing Then
           MyDeposit = 0
         Else
           MyDeposit = AllOrders.Item(1).Deposits
         End If  '// make sure deposits are set to 0
         
         If AcctData(0).ckTotal + MyDeposit = ChargeAmount Then
           '// only update if totals are different
         Else
          If OrderIs = "Delivery" Then
           See_Msg "Totals Updated", 5, ""
           UpdateDeliveryValues AcctData(0).CheckNo, False
          Else
           If IsDataCapInterface(EMVINTERFACE) And INCREMENTAL > "" Then
           
           Else
             UpdateDeliveryValues AcctData(0).CheckNo, True
           End If
          End If
         End If
         '// end section with totals
         If IsDataCapInterface(EMVINTERFACE) And INCREMENTAL > "" Then
          If Check_Status = MyOrderStatus(EditCheck) Then
           If ChargeAmount <= AcctData(0).ckTotal Then
            Emv_Back.r_Status = "UPDATED"
            Set DCAC = Nothing
            Exit Function
           End If
          End If
          '// make sure incremental auth gets updated at check close
         Else
          Emv_Back.r_Status = "UPDATED"
          Set DCAC = Nothing
          Exit Function
         End If
      End If
      '// if charge amount > 0
    End If
    '// above is block for datacap and PAX with preauth
    
    See_Msg "Sending for Adjustment", 5, ""
    With Emv_Back
       .r_AuthAmt = 0
       .r_CardBalance = 0
       .r_Message = "ABORTED"
       .r_Status = "ABORTED"
    End With
    Dim Sales As AllEMVSales
    Set Sales = New AllEMVSales
    If Sales.Item(CStr(AcctData(0).CheckNo)) Is Nothing Then
       Sales.add CStr(AcctData(0).CheckNo)
    End If
    
    Dim DeviceIp As String
    DeviceIp = "": ' Answer = ""
     If IsPAXInterface(EMVINTERFACE) Then
      If AcctData(0).RegisterName <> MyRegisterName Then
        DeviceIp = GetPaxIpAddress(AcctData(0).RegisterName)
        If DeviceIp = "" Then
          See_Msg "Cannot apply tip at this work station", 0, "Must Use " & AcctData(0).RegisterName
          Exit Function
        End If
      End If
     End If
     '// pax with not no preauth
     Sales.Item(CStr(AcctData(0).CheckNo)).OrgAuthCode = AcctData(0).AuthCode   '// old var being used as auth code
     Sales.Item(CStr(AcctData(0).CheckNo)).Cashier = AcctData(0).ServerName    'LoggedOnAs(0).EMPName
     Sales.Item(CStr(AcctData(0).CheckNo)).Term = AcctData(0).RegisterName
     Sales.Item(CStr(AcctData(0).CheckNo)).IpAddress = DeviceIp
     Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(ChargeAmount, "0.00")
     Sales.Item(CStr(AcctData(0).CheckNo)).TipAmt = AcctData(0).TipAmt
     Sales.Item(CStr(AcctData(0).CheckNo)).XMLString = AcctData(0).XMLDataString
     Sales.Item(CStr(AcctData(0).CheckNo)).TranType = "Credit"
     Sales.Item(CStr(AcctData(0).CheckNo)).TransNo = AcctData(0).CheckNo
     MyTipAtSale = AcctData(0).TipAmt
     
     Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
     If LoadEMVDevice(AcctData(0).RegisterName, EmvEquip) = "Failed" Then
        See_Msg "No Device Found", 5, "Check Device Setup"
        Exit Function
     End If
     Dim PMG As PAXMSG
     Set PMG = New PAXMSG
     Dim PXFO As EMVProcessLib.Process
     Set PXFO = New Process
     Dim VMG As PAXMSG
     Set VMG = New PAXMSG
     If IsDejavooInterface(EMVINTERFACE) Then
      If EmvEquip.Item(1).MaxQtyOffline = 1 Then 'used for regisiter with processor for processing fees
       If ProcessingFeesOn Then
         CreditFee = GetCurrentProcessingFee
         Sales.Item(CStr(AcctData(0).CheckNo)).PassThru = CCur(Round(CreditFee, 2)) * 100
         Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(CCur(Round(ChargeAmount, 2)) - CCur(Round(CreditFee, 2)), "0.00")
       End If
      End If
     End If
    Dim xml As MPxmlParse
    Set xml = New MPxmlParse
     xml.xml = AcctData(0).XMLDataString
     OriginalAuth = Val(xml.token("OriginalAuthAmt"))
     TPN = xml.token("TPN")
    Set xml = Nothing
    If IsDejavooInterface(EMVINTERFACE) Then
      If TPN > "" Then
       EmvEquip.Item(1).deviceId = TPN
       EmvEquip.Item(1).MerchantId = TPN
      Else
       TPN = ""
      End If
    End If
     
     If IsDataCapInterface(EMVINTERFACE) And INCREMENTAL > "" Or IsDataCapInterface(EMVINTERFACE) And Check_Option(140) = 1 And Check_Option(156) = 1 Then
        Results = PXFO.ProcessSale(Sales, DevSaleType_IncrementalAuthByRecordNo, PMG, EmvEquip)
        AdJustAndCapture = PMG.Item(1).ResultTxt
        If IsApproved(PMG.Item(1).ResultTxt) Then
           Sales.Item(CStr(AcctData(0).CheckNo)).XMLString = PMG.Item(1).ExtraData
           If EMVSUPPORTPREAUTH = 0 Then Results = PXFO.ProcessSale(Sales, DevSaleType_PreAuthCaptureByRecordNo, PMG, EmvEquip)
        End If
     Else
      If IsPAXInterface(EMVINTERFACE) Then
         Results = PXFO.ProcessSale(Sales, DevSaleType_PPostAuth, PMG, EmvEquip)
      ElseIf IsDejavooInterface(EMVINTERFACE) Then
      
       If OriginalAuth = ChargeAmount Then
         DoCapture = True
         Results = PXFO.ProcessSale(Sales, DevSaleType_djvCapture, PMG, EmvEquip)
         Trans_Log "Results from Capture from Orig Code " & Results
       Else
         Trans_Log "Sending Token for a new authorization. New Check Total " & Format(ChargeAmount, "0.00")
         Results = PXFO.ProcessSale(Sales, DevSaleType_djvUseToken, PMG, EmvEquip)
         Trans_Log "Results from sending token " & Results & " Result Txt " & PMG.Item(1).ResultTxt
         If UCase(PMG.Item(1).ResultTxt) = "FAILED" Then
           See_Msg "Token Sale Failed", 0, ""
           Trans_Log "Send Token Failed " & Results
           Exit Function
         End If
         If UCase(PMG.Item(1).ResultTxt) = "DECLINED" Then
             GetAccount = DCAC.GetAccountChgData(AcctData(), SvrPath.TerminalPath)
             Dim VEMV As AllEMVSales
             Set VEMV = New AllEMVSales
             VEMV.add AcctData(0).CheckNo
             VEMV.Item(CStr(AcctData(0).CheckNo)).OrgAuthCode = AcctData(0).AuthCode   '// old var being used as auth code
             VEMV.Item(CStr(AcctData(0).CheckNo)).Cashier = AcctData(0).ServerName    'LoggedOnAs(0).EMPName
             VEMV.Item(CStr(AcctData(0).CheckNo)).Term = AcctData(0).RegisterName
             VEMV.Item(CStr(AcctData(0).CheckNo)).IpAddress = DeviceIp
             VEMV.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(ChargeAmount, "0.00")
             VEMV.Item(CStr(AcctData(0).CheckNo)).TipAmt = 0
             VEMV.Item(CStr(AcctData(0).CheckNo)).XMLString = AcctData(0).XMLDataString
             VEMV.Item(CStr(AcctData(0).CheckNo)).TranType = "Credit"
             VEMV.Item(CStr(AcctData(0).CheckNo)).TransNo = AcctData(0).CheckNo
             Results = PXFO.ProcessSale(VEMV, DevSaleType_djvCapture, PMG, EmvEquip)
         ElseIf UCase(PMG.Item(1).ResultTxt) = "APPROVED" Then
             NewAuthCode = PMG.Item(1).AuthCode
             Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = OriginalAuth
             If ProcessingFeesOn Then
                Answer = PXFO.ProcessSale(Sales, DevSaleType_djvCapture, VMG, EmvEquip)
                Trans_Log "Void Capture " & Answer
             End If
             Answer = PXFO.ProcessSale(Sales, DevSaleType_djvVoid, VMG, EmvEquip)
             Trans_Log "Void Sale " & Answer
             RemoveCreditPayment CheckInUse
         End If
       End If
      End If
     End If

     If Results = "CONNECT ERROR" Then
        See_Msg Results, 5, UCase(PMG.Item(1).ResultTxt)
        Exit Function
     End If
     If IsPAXInterface(EMVINTERFACE) And PMG.Item(1).Message = "NEED REMOVE CARD FIRST" Then
       See_Msg PMG.Item(1).Message, 0, "Remove the card and click OK"
       Results = PXFO.ProcessSale(Sales, DevSaleType_PPostAuth, PMG, EmvEquip)
     End If
     If PMG.Item(1).ResultTxt = "DECLINED" Then
       If OriginalAuth = AcctData(0).ckTotal Then
         See_Msg "Charge Declined", 0, "New Payment Required"
         See_Msg "Old Charge will now be voided", 0, ""
         Delete_Payment_Record AcctData(0).CheckNo, AcctData(0).ServerName, True
         Emv_Back.r_AuthAmt = AcctData(0).ckTotal * -1
         '// if the originalauth
      End If
     End If
     Emv_Back.r_Status = UCase(PMG.Item(1).ResultTxt)
     
     If IsDejavooInterface(EMVINTERFACE) Then
      Dim xmls As MPxmlParse
      Set xmls = New MPxmlParse
      xmls.xml = PMG.Item(1).ExtraData
      FeeAmount = Format(Val(xmls.token("FeeAmount")), "0.00")
      Emv_Back.r_AuthAmt = Format(CCur(Val(PMG.Item(1).ApprovedAmount) + FeeAmount), "0.00")
     Else
      Emv_Back.r_AuthAmt = Format(CCur(Val(PMG.Item(1).ApprovedAmount)), "0.00")
     End If
     
     If NewAuthCode = "" Then
      Emv_Back.r_AuthCode = Sales.Item(CStr(AcctData(0).CheckNo)).OrgAuthCode
      Trans_Log "Using Original Auth Code " & Emv_Back.r_AuthCode & " Check No " & AcctData(0).CheckNo
     Else
      Emv_Back.r_AuthCode = NewAuthCode
      Trans_Log "Using New Auth Code " & Emv_Back.r_AuthCode & " Check No " & AcctData(0).CheckNo
     End If
     Emv_Back.r_Cardtype = AcctData(0).AccountName  ' stop name returning as undeterminded
     Emv_Back.r_CardBalance = Val(PMG.Item(1).RemainingBalance)
     Emv_Back.r_TipAmt = Val(PMG.Item(1).TipAmount)
     
     Emv_Back.r_Last4 = Trim$(PMG.Item(1).Last4)
     If Emv_Back.r_Last4 = "" Then Emv_Back.r_Last4 = CardNo
     If TPN > "" Then
       PMG.Item(1).ExtraData = PMG.Item(1).ExtraData & "<TPN>" & TPN & "</TPN>"
     End If
     If Val(PMG.Item(1).RefNum) > 0 Then
       Emv_Back.r_ReturnResult = PMG.Item(1).ExtraData & "<Id>" & PMG.Item(1).RefNum & "</Id>"
     Else
       Emv_Back.r_ReturnResult = PMG.Item(1).ExtraData
     End If
     Emv_Back.r_Message = PMG.Item(1).Message
     Emv_Back.r_CashBack = Val(PMG.Item(1).CashBack & "")
     If Emv_Back.r_Message = "" Then Emv_Back.r_Message = "Canceled"
     Trans_Log Emv_Back.r_Status & " " & PMG.Item(1).Message & " " & PMG.Item(1).ResultTxt
     '// display the message returned from processing
     If Emv_Back.r_Message = "OK" And Emv_Back.r_AuthAmt > 0 Then
        Emv_Back.r_Status = "APPROVAL"
     End If
    ' Emv_Back.r_Server = ReturnMyCashier
     Emv_Back.r_Server = AcctData(0).ServerName
     If IsApproved(Emv_Back.r_Status) Then
      If Emv_Back.r_Message = "PARTIAL AP" Then
          RollAuthorizationBack AcctData(0).CheckNo, OriginalAuth
      Else
          UpdateDeliveryValues AcctData(0).CheckNo, False
       '// only update phone orders
       'changed for datacap -- acc ref data changes
      End If
     '// start new code here
      Dim GL As GeneralLib.Utilities
      Set GL = New GeneralLib.Utilities
      Dim ChgCrd As GeneralLib.EMVSales
      Set ChgCrd = New GeneralLib.EMVSales
       If ChgCrd.Item(CStr(AcctData(0).CheckNo)) Is Nothing Then
         ChgCrd.add CStr(AcctData(0).CheckNo)
       End If
       With ChgCrd
          .Item(CStr(AcctData(0).CheckNo)).Cashier = ReturnMyCashier
          .Item(CStr(AcctData(0).CheckNo)).Term = MyRegisterName
          .Item(CStr(AcctData(0).CheckNo)).Media = Emv_Back.r_Cardtype
          .Item(CStr(AcctData(0).CheckNo)).AccountNo = Emv_Back.r_AuthCode
          If IsPAXInterface(EMVINTERFACE) Or IsDejavooInterface(EMVINTERFACE) Then
           .Item(CStr(AcctData(0).CheckNo)).SaleType = "Capture"
          Else
           .Item(CStr(AcctData(0).CheckNo)).SaleType = "EMV"
          End If
          If Emv_Back.r_SaleType = "Debit" Then
            .Item(CStr(AcctData(0).CheckNo)).ChargeType = ChargeTypeEnum.ctDebit
            .Item(CStr(AcctData(0).CheckNo)).Media = "Debit"
          ElseIf Emv_Back.r_SaleType = "EBT_FOODSTAMP" Then
           .Item(CStr(AcctData(0).CheckNo)).Media = "EBT"
           .Item(CStr(AcctData(0).CheckNo)).ChargeType = ChargeTypeEnum.ctEBT '// need a 2 to make sure we can void
          Else
            If EMVSUPPORTPREAUTH = 1 Then
              .Item(CStr(AcctData(0).CheckNo)).ChargeType = ChargeTypeEnum.ctPreAuth
            Else
              .Item(CStr(AcctData(0).CheckNo)).ChargeType = ChargeTypeEnum.ctSale
            End If
          End If
          .Item(CStr(AcctData(0).CheckNo)).SalesAmt = Emv_Back.r_AuthAmt
          If IsDejavooInterface(EMVINTERFACE) Or IsDataCapInterface(EMVINTERFACE) Then
            .Item(CStr(AcctData(0).CheckNo)).CheckTotal = Emv_Back.r_AuthAmt
          Else
           .Item(CStr(AcctData(0).CheckNo)).CheckTotal = Emv_Back.r_AuthAmt - Emv_Back.r_TipAmt
          End If
          .Item(CStr(AcctData(0).CheckNo)).TransNo = CStr(AcctData(0).CheckNo)
           CheckOwner = GetWhoOwnsTheCheck(LoggedOnAs(0).EMPName, AcctData(0).CheckNo)
            '// who owns the check
          .Item(CStr(AcctData(0).CheckNo)).ServerName = CheckOwner
          .Item(CStr(AcctData(0).CheckNo)).TipAmt = MyTipAtSale
          Emv_Back.r_TipAmt = MyTipAtSale
          .Item(CStr(AcctData(0).CheckNo)).XMLString = Emv_Back.r_ReturnResult
          .Item(CStr(AcctData(0).CheckNo)).CardNo = Emv_Back.r_Last4
          .Item(CStr(AcctData(0).CheckNo)).TimeOfSale = GetOriginalTimeOfSale(AcctData(0).CheckNo, .Item(CStr(AcctData(0).CheckNo)).ServerName, Emv_Back.r_Last4)
          If Not allcards Is Nothing Then
           If allcards.Item("-9999") Is Nothing Then
            .Item(CStr(AcctData(0).CheckNo)).gateWayIdNo = ""
           Else
           .Item(CStr(AcctData(0).CheckNo)).gateWayIdNo = allcards.Item("-9999").AccountOnFile
           allcards.Item("-9999").AccountOnFile = ""
           End If
          End If
          If IsDejavooInterface(EMVINTERFACE) Then
           If DoCapture Then  '// capture records is no changes were made
             MarkRecordCaptured AcctData(0).CheckNo
           End If
          End If
          '// used to save future phone order check number
         If IsDataCapInterface(EMVINTERFACE) And INCREMENTAL > "" Then
           If Check_Option(156) = 1 Then
              If Check_Status = MyOrderStatus(EditCheck) Then
                UpdateCreditCardData Emv_Back.r_ReturnResult, Emv_Back.r_AuthAmt, Emv_Back.r_TipAmt, Emv_Back.r_AuthCode, AcctData(0).CheckNo, Emv_Back.r_Last4, .Item(CStr(AcctData(0).CheckNo)).ServerName
                Exit Function
              ElseIf Check_Status = MyOrderStatus(CloseCheck) Then
                UpdateAuthCode Emv_Back.r_AuthCode, AcctData(0).CheckNo, .Item(CStr(AcctData(0).CheckNo)).ServerName, Emv_Back.r_Last4
              End If
           End If
         End If
         SaveEMVPayments .Item(CStr(AcctData(0).CheckNo)).TimeOfSale, AcctData(0).CheckNo, .Item(CStr(AcctData(0).CheckNo)).ServerName, .Item(CStr(AcctData(0).CheckNo)).CheckTotal, .Item(CStr(AcctData(0).CheckNo)).TipAmt, .Item(CStr(AcctData(0).CheckNo)).Media
         'Set apychg = New ver11DebitCards.Accounting
         Results = GL.ChargeAccount(ChgCrd, SvrPath.TerminalPath)
         If Results <> "Success" Then
          See_Msg "Failure to write to debit database", 0, Results
          Trans_Log "Failure to write to debit database"
         End If
         Set GL = Nothing
         Close_Open_Check AcctData(0).CheckNo, CheckOwner
       End With
     '// end herr
     Else
       If InStr(UCase(Emv_Back.r_Status), "DECLINE") > 0 Then
          RollAuthorizationBack AcctData(0).CheckNo, OriginalAuth
          If Emv_Back.r_AuthAmt > 0 Then Emv_Back.r_AuthAmt = 0
       End If

       If PMG.Item(1).Message = "NOT FOUND" Then
         See_Msg "Transaction Already Captured", 5, ""
       Else
         See_Msg PMG.Item(1).ResultCode, 5, PMG.Item(1).Message & " " & PMG.Item(1).ResultTxt
       End If
       Trans_Log PMG.Item(1).Message & " " & PMG.Item(1).ResultTxt
     End If
Set PMG = Nothing
Set VMG = Nothing
Set PXFO = Nothing
Set Sales = Nothing
Set DCAC = Nothing
'Set apychg = Nothing
Set EmvEquip = Nothing
Exit Function
ErrorHandler:
If Err.Number = 91 Then
  Error_Log 91, "Error Connecting To PinPad", "Out Of Scope", ""
  If See_Msg("Failed To Connect To PinPad", 4, "Yes to Retry Or No To Quit") = 6 Then
    Resume
  End If
  Exit Function
Else
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Out of Scope ")
End If
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function UpdateAuthCode(AuthCode As String, Check As Long, SvrName As String, Last4 As String) As String
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim OriginalAmt  As Currency
ConnectADODB DebitDb, Db
   Sql = "Update [Debit] Set [Account Number] = '" & AuthCode & "' Where  [Check Number] = " & Check & " and [Server Name] = '" & SvrName & "'"
   If Last4 > "" Then Sql = Sql & " and [Card Number] = '" & Last4 & "'"   'Credit Card Number last 4
   Sql = Sql & " and [Close Out Day] = 0"
   Db.Execute Sql
   Trans_Log "Update AuthCode " & AuthCode & "Update Incremental Account " & AuthCode
DbClose Db, RS
Exit Function
ErrorHandler:
Select Case Main_Error_Handler(Err.Description, Err.Number, "Update Incremental Auth")
       Case Is = 3
        Set Db = Nothing
       Case Is = 4
        Resume
       Case Is = 5
        Resume Next
End Select
End Function
Public Function GetOriginalTimeOfSale(Check As Long, SvrName As String, Last4 As String) As String
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
   GetOriginalTimeOfSale = ""
   ConnectADODB DebitDb, Db
   Sql = "Select [Time] from [Debit] Where  [Check Number] = " & Check & " and [Server Name] = '" & SvrName & "'"
   Sql = Sql & " and [Card Number] = '" & Last4 & "'"   'Credit Card Number last 4
   Sql = Sql & " and [Close Out Day] = 0"
   ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
      GetOriginalTimeOfSale = RS.fields(0).Value & ""
   Else
      GetOriginalTimeOfSale = Format(Time, "hh:mm:ss ampm")
   End If
DbClose Db, RS
Exit Function
ErrorHandler:
Select Case Main_Error_Handler(Err.Description, Err.Number, "Update Incremental Auth")
       Case Is = 3
        Set Db = Nothing
       Case Is = 4
        Resume
       Case Is = 5
        Resume Next
End Select
End Function

'Public Function IsHouseAccountInUse() As Boolean
'Dim Db              As ADODB.Connection
'Dim RS              As ADODB.Recordset
'IsHouseAccountInUse = False
'If ConnectADODB(HouseAcct, Db) = False Then
'  Exit Sub
'End If
'If ConnectADORS("Select * from [General Data]", Db, RS) = False Then Exit Sub
'If RS.EOF Then
'  IsHouseAccountInUse = False
'Else
'  IsHouseAccountInUse = True
'End If
'DbClose Db, RS
'Check to see if house accounts are in use.
'will not prompt for search for giftcard balance
'End Function

Public Function UpdateCreditCardData(xml As String, ChargeTotal As Currency, tip As Currency, Auth As String, Check As Long, Last4 As String, SvrName As String) As String
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim OriginalAmt  As Currency
UpdateCreditCardData = ""
ConnectADODB DebitDb, Db
   Sql = "Select * from [Debit] Where [Check Number] = " & Check & " and [Server Name] = '" & SvrName & "'"
   If Last4 > "" Then Sql = Sql & " and [Card Number] = '" & Last4 & "'"   'Auth Number
   Sql = Sql & " and [Close Out Day] = 0"
   If ConnectADORS(Sql, Db, RS) = False Then
     See_Msg "Error Connecting to Debit Database", 5, "Record No Updated..."
     Exit Function
   End If
   If Not RS.EOF Then
    RS.fields("Account Number").Value = Auth  '// auth code
    RS![Amount] = Format(ChargeTotal, "0.00")
    RS![tip] = Format(tip, "0.00")
    OriginalAmt = Val(RS![Note] & "")
    If xml > "" Then RS![XMLString] = xml & "<OriginalAuthAmt>" & Format(OriginalAmt, "0.00") & "</OriginalAuthAmt>"
   RS.Update
  End If
  Trans_Log "EMV Sale " & tip & " " & xml & " " & "Update Incremental Account"
  UpdateCreditCardData = "Success"
DbClose Db, RS
Exit Function
ErrorHandler:
Select Case Main_Error_Handler(Err.Description, Err.Number, "Update Incremental Auth")
       Case Is = 3
        Set Db = Nothing
       Case Is = 4
        Resume
       Case Is = 5
        Resume Next
End Select

End Function
Public Sub SaveReOpenEMVPayment(CheckNo As Long)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB DebitDb, Db
        Sql = "Select * from [Debit] Where [Close Out day] = 0 and [Check Number] = " & CheckNo
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  SaveEMVPayments RS.fields("Time").Value, CheckNo, RS.fields("Server Name").Value, RS.fields("Amount").Value, RS.fields("Tip").Value, RS.fields("Account Name").Value
End If
DbClose Db, RS
End Sub
Public Sub SaveEMVPayments(TimeOfPayment As String, CheckNo As Long, SrvName As String, ChargeAmt As Currency, TipAmt As Currency, MediaType As String)
On Error GoTo ERH
    Dim MB As String
    Dim GL As GeneralLib.Utilities
    Set GL = New GeneralLib.Utilities
    MB = GL.SaveEMVPayments(TimeOfPayment, CStr(CheckNo), SrvName, ChargeAmt, TipAmt, MediaType)
    Set GL = Nothing
Exit Sub
ERH:
Error_Log Err.Number, Err.Description, "Apply Tip EMV", "None"
Err.Clear
End Sub
Public Function ReturnDepartmentName(ByVal ModDepartmentName As String, ByVal MenuDepartmentName As String) As String
    Dim Db As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim RS As ADODB.Recordset
    Dim sSql As String

    On Error GoTo ErrorHandler
    ' Establish connection
    If Not ConnectADODB("MSaleData", Db) Then
        ReturnDepartmentName = MenuDepartmentName
        Exit Function
    End If

    ' Prepare the SQL command and use parameters to avoid SQL injection
    sSql = "SELECT [Report Name] FROM [Report Categories] WHERE [Report Name] = ?"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = Db
        .CommandText = sSql
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("param1", adVarChar, adParamInput, 255, ModDepartmentName)
    End With

    ' Execute the command
    Set RS = cmd.Execute

    ' Check if data is found
    If Not RS.EOF Then
        ReturnDepartmentName = ModDepartmentName
    Else
        ReturnDepartmentName = MenuDepartmentName
    End If

CleanUp:
    ' Clean up database objects
    If Not RS Is Nothing Then
        If RS.State = adStateOpen Then RS.Close
        Set RS = Nothing
    End If
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not Db Is Nothing Then
        If Db.State = adStateOpen Then Db.Close
        Set Db = Nothing
    End If
    Exit Function

ErrorHandler:
    ' Handle errors
    ReturnDepartmentName = MenuDepartmentName
    Resume CleanUp
End Function

Public Function ProcessEMVCharge(ChargeAmt As Currency, CheckInUse As Long) As DeliveryPayment
 'On Error GoTo errorhandler
 ProcessEMVCharge.QuitTransaction = "Exit"
 ProcessEMVCharge.ApprovalAmount = 0
 Dim CardSlip As String
   Dim OrderTaker As String
   OrderTaker = LoggedOnAs(0).EMPName
   If OutOfScope(OrderTaker, "EMV_Sale", Format(ChargeAmt, "0.00"), OOCCreditSale) = "Exit" Then Exit Function
   
    ProcessEMVCharge.Approval = Emv_Back.r_Status
    If IsApproved(Emv_Back.r_Status) Then
       ProcessEMVCharge.Approval = "APPROVAL"
       If Emv_Back.r_AuthAmt < ChargeAmt Then
        If GetConfigInfo("Phone Orders", "Stop With Insufficient Funds", "Yes", "FUNCTIONS.INI") = "ENABLED" Then
          See_Msg "Partial Authorization Voiding Charge", 0, "Insufficient Funds"
          Delete_Payment_Record Check_no, LoggedOnAs(0).EMPName, True
          ProcessEMVCharge.ApprovalAmount = 0
          ProcessEMVCharge.Approval = "Declined"
          Exit Function
        End If
        See_Msg "Partial Approval", 5, "Balance Due " & Format(ChargeAmt - Emv_Back.r_AuthAmt, "0.00")
        ProcessEMVCharge.ApprovalAmount = Emv_Back.r_AuthAmt
       Else
          See_Msg "Approved...", 5, ""
          ProcessEMVCharge.ApprovalAmount = ChargeAmt
          ProcessEMVCharge.Approval = "APPROVAL"
       End If
       If Check_Option(253) = 0 Or OrderIs = "Pick Up" Then
         If IsDataCapInterface(EMVINTERFACE) Then
          CardSlip = FormatEMVReceipt(Emv_Back.r_ReturnResult)
          If CardSlip = "" Then GoTo MSaleSlip
          PrintChitSlip CardSlip
         Else
MSaleSlip:
          Dim CCD As CCData
          With CCD
            .HeaderData = "No"
            .HeaderReason = "EMV-"
            .AuthNo = Emv_Back.r_AuthCode
            .SaleAmt = Emv_Back.r_AuthAmt '- Emv_Back.r_TipAmt
            .TipAmt = Emv_Back.r_TipAmt
            .AccNo = Emv_Back.r_Last4
            .CustName = ""
            .CType = Emv_Back.r_Cardtype
            .ServerName = Emv_Back.r_Server
            .TBLName = Table_Number
            .ChkNO = CStr(CheckInUse)
            .ProcessorResponse = "Approved"
            .Datetime = Emv_Back.r_ReturnResult
            .ckTotal = Format(Chk_Sales_data(5), "0.00")
          End With
          PrintSlip CCD
         End If
       End If
    ElseIf Emv_Back.r_Status = "DECLINED" Then
            See_Msg "Declined...", 0, Emv_Back.r_Message
            If IsDataCapInterface(EMVINTERFACE) Then
               PrintChitSlip FormatEMVReceipt(Emv_Back.r_ReturnResult)
            End If
          Exit Function
    Else
        See_Msg Emv_Back.r_Message, 0, ""
        Exit Function
    End If
    If Check_Option(0) = 0 Then Emv_Back.r_TipAmt = 0
    '// will drop tip on consolidated credit card slip
    allcards.Item(CStr(CheckInUse)).ChargeAmt = Format(Val(Emv_Back.r_AuthAmt), "0.00")    '  ChargeAmount + Format(Val(xmls.token("CashBack")), "0.00") ' Emv_Trans.CashBack  'Val(xmls.token("CashBack"))
    Tenders(0).TAmt = Format(Val(Emv_Back.r_AuthAmt), "0.00")
    Tenders(0).TType = TenderTypes.EVM
    Save_Payments Format(Val(Emv_Back.r_AuthAmt), "0.00"), CheckInUse
End Function
Public Function IsBetween(CurrentTimeIn As String, CurrentTimeOut As String, StartTime As String, EndTime As String) As Boolean
 IsBetween = False
 If IsDate(CurrentTimeIn) And IsDate(CurrentTimeOut) Then
    If CDate(CurrentTimeIn) >= CDate(StartTime) Then
         If CDate(CurrentTimeIn) < CDate(EndTime) Then
           IsBetween = True
           Exit Function
         End If
    End If
    If CDate(CurrentTimeIn) < CDate(EndTime) Then
         If CDate(CurrentTimeIn) >= CDate(StartTime) Then
           IsBetween = True
           Exit Function
         End If
    End If
    
    If CDate(CurrentTimeOut) <= CDate(EndTime) Then
         If CDate(CurrentTimeOut) > CDate(StartTime) Then
           IsBetween = True
           Exit Function
         End If
    End If
    If CDate(CurrentTimeOut) > CDate(StartTime) Then
         If CDate(CurrentTimeOut) <= CDate(EndTime) Then
           IsBetween = True
           Exit Function
         End If
    End If
 End If
End Function
Public Function RoundUpMileage(ByVal ValueIn As Double) As Integer
Dim Temp As Double
Temp = Int(ValueIn)
If Temp <> ValueIn Then
   Temp = Temp + 1
End If
RoundUpMileage = Temp
End Function
Public Function GetZoneName(DrvMiles As String, DrvTime As String, DrvResult As DrvZoneData) As String
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim GetDist() As String
Dim Travel As Single
ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Db
  Sql = "Select * from [Delivery Distance]"
ConnectADORS Sql, Db, Tb
If Not Tb.EOF Then
 If Tb.fields("Name").Value & "" = "Miles" Then
   GetDist = Split(DrvTime, " ")
   Travel = Val(GetDist(1))
 Else
   GetDist = Split(DrvMiles, " ")
   Travel = Val(GetDist(1))
 End If
 DrvResult.ZoneDistance = Travel
   Sql = "Select * from [Delivery Zones] Where " & RoundUpMileage(Travel) & " Between [Start] and [End]"
 ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
   DrvResult.ZoneName = Tb.fields("Name").Value
  Else
   DrvResult.ZoneName = "Out of Area"
  End If
End If
DbClose Db, Tb
End Function
Public Function LookUpZone(DrvDist As Single) As String
DrvDist = RoundUpMileage(DrvDist)
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
LookUpZone = GL.LookUpZone(DrvDist)
Set GL = Nothing
End Function
Public Function GetTimeRecordId(EmployeeName As String) As ReturnTimeValues
On Error GoTo ERH
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim Hoursworked As Single
Dim Total_Time  As Single
OpenDb TimeRecordsDb, Db
    Sql = "Select * From [Employee Records] Where [Employee Name] = '" & EmployeeName & "'"
    Sql = Sql & " AND [Time In] > '' and [Time Out] = ''"
    Sql = Sql & " Order By TRY_CAST([Shift Date] AS date) desc"
ConnectADORS Sql, Db, RS
    If RS.EOF Then
      GetTimeRecordId.Results = "No Open Time Records"
    End If
    If RecordsCount(RS) > 1 Then
      GetTimeRecordId.Results = "Please correct time punches"
    End If
    If Not RS.EOF Then
     Hoursworked = CalculateTimeRecord(RS![Shift date] & "", Date, RS![Time In] & "", RS![Break Out] & "", RS![Break In] & "", Format(Time, "HH:MM AmPm"), Total_Time)
     If (Total_Time / 60) > 16 Then
       See_Msg "Invalid Punch Time", 5, "Have a manager correct your time"
       GetTimeRecordId.Results = "Please correct time punches"
       GetTimeRecordId.RecordId = ""
     Else
      GetTimeRecordId.Results = "Success"
      GetTimeRecordId.RecordId = RS.fields("RecordId").Value & ""
      GetTimeRecordId.ShiftDate = CStr(RS.fields("Shift Date").Value)
     End If
    End If
DbClose Db, RS
Exit Function
ERH:
Err.Clear
End Function
Public Function IHaveOverTime(SearchName As String, CloseDate As String) As Boolean
On Error GoTo ERH
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim DayNo       As Integer
Dim Total_Time  As Single
Dim TimeWorked  As Single
Dim Hoursworked As String
IHaveOverTime = False
TimeWorked = 0
DayNo = GetDateToNo(CloseDate)
OpenDb TimeRecordsDb, Db
    Sql = "Select * From [Employee Records] Where [Employee Name] = '" & SearchName & "'"
    Sql = Sql & " AND TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(CloseDate) - (DayNo - 1) & " " & GetStartTime & "'"
    Sql = Sql & " and '" & SetEndingDate(CDate(CloseDate)) & " " & SetEndingTime & "'"
ConnectADORS Sql, Db, RS
    Do While Not RS.EOF
        Hoursworked = CalculateTimeRecord(RS![Shift date] & "", RS![Punch out Date] & "", _
                            RS![Time In] & "", RS![Break Out] & "", RS![Break In] & "", RS![Time Out] & "", Total_Time)
        TimeWorked = TimeWorked + Val(Hoursworked)
      RS.MoveNext
    Loop
DbClose Db, RS
If TimeWorked > Val(Input_Option(49)) Then
   IHaveOverTime = True
Else
   IHaveOverTime = False
End If
Exit Function
ERH:
Err.Clear
End Function
'//
Public Function CalculateTimeRecord(Punch_InDate As String, Punch_OutDate As String, Time_On As String, Brk_Out As String, Brk_In As String, Time_Out As String, TotalTime As Single) As String
On Error Resume Next
Dim TotalWorked As Long
Dim TotalBreak As Long
Dim TotalHours As Long
Dim TotalMins As Long
Dim RegTime As String
Dim PaidBreak As Long
Dim InDate As Date
Dim OutDate As Date
Dim InTime As Date
Dim OutTime As Date
Dim InBrk As Date
Dim OutBrk As Date
Dim ManditoryBreak As Long
'Time_On , Brk_Out , Brk_In  Time_Out  TotalTime
'Incomming var from database

CalculateTimeRecord = "": TotalTime = 0  'Makes sure var are Nothing
If Time_On = "00:00" Or Time_Out = "00:00" Or Brk_Out = "00:00" Or Brk_In = "00:00" Then Exit Function
'exits if time record was deleted

If Trim(Punch_InDate) = "" Then Punch_InDate = CStr(Date)
InDate = CDate(Punch_InDate)
If Err Then
 InDate = Format(Date, "MM/DD/YYYY")
 Err.Clear
End If

If Trim(Punch_OutDate) = "" Then Punch_OutDate = CStr(Date)
OutDate = CDate(Punch_OutDate)
If Err Then
 OutDate = Format(Date, "MM/DD/YYYY")
 Err.Clear
End If
'Verifies that punch dates are valid
InTime = CDate(Time_On) + InDate
If Err Then
   Exit Function
End If
If Trim(Time_Out) = "" Then Time_Out = CStr(Time)
OutTime = OutDate + CDate(Time_Out)
If DateDiff("n", InTime, OutTime) < 0 Then
  OutTime = CDate(OutDate) + 1 + CDate(Time_Out)
End If

If Err Then
   OutTime = Date + Time
   Err.Clear
End If
PaidBreak = Val(Input_Option(4)) '//  Option (4) = paid break
ManditoryBreak = Val(Input_Option(7)) '//  Option (7) = manditorybreak out
TotalBreak = 0  'set break to 0
If Trim(Brk_In) > "" And Trim(Brk_Out) > "" Then
    If Format(Brk_Out, "MM/DD/YYYY") < CDate(InDate) Then
       Brk_Out = Brk_Out + InDate
    End If
    
    OutBrk = CDate(Brk_Out)
    If Err Then
     TotalBreak = 0
     GoTo NoBreak
     Err.Clear
    End If
    InBrk = CDate(Brk_In)
    If Format(Brk_In, "MM/DD/YYYY") < CDate(InDate) Then
       Brk_In = Brk_In + InDate
    End If
    
    If Err Then
      TotalBreak = 0
      GoTo NoBreak
      Err.Clear
    End If
    TotalBreak = DateDiff("n", Brk_Out, Brk_In, vbMonday, vbFirstJan1)
    If TotalBreak < 0 Then
      Brk_In = DateAdd("d", 1, Brk_In)
      TotalBreak = DateDiff("n", Brk_Out, Brk_In, vbMonday, vbFirstJan1)
    End If
    If PaidBreak > 0 And TotalBreak <= PaidBreak Then
      TotalBreak = 0
    Else
      TotalBreak = TotalBreak - PaidBreak
    End If
End If
NoBreak:
TotalWorked = DateDiff("n", InTime, OutTime, vbMonday, vbFirstJan1)
TotalWorked = TotalWorked - TotalBreak
If TotalWorked > 0 Then
   TotalTime = TotalWorked
End If
If TotalWorked > 0 Then
   CalculateTimeRecord = Format(TotalWorked / 60, "0.00")
Else
   CalculateTimeRecord = 0
End If
ESub:
ErrorHandler:
Exit Function
'Resume Next
End Function

Private Function ValidateDate(ByVal inputDate As String) As Date
    If Not IsDate(inputDate) Then
        ValidateDate = Date
    Else
        ValidateDate = CDate(inputDate)
    End If
End Function
Private Function ValidateDateTime(ByVal inputTime As String, ByVal baseDate As Date) As Date
    ValidateDateTime = baseDate + CDate(inputTime)
End Function

Public Function CalculateBreakTime(ByVal outBreak As String, ByVal inBreak As String, ByVal PaidBreak As Long) As Long
    If Not IsDate(outBreak) Or Not IsDate(inBreak) Then
      CalculateBreakTime = 0
      Exit Function
    End If
    Dim breakDuration As Long
    breakDuration = DateDiff("n", outBreak, inBreak)
    'If breakDuration <= PaidBreak Then
    '    CalculateBreakTime = 0
    'Else
        CalculateBreakTime = breakDuration
    'End If
End Function
'//
Public Function GetTextFromINI(TextIn As String) As String
  GetTextFromINI = GetConfigInfo("Labels", TextIn, "", "FUNCTIONS.INI")
  If GetTextFromINI = "" Then GetTextFromINI = TextIn
End Function
Public Function IsPickUpDate(ItemIn As String, PT As PhoneTime) As String
   Select Case PT
       Case Is = PhoneTime.PDate
           If Not IsDate(ItemIn) Then
              IsPickUpDate = Format(Date, "MM/DD/YYYY")
           Else
              IsPickUpDate = Format(ItemIn, "MM/DD/YYYY")
           End If
           Exit Function
       Case Is = PhoneTime.PTime
           If Not IsDate(ItemIn) Then
              IsPickUpDate = Format(Time, "HH:MM ampm")
           Else
              IsPickUpDate = Format(ItemIn, "HH:MM ampm")
           End If
   End Select
End Function
Public Function FormatPhoneNumber(PHONENO As String) As String
    If InStr(PHONENO, "-") = 0 Then
         Select Case Len(PHONENO)
             Case Is = 7
                PHONENO = Left(PHONENO, 3) & "." & Right(PHONENO, 4)
             Case Is = 10
                PHONENO = Left(PHONENO, 3) & "." & Mid(PHONENO, 4, 3) & "." & Right(PHONENO, 4)
             Case Is = 11
                PHONENO = Mid(PHONENO, 2, 3) & "." & Mid(PHONENO, 5, 3) & "." & Right(PHONENO, 4)
           End Select
    End If
    If InStr(PHONENO, "-") > 0 Then
       PHONENO = Replace(PHONENO, "-", ".")
    End If
    If Check_Option(24) = 1 Then
      If Len(PHONENO) = 8 Then
        If Left(StorePhone, 3) = "800" Or Left(StorePhone, 3) = "888" Then
          
        Else
         PHONENO = Left(StorePhone, 3) & "." & PHONENO
        End If
      End If
    End If
    FormatPhoneNumber = PHONENO
End Function
Public Function AddDiscountDepartmentTotal(NIP As Integer, DiscountName As String, DiscountAmt As Currency, TX1 As Integer, TX2 As Integer, SChg As Single, DiscId As Long)
Dim HMItems     As Integer
Dim SrcFor      As Integer
Dim TotalDepAmt As Currency
Dim DiscountId        As Long

TotalDepAmt = 0
Dim MySummary() As Variant
Store_Grid Order.Ticket(0)
SrcFor = GetDepartmentSeatTotal(NIP, MySummary())
 For HMItems = 0 To SrcFor - 1
  TotalDepAmt = TotalDepAmt + MySummary(1, HMItems)
 Next HMItems
 For HMItems = 0 To SrcFor - 1
  If MySummary(1, HMItems) > 0 And TotalDepAmt > 0 Then
     MySummary(3, HMItems) = (MySummary(1, HMItems) / TotalDepAmt) * DiscountAmt
     Grid_array(NIP, NextRow(NIP), 0) = Order_Qty
     Grid_array(NIP, NextRow(NIP), 1) = Left(MySummary(2, HMItems) & " " & DiscountName, 20)
     Grid_array(NIP, NextRow(NIP), 2) = Format(MySummary(3, HMItems), "0.00")
     Grid_array(NIP, NextRow(NIP), 3) = 1               'Expanded Print
     Grid_array(NIP, NextRow(NIP), 4) = 0               'Remote Printing
     Grid_array(NIP, NextRow(NIP), 5) = ""  'Remote Printer 1
     Grid_array(NIP, NextRow(NIP), 6) = ""  'Remote Printer 2
     Grid_array(NIP, NextRow(NIP), 7) = "" 'Remote Printer 3
     Grid_array(NIP, NextRow(NIP), 8) = ""  'Remote Printer 4
     Grid_array(NIP, NextRow(NIP), 9) = ""  'Remote Printer 5
     Grid_array(NIP, NextRow(NIP), 10) = TX1           'Sales Tax Table
     Grid_array(NIP, NextRow(NIP), 11) = TX2             'Special Tax Table
     Grid_array(NIP, NextRow(NIP), 12) = SChg            'Liquor Tax
     Grid_array(NIP, NextRow(NIP), 13) = 0
     Grid_array(NIP, NextRow(NIP), 14) = AdvanceSelection
     Grid_array(NIP, NextRow(NIP), 19) = -1
     Selection_Number = Val(Grid_array(NIP, NextRow(NIP), 14))
     Grid_array(NIP, NextRow(NIP), 15) = 1000
     Grid_array(NIP, NextRow(NIP), 16) = 0   'Printed
     Grid_array(NIP, NextRow(NIP), 17) = "Discount"  ' DisType    'discount type
     Grid_array(NIP, NextRow(NIP), 18) = 0
     Grid_array(NIP, NextRow(NIP), 20) = LoggedOnAs(0).ManagerOk
     Grid_array(NIP, NextRow(NIP), 21) = ""  '= Void Why
     Grid_array(NIP, NextRow(NIP), 22) = ""  '= Void Time
     Grid_array(NIP, NextRow(NIP), 23) = "Discount"
     Grid_array(NIP, NextRow(NIP), 24) = NIP
     Grid_array(NIP, NextRow(NIP), 25) = DiscountId
     Grid_array(NIP, NextRow(NIP), 26) = MySummary(2, HMItems)
     Grid_array(NIP, NextRow(NIP), 27) = 0 '
     Grid_array(NIP, NextRow(NIP), 28) = "" 'alt print
     Grid_array(NIP, NextRow(NIP), 29) = 0
     Grid_array(NIP, NextRow(NIP), 30) = Left(MySummary(2, HMItems) & " " & DiscountName, 20) 'GetMenu(0)
     Grid_array(NIP, NextRow(NIP), 31) = ""
     Grid_array(NIP, NextRow(NIP), 32) = MySummary(4, HMItems)  'inclusive Tax Field
     Grid_array(NIP, NextRow(NIP), 33) = 0
     Grid_array(NIP, NextRow(NIP), 34) = CoarseLineIsSet
     NextRow(NIP) = NextRow(NIP) + 1
  End If
 Next HMItems
 Order.Load_Grid
End Function
Public Function GetDeliveryPayment(CheckNo As Long) As String
On Error Resume Next
ReDim How_Paid(0)
How_Paid(0).PaymentName = ""
How_Paid(0).PaymentAmt = 0  'Format(Chk_Sales_data(5), "0.00")
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
  Sql = "Select * from [Debit] "
  Sql = Sql & " Where [Check Number] = " & CheckNo
  Sql = Sql & " and [Close Out Day] = 0"
 Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(DebitDb), "Debit")
 If Not Tb Is Nothing Then
    If Not Tb.BOF Then Tb.MoveFirst
 Else
    See_Msg "error connecting", 5, ""
    Exit Function
 End If
 If Not Tb.EOF Then
  How_Paid(0).PaymentName = "Paid"
  How_Paid(0).PaymentAmt = Format(Chk_Sales_data(5), "0.00")
 End If
DbClose Nothing, Tb: Set GRS = Nothing
End Function
Public Function LoadPayments(LoadPayment() As Variant) As Currency
On Error GoTo ERH
Dim FP As Integer
Dim CP As Integer
LoadPayments = 0
ReDim LoadPayment(1, 0)
CP = UBound(How_Paid)
If CP > 0 Then
 ReDim LoadPayment(1, CP)
End If
For FP = 0 To CP
   LoadPayment(0, FP) = How_Paid(FP).PaymentName
   LoadPayment(1, FP) = How_Paid(FP).PaymentAmt
   LoadPayments = LoadPayments + How_Paid(FP).PaymentAmt
Next FP

Exit Function
ERH:
End Function
Public Function QualifiesForDayOt(JobCode As String, EMPName As String, EMPRate As Currency) As Boolean
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
QualifiesForDayOt = True
 If InStr(EMPName, " ") = 0 Then Exit Function
 Dim EMP() As String
 EMP = Split(EMPName, " ")
 Sql = "Select * from [Job Codes] Where [Job Code] ='" & JobCode & "'"
 Sql = Sql & " and [First Name] = '" & EMP(0) & "' and [Last Name] = '" & EMP(1) & "'"
 If ConnectADODB(SvrPath.LocalTerminalPath & "EMPLOYEE.MDB", Db) = False Then Exit Function
 If ConnectADORS(Sql, Db, RS) = False Then Exit Function
 If Not RS.EOF Then
    If IsNull(RS.fields("Day OT").Value) Then
     QualifiesForDayOt = True
    Else
     If Val(RS.fields("Day OT").Value & "") = 1 Then
      QualifiesForDayOt = True
     Else
      QualifiesForDayOt = False
     End If
    End If
 End If
 DbClose Db, RS
End Function
Public Function PayAtTheTable(CheckNo As Long) As String
On Error GoTo ErrorHandler
Dim R               As String
Dim xml             As MPxmlParse
Dim SeqNo           As String
Dim TableNo         As String
Dim Results         As String
Dim Check_TotalAmt  As MyCheckTotal
Dim EmvEquip        As EMVProcessLib.EmvDeviceHardWare
Dim LUPRsp          As String

Check_Status = MyOrderStatus(EditCheck)
TableNo = GetCheckTotal(CheckNo, Check_TotalAmt)
If TableNo = "Failed" Then
   See_Msg "Check Not Found", 5, "Nothing To Pay"
   PayAtTheTable = "Check No Found"
   Exit Function
End If

'If PAY_AT_TABLE = "SLINGSHOT" Then
     '// No longer supported
'Else
     PayAtTheTable = "Exit"
     Find_What = "Select PinPad"
     Load_What = "Pay@Table"
     Select_What.Show 1
     If GetSelected = "" Then
      See_Msg "Nothing Selected", 5, ""
      PayAtTheTable = "No Device Selected"
      Exit Function
     End If
     Dim Sales As AllEMVSales
     Dim PMG As PAXMSG
     Dim CA As ver11DebitCards.AccountInfo
     Set Sales = New AllEMVSales
      If Sales.Item(CStr(Check_no)) Is Nothing Then
        Sales.add CStr(Check_no)
      End If
      Sales.Item(CStr(Check_no)).SalesAmt = Format(Check_TotalAmt.ChkTotal, "0.00")
      Sales.Item(CStr(Check_no)).TipAmt = 0
      Sales.Item(CStr(Check_no)).TaxAmt = Check_TotalAmt.TaxTotal
      Sales.Item(CStr(Check_no)).Cashier = LoggedOnAs(0).EMPName
      Sales.Item(CStr(Check_no)).Term = MyRegisterName
      Sales.Item(CStr(Check_no)).TranType = "Credit"
      Sales.Item(CStr(Check_no)).TransNo = CStr(CheckNo)
      Sales.Item(CStr(Check_no)).CashBack = 0
      If LoadEMVDevice(GetSelected, EmvEquip) = "Failed" Then
        See_Msg "No Device Found", 5, "Check Device Setup"
        Exit Function
      End If
      Set PMG = New PAXMSG
      Dim PXFO As EMVProcessLib.Process
      Set PXFO = New EMVProcessLib.Process
      If IsDataCapInterface(EMVINTERFACE) Then
          Results = PXFO.ProcessSale(Sales, DevSaleType_Checkout, PMG, EmvEquip)
      End If
      If Results = "CONNECT ERROR" Then
       Call PXFO.ProcessSale(Sales, DevSaleType_XVoidbyRecordNo, PMG, EmvEquip)
       See_Msg Results, 5, ""
       Exit Function
      End If
      Set xml = New MPxmlParse
      xml.xml = Results
      If xml.token("CmdStatus") = "Success" Then
       SeqNo = xml.token("PayAtTheTableID")
       StoreTableSeqNo SeqNo, CheckNo, Results, EmvEquip
       PayAtTheTable = "Success"
      Else
       See_Msg "PinPad Failed To Update", 5, xml.token("CmdStatus")
       PayAtTheTable = "Failed"
      End If
      Set xml = Nothing
   'End If
Exit Function
ErrorHandler:
If Err.Number = 91 Then
  Error_Log 91, "Error Connecting to Pin Pad", "Pin Pad Pay@Table", ""
  If See_Msg("Failed To Connect To PinPad", 4, "Yes to Retry Or No To Quit") = 6 Then
    Resume
  End If
  Exit Function
Else
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Out of Scope ")
End If
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetTransformerShift(InShiftLvl As Long, InCategory As String) As Long
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
GetTransformerShift = Val(GL.GetTransformerShift(InShiftLvl, InCategory))
Set GL = Nothing
End Function
Public Function TransformPLU(ShiftPLUno As Long, MenuPlu As String) As Long
  If ShiftPLUno > 0 Then
    If MenuPlu > 0 And Len(CStr(ShiftPLUno)) < 4 Then
      TransformPLU = Val(CStr(ShiftPLUno) & CStr(MenuPlu))
    Else
      TransformPLU = ShiftPLUno
    End If
  Else
    TransformPLU = MenuPlu
  End If
  ShiftPLUno = 0
End Function
Public Function IssuePLUNO(ItemName As String, PLUNO As String) As String
 If Not IsNumeric(PLUNO) Then
  IssuePLUNO = 0
 Else
  IssuePLUNO = PLUNO
 End If
End Function
Public Function Is64bit() As Boolean
    Dim handle As Long, bolFunc As Boolean
    bolFunc = False
    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")
    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function
Public Function GetDepartmentNo(Item As String) As Long
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
If Len(Item) > 3 And Left$(Item, 3) = "EBT" Then
 Item = Right$(Item, Len(Item) - 4)  '// looks up department name in the ini and returns the department number for EBT
 GetDepartmentNo = Val(GetConfigInfo("EBT", "Get EBT Department No For " & Item, CStr(GL.GetDepartmentNo(Item)), ""))
Else
 GetDepartmentNo = GL.GetDepartmentNo(Item)
End If
Set GL = Nothing
End Function
Public Function WhoIsManager() As String
If LoggedOnAs(0).EMPName <> LoggedOnAs(0).ManagerOk And LoggedOnAs(0).ManagerOk > "" Then
  WhoIsManager = LoggedOnAs(0).ManagerOk
Else
  WhoIsManager = LoggedOnAs(0).EMPName
End If
End Function
Public Function GetRevenueCenter(TerminalCenter As String, EmpAsCenter As String, TableAsCenter As String) As String
On Error GoTo ERH
GetRevenueCenter = TerminalCenter
If InStr(TableAsCenter, "SN") > 0 Then
  TableAsCenter = Trim$(Left$(TableAsCenter, InStr(TableAsCenter, "SN") - 2))
  'Added for QCR Make sure Rev center is updated on split checks 08/15/2018
End If
If TableAsCenter > "" Then
  If Not ALLRevCenter.Item(TableAsCenter) Is Nothing Then
   GetRevenueCenter = ALLRevCenter.Item(TableAsCenter).PrinterName  '// printer name is alias for Revenue center
   Exit Function
  End If
End If
If EmpAsCenter > "" Then
 If Not ALLRevCenter.Item(EmpAsCenter) Is Nothing Then
  GetRevenueCenter = ALLRevCenter.Item(EmpAsCenter).PrinterName   '// printer name is alias for Revenue center
  Exit Function
 End If
End If
ERH:
GetRevenueCenter = TerminalCenter
End Function
Public Function OKTODELETE() As Boolean
OKTODELETE = True
Dim AL As Integer
Dim i  As Integer
For i = 0 To NextRow(SeatNumber) - 1
   If Grid_array(SeatNumber, i, 2) < 0 Then
      If Val(Grid_array(SeatNumber, i, 0)) > 0 Then
      ' OKTODELETE = False
       Exit For
      End If
   End If
Next i
End Function
Public Function GetSvrJobCode(SvrName As String, AuditDate As String, ShiftStart As String) As String ' TipOutPosition
Dim Db As ADODB.Connection
Dim OTB As ADODB.Recordset 'must use desc order to get last job worked.
ConnectADODB TimeRecordsDb, Db
Sql = "Select * from [Employee Records] where [Employee name] = '" & SvrName & "'"
Sql = Sql & " and TRY_CAST([Shift Date] AS date) = '" & AuditDate & "' Order by TRY_CAST([Time In] AS datetime) desc"
ConnectADORS Sql, Db, OTB
   If Not OTB.EOF Then
     GetSvrJobCode = OTB![Department Name] & ""
   End If
DbClose Db, OTB
End Function
Public Function GetIpAddress(TerminalName As String) As String
On Error GoTo ERH
If MyRegisterName = TerminalName Then
  GetIpAddress = MyRegisterName
  Exit Function
End If

Dim GetIP As GeneralLib.Utilities
Set GetIP = New GeneralLib.Utilities
GetIpAddress = GetIP.GetIPAddressByComputerName(TerminalName)
If InStr(GetIpAddress, "Fail") > 0 Then
  GetIpAddress = TerminalName
End If
Set GetIP = Nothing
Exit Function
ERH:
GetIpAddress = MyRegisterName
Set GetIP = Nothing
End Function
Public Function IsTermialThere(SAddress As String) As Boolean
IsTermialThere = True
Dim MyIpAddress As String
MyIpAddress = GetIpAddress(SAddress)
If IsValidIPAddress(MyIpAddress) Then
 Dim GRS As GeneralLib.Utilities
 Set GRS = New GeneralLib.Utilities
 IsTermialThere = GRS.IsTerminalPresent(MyIpAddress)
 Set GRS = Nothing
End If
End Function
Public Function TerminalIsPresent(StationName As String) As Boolean
 On Error GoTo ERH
 TerminalIsPresent = True
 If UCase(ReturnComputerName.MachineName) = UCase(StationName) Then
    TerminalIsPresent = True
    Exit Function
 End If
 If DisablePing = "ENABLED" Then TerminalIsPresent = True: Exit Function
 TerminalIsPresent = IsTermialThere(StationName)
 Exit Function
ERH:
End Function
' Function to check if an IP address is valid
Public Function IsValidIPAddress(ByVal IpAddress As String) As Boolean
    Dim octets() As String
    Dim i As Integer
    Dim num As Integer
    
    ' Split the IP address into its four octets
    octets = Split(IpAddress, ".")
    
    ' Check if there are exactly four octets
    If UBound(octets) <> 3 Then
        IsValidIPAddress = False
        Exit Function
    End If
    
    ' Check each octet
    For i = 0 To 3
        ' Ensure the octet is numeric and within the valid range
        If IsNumeric(octets(i)) Then
            num = CInt(octets(i))
            If num < 0 Or num > 255 Then
                IsValidIPAddress = False
                Exit Function
            End If
        Else
            IsValidIPAddress = False
            Exit Function
        End If
    Next i
    
    ' If all checks pass, the IP address is valid
    IsValidIPAddress = True
End Function
' Example usage
Public Function GetMyIpAddress(TName As String) As String
GetMyIpAddress = TName
 Select Case UCase(TName)
    Case Is = UCase(SvrPath.TerminalName)
      GetMyIpAddress = SvrPath.TerminalIpAddress
      Exit Function
    Case Is = UCase(SvrPath.TerminalName)
      GetMyIpAddress = SvrPath.TerminalIpAddress
      Exit Function
 End Select
End Function
'// payroll functions
Public Function GetStartPayRollTime() As String
 If Not IsDate(Input_Option(84)) Then   '// end of business day  -- default is 4:00 am
   GetStartPayRollTime = "04:00 am"
 Else
   GetStartPayRollTime = DateAdd("n", 1, Input_Option(84))
 End If
End Function
Public Function GetEndPayroll() As String
 If Not IsDate(Input_Option(84)) Then  '// end of business day  -- default is 4:00 am
   GetEndPayroll = "03:59 am"
 Else
   GetEndPayroll = DateAdd("n", -1, Input_Option(84))
 End If
End Function
Public Function GetEndPayRollDate(SearchDate) As String
   Dim EndTime As String
   If Not IsDate(SearchDate) Then SearchDate = Format(Date, "MM/DD/YYYY")
   If Not IsDate(GetEndPayroll) Then
     EndTime = "4:00 am"
     GetEndPayRollDate = CDate(SearchDate) + 1
     Exit Function
   End If
   Select Case Hour(GetEndPayroll)
      Case Is = 0, 12, 1, 2, 3, 4, 5, 6, 7, 8
        GetEndPayRollDate = CDate(SearchDate) + 1
      Case Else
        GetEndPayRollDate = SearchDate
   End Select
End Function
'// end payroll functions
'// start and end of business day
Public Function GetStartTime() As String
 If Not IsDate(Input_Option(84)) Then  '// end of business day  -- default is 4:00 am
   GetStartTime = DateAdd("n", 1, "04:00 am")
 Else
   GetStartTime = DateAdd("n", 1, Input_Option(84))
 End If
End Function
Public Function SetEndingTime() As String
      If Not IsDate(Input_Option(84)) Then '// end of business day  -- default is 4:00 am
       SetEndingTime = "4:00 am"
      Else
       SetEndingTime = Input_Option(84)
      End If
End Function
Public Function SetEndingDate(SearchDate As String) As String
   Dim EndTime As String
   If Not IsDate(SearchDate) Then SearchDate = Format(Date, "MM/DD/YYYY")
   If Not IsDate(Input_Option(84)) Then '// end of business day  -- default is 4:00 am
     EndTime = "4:00 am"
     SetEndingDate = CDate(SearchDate) + 1
     Exit Function
   End If
   Select Case Hour(Input_Option(84))
      Case Is = 0, 12, 1, 2, 3, 4, 5, 6, 7, 8
        SetEndingDate = CDate(SearchDate) + 1
      Case Else
        SetEndingDate = SearchDate
   End Select
End Function
Public Function GetBusinessDay() As String
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
GetBusinessDay = GL.FindBusinessDay
Set GL = Nothing
End Function
'// end of business day

Public Sub RecordJobCodeTips(ClosingDate As String)
On Error GoTo ErrorHandler
'Make sure delete for open audit
'If Check_Option(206) = 0 Then Exit Sub
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim T               As Integer
Dim THrs            As Single
Dim T_Dollars       As Currency
ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
Sql = "Select Distinct [Tip Share Job Code], Sum([Tip Share]) as TipOut from [Tip Share] Where [Close out day] = 0"
Sql = Sql & " Group By [Tip Share Job Code]"
ConnectADORS Sql, Db, RS
'T = RecordsCount(RS)
Do While Not RS.EOF
  If RS![TipOut] > 0 Then
    THrs = GetTipShareHours(RS![Tip Share Job Code] & "", ClosingDate)
    Trans_Log RS![Tip Share Job Code] & " Hrs " & THrs & " Dollars " & RS![TipOut]
    If THrs > 0 And RS![TipOut] > 0 Then
      T_Dollars = Format(RS![TipOut] / THrs, "0.00")
      DeclareTipShare RS![Tip Share Job Code], T_Dollars, ClosingDate
      Trans_Log RS![Tip Share Job Code] & " Dollars " & T_Dollars
    End If
  End If
 RS.MoveNext
Loop
DbClose Db, RS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "RecordJobCodeTips ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function GetTipShareHours(JobCode As String, ClosingDate As String) As Single
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim Hoursworked As Single
GetTipShareHours = 0: Hoursworked = 0
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Hoursworked = 0
OpenDb TimeRecordsDb, Db
 Sql = "Select * From [Employee Records] Where [Department Name] = '" & JobCode & "' and TRY_CAST([Shift Date] AS date) = '" & CDate(ClosingDate) & "'"
ConnectADORS Sql, Db, Tb
       Do While Not Tb.EOF
        Hoursworked = Hoursworked + CalculateTimeRecord(CStr(Tb.fields("Shift date").Value & ""), CStr(Tb.fields("Punch Out Date").Value & ""), CStr(Tb.fields("Time In").Value & ""), CStr(Tb.fields("Break Out").Value & ""), CStr(Tb.fields("Break In").Value & ""), CStr(Tb.fields("Time Out").Value & ""), 0)
        Trans_Log JobCode & " " & Hoursworked
        Tb.MoveNext
     Loop
DbClose Db, Tb
GetTipShareHours = Hoursworked
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Calculate Labor Tip Share ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub RecordTipShareTo_History(Db As ADODB.Connection, ShiftDate As String, Emp_Name As String, ShiftNo As Long, ShareAmt As Currency)
On Error GoTo ERH
Dim RS          As ADODB.Recordset
  Sql = "Select * from [Tip Income] Where TRY_CAST([date] AS date) = '" & ShiftDate & "' and [Server Name] = '" & Emp_Name & "'"
  Sql = Sql & " and [Shift] = " & ShiftNo
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
   RS.fields("Tip Share").Value = Format(ShareAmt, "0.00")
   RS.Update
End If
DbClose Nothing, RS
Exit Sub
ERH:
Err.Clear
End Sub
Public Sub DeclareTipShare(JobCode As String, HourlyRate As Currency, ClosingDate As String)
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim sDb         As ADODB.Connection
Dim Hoursworked As Single
Hoursworked = 0
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
ConnectADODB "MSaleData", sDb
OpenDb TimeRecordsDb, Db
 Sql = "Select * From [Employee Records] Where [Department Name] = '" & JobCode & "' and TRY_CAST([Shift Date] AS date) = '" & ClosingDate & "'"
ConnectADORS Sql, Db, Tb
       Do While Not Tb.EOF
         Hoursworked = CalculateTimeRecord(CStr(Tb(2)) & "", CStr(Tb(7)) & "", CStr(Tb(3)) & "", CStr(Tb![Break Out] & ""), CStr(Tb![Break In] & ""), CStr(Tb(6)) & "", 0)
         Tb![Tip Share] = Format(Hoursworked * HourlyRate, "0.00")
         Tb.Update
         RecordTipShareTo_History sDb, Tb.fields("Shift Date").Value, Tb.fields("Employee Name").Value, Tb.fields("Shift").Value, Tb.fields("Tip Share").Value
       Tb.MoveNext
     Loop
DbClose Db, Tb
DbClose sDb, Nothing
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "DeclareTipShare ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Public Function GetPhoneOrderSales() As Currency  '// this function is used to add open phone sales to current day sales report
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim OP As ADODB.Recordset
GetPhoneOrderSales = 0
ConnectADODB CallInOrderDb, Db
ConnectADORS "Select * from [Open Checks] Where [Presented] = 1", Db, OP
       Do While Not OP.EOF
        GetPhoneOrderSales = GetPhoneOrderSales + (OP(6) - OP(11)) - (OP(10) - OP(15))
        OP.MoveNext
       Loop
DbClose Db, OP
Exit Function
ERH:
Err.Clear
DbClose Db, OP
End Function
Public Function ReturnTabs(SearchHow As String) As Integer
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim TBGA As String
Dim NOTs As Long
Dim PIG  As Long
ConnectADODB CheckStatDb, Db
If Check_Option(93) = 0 Then  'Make sure on transfer server will see own check
   If SvrPath.GroupAudit Then
      TBGA = " Where [Close Out Day] = 0 and [Transfer From] = '" & SvrPath.TerminalGroupName & "' or [server Name] = '" & LoggedOnAs(0).EmpAssignedTo & "'"
   Else
      TBGA = " Where [Close Out Day] = 0 and [Transfer From] = '" & SvrPath.TerminalGroupName & "' or [server Name] = '" & LoggedOnAs(0).EMPName & "'"
   End If
Else
      TBGA = " Where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EMPName & "'"
End If
If Check_Option(140) = 1 Then  'allow open tabs from QSR Options
   If Check_Option(93) = 1 Then
      Sql = "Select * from [Open Checks] where [Close Out Day] = 0 and [Register Name] = '" & MyRegisterName & "' order by [Table Number] asc"
   Else
      Sql = "Select * from [Open Checks] Where [Close Out Day] = 0 and [Transfer From] = '" & SvrPath.TerminalGroupName & "' order by [Table Number] asc"
   End If
Else
   If Val(LoggedOnAs(0).EmpAccessLevel) > Access_Required("View All Checks") Then  'And Check_Option(93) = 0
       Sql = "Select * from [Open Checks] where [Close Out Day] = 0 order by [Table Number] asc"
   Else
       Sql = "Select * from [Open Checks] " & TBGA & " order by [Table Number], [check number] asc"
   End If
End If
ConnectADORS Sql, Db, Tb
NOTs = RecordsCount(Tb)
DbClose Db, Tb
ReturnTabs = NOTs
Exit Function
ErrorHandler:
Err.Clear
End Function

Public Function RouteVideo() As Boolean
If Check_Option(102) = 1 Then
 If FileCheck(SvrPath.TerminalPath & "VideoRouting.ON") Then
   RouteVideo = True
 Else
   RouteVideo = False
 End If
Else
  RouteVideo = False
End If
'// used for kitchen video routings
'// org used for wendys center grill stores
End Function
Public Function RoutingIs() As Boolean
If Check_Option(102) = 1 Then
 If FileCheck(SvrPath.TerminalPath & "Routing.ON") Then
   RoutingIs = True
 Else
   RoutingIs = False
 End If
Else
   RoutingIs = False
End If
End Function
Public Function ParseIpCall(DataIn As String) As String
    On Error Resume Next
    Dim SD      As Integer
    Dim oLine   As String
    Dim oPhone  As String
    Dim oName   As String
        InStr(DataIn, "$") = SD
        For SD = Len(DataIn) To 1 Step -1
         If Mid(DataIn, SD, 1) = "$" Then Exit For
        Next SD
        oLine = Mid$(DataIn, SD + 1, 2)
        oPhone = Mid$(DataIn, SD + 32, 16)
        oPhone = Trim$(Replace(oPhone, "-", "."))
        If InStr(UCase(DataIn), "PRIVATE") > 0 Then
          oName = "Private"
        ElseIf InStr(UCase(DataIn), "OUT-OF-AREA") > 0 Then
          oName = "Out-of-Area"
        Else
          oName = Right$(DataIn, 15)
        End If
        ParseIpCall = oLine & "|" & oPhone & "|" & oName & "|" & "Ring"
        '// parse caller id calls for display
End Function
Public Function GetConfigInfo(Header As String, AccessLine As String, defaultValue As String, FileName As String) As String
If MiscFunctions Is Nothing Then
 GetConfigInfo = defaultValue
 BuildFunctionINI Header, AccessLine, defaultValue
ElseIf MiscFunctions.Item(AccessLine) Is Nothing Then
 GetConfigInfo = defaultValue
 BuildFunctionINI Header, AccessLine, defaultValue
Else
 GetConfigInfo = MiscFunctions.Item(AccessLine).m_AccessValue
End If
Select Case UCase(GetConfigInfo)
  Case Is = "ON", "YES", "TRUE", "ENABLED"
     GetConfigInfo = "ENABLED"
  Case Is = "OFF", "NO", "FALSE", "DISABLED"
     GetConfigInfo = "DISABLED"
  Case Else
    '//do nothing
  End Select
  '// used for MiscFunctions from misc items database
End Function
Public Sub BuildFunctionINI(Section As String, FunctionName As String, FunctionValue As String)
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB("MSaleData", Db) = False Then
   See_Msg "Error Connecting to MSaleData", 5, ""
End If
Dim MySql As String
MySql = "Select * from MiscFunctions where FunctionName = '" & FunctionName & "'"
If ConnectADORS(MySql, Db, RS) = False Then
   See_Msg "Error Connecting to MSaleData", 5, "Record Set"
End If
If RS.EOF Then
   RS.AddNew
   RS.fields("SectionName").Value = Section
   RS.fields("FunctionName").Value = FunctionName
   RS.fields("FunctionValue").Value = FunctionValue
   RS.Update
End If
DbClose Db, RS
Exit Sub
ERH:
Trans_Log "Error Writing to Misc Functions " & FunctionName
DbClose Db, RS
End Sub
Public Function LoadMiscFunctions()
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Set MiscFunctions = Nothing
If ConnectADODB("MSaleData", Db) = False Then
   See_Msg "Error Connecting to MSaleData", 5, ""
End If
Sql = "Select * from MiscFunctions Where FunctionValue > '' "
If ConnectADORS(Sql, Db, RS) = False Then
   See_Msg "Error Connecting to MSaleData", 5, "Record Set"
End If
If Not RS.EOF Then
 Set MiscFunctions = New AllLevels
End If
Do While Not RS.EOF
   If MiscFunctions.Item(RS.fields("FunctionName").Value) Is Nothing Then
     MiscFunctions.add RS.fields("FunctionName").Value
   End If
   MiscFunctions.Item(RS.fields("FunctionName").Value).m_AccessValue = RS.fields("FunctionValue").Value & ""
   RS.MoveNext
Loop
DbClose Db, RS
Exit Function
ERH:
See_Msg Err.Description, 5, "Loading Misc Function"
Trans_Log "Error Loading Misc Functions " & Err.Description
Resume Next
End Function
Public Function GETSVRDWNINFO(Header As String, INFO As String, FileName As String) As String
Dim Buff As String
Dim Ret  As Integer
Buff = String(255, 0)
Ret = GetPrivateProfileString(Header, INFO, "", Buff, 256, FileName)
GETSVRDWNINFO = Replace(Buff, Chr(0), "")
End Function
Public Function GetLastOrderDate(CName As String, CPhone As String) As String
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
   GetLastOrderDate = GL.ReturnLastVisit(CName, CPhone)
Set GL = Nothing
End Function
Public Function CountCoupons(C_Name As String, Limit As Integer) As Boolean
Dim FC      As Integer
Dim Counter As Integer
Counter = 0
If Limit = 0 Then CountCoupons = True: Exit Function
  With Order.Ticket(0)
    For FC = 0 To NextRow(SeatNumber) - 1
      If Trim(C_Name) = Trim(.TextMatrix(FC, 1)) Then
         Counter = Counter + 1
      End If
    Next FC
  End With
  If Counter >= Limit Then
    CountCoupons = False
  Else
    CountCoupons = True
  End If
End Function
Public Function GetRoutedDestinationPrinter() As String
 GetRoutedDestinationPrinter = ""
 If GetSetting("Micro$ale", "Register Setup", OrderIs & " Printer") > "" Then
  If Not AllPrinter.Item(GetSetting("Micro$ale", "Register Setup", OrderIs & " Printer Name")) Is Nothing Then
    GetRoutedDestinationPrinter = AllPrinter.Item(GetSetting("Micro$ale", "Register Setup", OrderIs & " Printer Name")).PrinterName
  End If
 End If
End Function
Public Function GetRefundNo() As Long
GetRefundNo = 1
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
 If ConnectADODB(FinancialDb, Db) = False Then Exit Function
    Sql = "Select top 1 [check number] from Refund where [close out day] = 0 "
    Sql = Sql & " order by [check number] desc"
 If ConnectADORS(Sql, Db, RS) = False Then Exit Function
 If Not RS.EOF Then
   GetRefundNo = Val(RS![Check Number] & "") + 1
 Else
   GetRefundNo = 1
 End If
 DbClose Db, RS
 If ReturnKVSTerminalId > 0 Then
  If GetRefundNo < 100 Then
  GetRefundNo = Val(ReturnKVSTerminalId * 10 & GetRefundNo)
  End If
 End If
End Function
Public Function AdminIsOn() As Boolean
On Error Resume Next
Dim Db         As ADODB.Connection
Dim RS         As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & "MSaleData", Db
  Sql = "Select * from [Administrator]"
OpenTb Sql, RS, Db
 If Not RS.EOF Then
  AdminIsOn = True
 Else
  AdminIsOn = False
 End If
DbClose Db, RS
End Function
Public Function AdminPasswordOK(User As String) As Boolean
'On Error GoTo errorhandler
  Dim Db         As ADODB.Connection
  Dim RS         As ADODB.Recordset
  AdminPasswordOK = True
  OpenDb SvrPath.TerminalPath & "Employee.mdb", Db
    Sql = "Select * from [Administrator] where Ucase([Employee Name]) = '" & UCase(User) & "'"
  OpenTb Sql, RS, Db
      If Not RS.EOF Then
        If IsDate(RS![Date]) Then
           If DateDiff("d", RS![Date], Date) > 90 Then
             AdminPasswordOK = False
             See_Msg "New Password Required", 0, "Password Has Expired !! "
           Else
             AdminPasswordOK = True
           End If
        End If
      End If
      DbClose Db, RS
End Function
Public Function GetAdministrator(Accessfor As String) As Boolean
On Error GoTo ErrorHandler
  Dim Db         As ADODB.Connection
  Dim RS         As ADODB.Recordset
  Dim UserName   As String
  Dim PassWord   As String
  Dim DecryptKey As String
  GetAdministrator = False
  UserName = ShowKeyBoard("Enter User Name")
  If UserName > "" Then
     PassWord = ShowKeyBoard("Enter Password")
     If PassWord = "" Then
       Exit Function
     End If
  Else
    Exit Function
  End If
    OpenDb SvrPath.TerminalPath & "Employee.mdb", Db
      Sql = "Select * from [Administrator] where Ucase([User Name]) = '" & UserName & "'"
    OpenTb Sql, RS, Db
      If Not RS.EOF Then
         DecryptKey = StoreName & " " & StoreAddress
         If UCase(UserName) = UCase(RS![User Name] & "") And UCase(PassWord) = UCase(V8Decrypt(RS![PassWord] & "", DecryptKey)) Then
            RS![Access] = 1
            RS![Time] = Format(Time, "HH:mm ampm")
            RS.Update
            GetAdministrator = AdminPasswordOK(LoggedOnAs(0).EMPName)
         Else
              If Not IsDate(RS![Time]) Then
                RS![Access] = 1
                RS![Time] = Format(Time, "HH:mm ampm")
                RS.Update
                See_Msg "No Match", 5, "Access Denied."
                GetAdministrator = False
                DbClose Db, RS
                Exit Function
              End If
              If IsDate(RS![Time]) And DateDiff("n", RS![Time], Time) > 30 Then
                RS![Access] = 1
                RS![Time] = Format(Time, "HH:mm ampm")
                RS.Update
                See_Msg "Wrong Password", 5, "Access Denied."
                GetAdministrator = False
                DbClose Db, RS
                Exit Function
              End If
              If Val(RS![Access] & "") >= 4 Then
               If DateDiff("n", RS![Time], Time) < 30 Then
                 See_Msg "Account Locked for 30 Minutes", 0, "Time Remaining " & (30 - DateDiff("n", RS![Time], Time))
                 DbClose Db, RS
                 Exit Function
               End If
              End If
              RS![Access] = Val(RS![Access] & "") + 1
              RS![Time] = Format(Time, "HH:mm ampm")
              RS.Update
              See_Msg "No Match", 5, "Access Denied."
              GetAdministrator = False
         End If
      Else
         See_Msg "No Match", 5, "Access Denied."
         GetAdministrator = False
      End If
DbClose Db, RS
Exit Function
ErrorHandler:
Sleep 0
See_Msg "Error", 0, Err.Description
GetAdministrator = False
End Function
Public Function V8Decrypt(DataIn As String, PWord As String) As String
On Error GoTo ERH
If DataIn = "" Then
   V8Decrypt = ""
   Exit Function
End If
V8Decrypt = DataIn
Dim strTmp       As String
Dim arData()     As Byte     ' added 08-Jan-2001 KCI
Dim arPassWord() As Byte
Dim cCrypto      As CryptKci.clsCryptoAPI
Set cCrypto = New CryptKci.clsCryptoAPI
arPassWord = cCrypto.StringToByteArray(PWord)
cCrypto.PassWord = arPassWord()
strTmp = cCrypto.ConvertStringFromHex(DataIn)
arData = cCrypto.StringToByteArray(strTmp)
cCrypto.InputData = arData()
If cCrypto.Decrypt(2, 1) Then
   arData = cCrypto.OutputData
   V8Decrypt = cCrypto.ByteArrayToString(arData)
End If
Set cCrypto = Nothing    ' Free the Crypto class from memory
strTmp = String$(250, 0)  ' overwrite data in temp variable
Exit Function
ERH:
Err.Clear
End Function
Public Function GetCountDownQty(ItemName As String) As Long
On Error GoTo ERH
Dim PST As Integer
PST = InStr(ItemName, ")")
GetCountDownQty = 0
If PST > 0 And Left(ItemName, 1) = "(" Then
   If IsNumeric(Mid(ItemName, 2, PST - 2)) Then
    GetCountDownQty = Val(Mid(ItemName, 2, PST - 2))
   Else
    GetCountDownQty = 0
   End If
End If
Exit Function
ERH:
GetCountDownQty = 0
End Function
Public Function GetItemCost(ByVal MenuItem As String, ByVal qty As Single) As Currency
On Error GoTo ErrorHandler
If qty = 0 Then qty = 1
If Not AllMenuItems(MenuItem) Is Nothing Then
 If Check_Option(69) = 1 Then
   GetItemCost = ReturnCost(MenuItem, qty)
 Else
   GetItemCost = AllMenuItems(MenuItem).ItemCost * qty
 End If
Else
 If Check_Option(69) = 1 Then
   GetItemCost = ReturnCost(MenuItem, qty)
 End If
End If
Exit Function
ErrorHandler:
Err.Clear
'Cost get droped on open checks split
End Function
Public Function ReturnCost(ByVal MenuItem As String, ByVal qty As Single) As Currency
Dim WTC As FoodCostSync.Upload
Set WTC = New FoodCostSync.Upload
ReturnCost = WTC.GetItemCost(SvrPath.TerminalPath, RPMDb, qty, MenuItem, 0)
Set WTC = Nothing
'// used to look up cost for wasted items.
End Function
Public Function GetPLUNO(ByVal MenuItem As String) As String
On Error Resume Next
If Not AllMenuItems(MenuItem) Is Nothing Then
    GetPLUNO = Trim$(AllMenuItems(MenuItem).BARCODE)
    If GetPLUNO = "None" Then GetPLUNO = ""
Else
    GetPLUNO = ""
End If
'// find menu item plu
End Function
Public Function GetCountDownItem(ByVal ItemName As String) As String
On Error GoTo ErrorHandler
Dim PST As Integer
PST = InStr(ItemName, ")")  'Stops Error 6 Johnny Customer Requests 9.08.2010
If PST > 0 And Left(ItemName, 1) = "(" Then
   If IsNumeric(Mid(ItemName, 2, PST - 2)) Then
    
   Else
    GetCountDownItem = ItemName
    Exit Function
   End If
End If
PST = InStr(ItemName, ")")  'Stops Error 6 Johnny Customer Requests 9.08.2010
If PST > 0 And Left(ItemName, 1) = "(" Then
   GetCountDownItem = Trim$(Mid(ItemName, PST + 1, Len(ItemName) - PST))
   If GetCountDownItem = "" Then GetCountDownItem = ItemName
Else
   GetCountDownItem = ItemName
End If
Exit Function
ErrorHandler:
GetCountDownItem = ItemName
Err.Clear
End Function
Public Function GetRandomCoupon() As String
GetRandomCoupon = (CStr(RegisterCheckSum) & (((Val(Format(Date, "y")) + Val(Format(Time, "hh"))) + Val(Format(Time, "mm"))) + Val(Format(Time, "ss"))))
Sleep 0
'// builds random coupons with serial numbers for usage tracking
End Function
Public Function RegisterCheckSum() As Long
Dim R As Integer
For R = 1 To Len(MyRegisterName)
   RegisterCheckSum = RegisterCheckSum + Asc(Mid(MyRegisterName, R, 1))
Next R
'// part of random coupons.
End Function
Public Function CheckForOpenTabs(Svr As String) As Boolean
Dim ThisRegister As String
If Check_Option(0) = 1 Then  '// quick service option
  ThisRegister = MyRegisterName
Else
  ThisRegister = ""
End If
'// check option(0) = 1 is for quick service
'// tabs are sorted by terminal for quick service and by server for full service.
'// bars option (39) is the exception is can share tabs by register group or not
'// option (93) will disable bar tab sharing bartender only see's their tabs
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
  CheckForOpenTabs = GL.IHaveOpenChecks(Svr, ThisRegister)
Set GL = Nothing
End Function
Public Function GetTimePlayCharge(ItemIn As String, StartTime As String, ItemStrName As String) As Currency
'// Menu item should have the hourly close set
'// software will charge by the minute
'// timed pricing can span through timed events and charge different prices by day part.
'// Prices will reflect each event played through -- regular time, happy hour etc.
Dim CostPerMin As Currency
ItemStrName = ""
If AllMenuItems(ItemIn).RegularPrice = 0 Then
   GetTimePlayCharge = 0
   Exit Function
End If
CostPerMin = AllMenuItems(ItemIn).RegularPrice / 60
ItemStrName = ItemIn & " (" & (ReturnTotalMins(CDate(StartTime), CDate(Time & Date))) & ") @ " & Format(CostPerMin, "0.00")
GetTimePlayCharge = Format(Val(CostPerMin * ReturnTotalMins(CDate(StartTime), CDate(Time & Date))), "0.00")
End Function
Public Function GetSortOrder(ItemIn As String, CatSort As Integer) As Integer
If Check_Option(115) = 1 Then  '//Menu items are in order as entered on the check
   GetSortOrder = 0
Else
  If AllMenuItems(ItemIn) Is Nothing Then '// Menu items will be sorted in category order
      GetSortOrder = CatSort
      Exit Function
  End If
  If Left$(AllMenuItems(ItemIn).ItemInfo, 6) = "Coarse" Then      '// coarse line can change sort order based on Coarse
    GetSortOrder = Val(Right$(AllMenuItems(ItemIn).ItemInfo, 1))
    Exit Function
  End If
  If AllMenuItems(ItemIn).SortOrder > 0 Then   'menu items can be set in a specific order - setting is programmed menu items
     GetSortOrder = AllMenuItems(ItemIn).SortOrder
  Else
     GetSortOrder = CatSort
  End If
End If
End Function
Public Function ReturnGroupName(DataIn As String, GroupNameIn As String) As String
          If GroupNameIn = "" Then GroupNameIn = ModGroupName
          If Left(GroupNameIn, 1) = "!" Then
            DataIn = Mid(GroupNameIn, 2, Len(GroupNameIn) - 1) & " " & DataIn
          ElseIf Left(GroupNameIn, 1) = "@" Then
            DataIn = Mid(GroupNameIn, 2, Len(GroupNameIn) - 1)
          Else
            DataIn = Trim$(DataIn) & " " & GroupNameIn
          End If
          If Len(DataIn) > 20 Then DataIn = Left(DataIn, 20)
          ReturnGroupName = DataIn
End Function
Public Function ReturnGroupNameSuffix(DataIn As String, GroupNameIn As String) As String
          DataIn = GroupNameIn & " " & DataIn
          If Len(DataIn) > 20 Then DataIn = Left(DataIn, 20)
          ReturnGroupNameSuffix = DataIn
End Function
Public Function PSTRule(DeptName As String) As Boolean
Dim f As Integer
For f = 0 To UBound(DeptEvents, 1)
  If DeptName = DeptEvents(f).DepartmentName Then
   If DeptEvents(f).PSTTaxRule = -1 Then
      PSTRule = True
   Else
      PSTRule = False
   End If
   Exit Function
  End If
Next f
PSTRule = False
'// taxing rules for candian sales tax
End Function
Public Function TestRules()
PSTMinTaxOn = False
Dim N As Integer
Dim R As Integer
For N = 0 To Number_In_Party - 1
 For R = 0 To NextRow(N) - 1
   If PSTRule(CStr(Grid_array(N, R, 26))) = True Then
   PSTMinTaxOn = True
   Exit Function
   End If
 Next R
Next N
End Function
Public Function InventoryIsOn() As Boolean
'On Error GoTo errorhandler
'If GetConfigInfo("MenuItems", "Count Down Always", "OFF", "FUNCTIONS.INI") = "ENABLED" Then
'  InventoryIsOn = True
'  InventoryIsSet = True
'  Exit Function
'End If
InventoryIsOn = False
'If Check_Option(0) = 1 Then Exit Function
'If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Function
On Error GoTo ErrorHandler

Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim InvRecord  As ADODB.Recordset
Sql = "Select [Menu Item Name] from [MenuItemCountDown] where [Quick Pick] = 0 and [Inventory Amount] <> 9999"
Set InvRecord = GRS.ReturnRecordset(Sql, ReturnAdoConnection("MSaleData"), "MenuItemCountDown")
    If Not InvRecord.EOF Then
       InventoryIsOn = True
    Else
       InventoryIsOn = False
    End If
DbClose Nothing, InvRecord
Set GRS = Nothing
InventoryIsSet = InventoryIsOn
Exit Function
ErrorHandler:
InventoryIsOn = False
'// used for menu count down -- out of stock inventory
End Function
Public Function Return_Item(Menu_Item As Variant, qty As Single) As Single
On Error GoTo ErrorHandler
If AllMenuItems(Menu_Item) Is Nothing Then Exit Function
 If AllMenuItems(Menu_Item).InventoryAmount <> 9999 Then
    AllMenuItems(Menu_Item).InventoryAmount = AllMenuItems(Menu_Item).InventoryAmount + qty
 Else
    Exit Function
 End If
 If Menu_Item <> AllMenuItems(Menu_Item).MenuItemName Then
  Trans_Log Menu_Item & " " & AllMenuItems(Menu_Item).MenuItemName
 End If
 Dim InvMItem   As ADODB.Connection
 Dim GRS        As GeneralLib.Utilities
 Set GRS = New GeneralLib.Utilities
 Dim Rst        As String
 
 Sql = "Update [MenuItemCountDown] Set [Inventory Amount] = [Inventory Amount] + " & qty & " where [Menu Item Name] = '" & Menu_Item & "'"
 Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection("MSaleData"))
 Set GRS = Nothing
Exit Function
ErrorHandler:
Error_Log Err.Number, Err.Description, "Return Item", "None"
End Function
Public Function CountDownQty(MenuItemName As String) As Long
On Error GoTo ErrorHandler
 CountDownQty = 9999
 Dim GRS As GeneralLib.Utilities
 Set GRS = New GeneralLib.Utilities
  CountDownQty = Nz(GRS.CountDownQty(MenuItemName))
 Set GRS = Nothing
Exit Function
ErrorHandler:
Error_Log Err.Number, Err.Description, "CountDownQty", "None"
End Function
Public Function SetCountDownQty(MenuItemName As String, NewCount As Integer) As Long
On Error GoTo ErrorHandler
 Dim Rst        As String
 Dim GRS As GeneralLib.Utilities
 Set GRS = New GeneralLib.Utilities
 Sql = "Update [MenuItemCountDown] Set [Inventory Amount] = " & NewCount & " where [Menu Item Name] = '" & MenuItemName & "'"
 Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection("MSaleData"))
Set GRS = Nothing
Exit Function
ErrorHandler:
Error_Log Err.Number, Err.Description, "setcountdownqty", "None"
End Function
Public Function Check_Inventory(Amount As Long, Menu_Item As String, OrderQty As Single) As Boolean
On Error GoTo ErrorHandler
      Dim InvRecord As ADODB.Recordset
      Dim OnHand    As Integer
      Check_Inventory = False
      
      If InventoryIsSet = False Then Exit Function
      
      If Nz(GetSetting("Micro$ale", "Register Setup", "Common Server Failed")) = 1 Then Exit Function
      
      If UCase(LoggedOnAs(0).EmpLoggedOnAs) = "TRAINING MODE" Then Exit Function
      
      If AllMenuItems(Menu_Item) Is Nothing Then Exit Function
      
      Dim GRS As GeneralLib.Utilities
      Set GRS = New GeneralLib.Utilities
      
       Sql = "Select [Inventory Amount] from [MenuItemCountDown] where [Inventory Amount] <> 9999 and [Menu Item Name] = '" & Menu_Item & "'"
      
      Set InvRecord = GRS.ReturnRecordset(Sql, ReturnAdoConnection("MSaleData"), "MenuItemCountDown")
       If Not InvRecord.EOF Then
        If OrderQty = 0 Then
          If Nz(InvRecord![Inventory Amount]) = 0 Then
           Check_Inventory = True
          End If
          AllMenuItems(Menu_Item).InventoryAmount = Nz(InvRecord![Inventory Amount])
        Else
         OnHand = Nz(InvRecord![Inventory Amount])
         
         Amount = OnHand - OrderQty
         If Amount < 0 Then
          
          Order_Qty = IIf(OnHand < 0, 0, OnHand)
          
          Amount = IIf(Amount < 0, 0, Amount)
          
          See_Msg Order_Qty & " Available", 5, "Order Exceeded Amount Available"
          Check_Inventory = True
         
         Else
          Sql = "Update [MenuItemCountDown] set [Inventory Amount] = [Inventory Amount] - " & CInt(OrderQty)
          Sql = Sql & " Where [Menu Item Name] = '" & Menu_Item & "'"
          
          Dim Rst As String
          
          Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection("MSaleData"))
          
          AllMenuItems(Menu_Item).InventoryAmount = AllMenuItems(Menu_Item).InventoryAmount - OrderQty
         End If
        End If
        AllMenuItems(Menu_Item).InventoryAmount = IIf(AllMenuItems(Menu_Item).InventoryAmount < 0, 0, AllMenuItems(Menu_Item).InventoryAmount)
       End If
       DbClose Nothing, InvRecord '
       Set GRS = Nothing
Exit Function
ErrorHandler:
Error_Log Err.Number, Err.Description, "Check Inventory", "None"
Err.Clear
DbClose Nothing, InvRecord
End Function
Public Function AddExchangeInventory(MenuItem As String, QtySold As Integer) As String
On Error GoTo ErrorHandler
AddExchangeInventory = ""
If LoggedOnAs(0).EmpLoggedOnAs = "TRAINING MODE" Then Exit Function

Dim InvMItem   As ADODB.Connection
Dim InvRecord  As ADODB.Recordset
Dim PLUNO      As String

PLUNO = GetPLUNO(MenuItem)

If PLUNO = "None" Or PLUNO = "" Then Exit Function

Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
 AddExchangeInventory = GL.AddExchangeInventory(MenuItem, PLUNO, QtySold)
Set GL = Nothing

Exit Function
ErrorHandler:
Error_Log Err.Number, Err.Description, "AddExchangeInventory", "None"
AddExchangeInventory = 0
'// used for drink exchange
End Function
Public Function ReturnExChangePrice(PLUNO As String, PriceIn As Currency) As Currency
On Error GoTo ErrorHandler
ReturnExChangePrice = 0
If LoggedOnAs(0).EmpLoggedOnAs = "TRAINING MODE" Then Exit Function

If PLUNO = "None" Or Trim$(PLUNO) = "" Then
  ReturnExChangePrice = PriceIn
  Exit Function
End If

Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
 ReturnExChangePrice = Val(GL.ReturnExChangePrice(PLUNO, PriceIn))
Set GL = Nothing
Exit Function
ErrorHandler:
Error_Log Err.Number, Err.Description, "ReturnExChangePrice", "None"
ReturnExChangePrice = PriceIn
'// used with drink exchange
End Function
Public Function X10HowManyPlayer(ByVal TableNo As String, ByVal MyCheck As Long) As Integer
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
 X10HowManyPlayer = GL.X10HowManyPlayer(TableNo, MyCheck)
Set GL = Nothing
Exit Function
End Function
Public Sub X10EndPlay(ByVal ChkNO As Long, ByVal SelectionNo As Integer)
On Error GoTo ErrorHandler
Dim Response    As String
Dim GL          As GeneralLib.Utilities

Set GL = New GeneralLib.Utilities
Response = GL.X10EndPlay(ChkNO, SelectionNo)
Set GL = Nothing
Exit Sub
ErrorHandler:
Error_Log Err.Number, Err.Description, "X10EndPlay", "None"
Err.Clear
'// used for table rentals or any other item that pays by time
End Sub
Public Function ElapsedTime(ByVal tStart As Date, ByVal tStop As Date) As String
On Error GoTo ELAPSEDTIME_ERROR

    Dim totalSeconds As Long
    Dim minutes As Long
    Dim seconds As Long

    ' If stop time is earlier than start, assume it crossed midnight
    If tStop < tStart Then
        tStop = DateAdd("d", 1, tStop)
    End If

    ' Get total seconds between start and stop
    totalSeconds = DateDiff("s", tStart, tStop)

    ' Break into minutes and seconds
    minutes = totalSeconds \ 60          ' integer division
    seconds = totalSeconds Mod 60

    ' Format as M.SS  (e.g., 3.07, 12.45, 135.09)
    ElapsedTime = Format$(minutes, "0") & "." & Format$(seconds, "00")
    Exit Function

ELAPSEDTIME_ERROR:
    ' In case of bad input, return blank (or whatever you prefer)
    ElapsedTime = ""
End Function
Public Function ElapsedMinutes(tStart As Variant, tStop As Variant) As Long
On Error GoTo END_ELAPSEDTIME
Dim dtr, dtl, jml As Long
Dim EH As Long
Dim EM As Long
Dim ES As Long
dtl = (Hour(tStart) * 3600) + (Minute(tStart) * 60) + (Second(tStart))
dtr = (Hour(tStop) * 3600) + (Minute(tStop) * 60) + (Second(tStop))
If tStop < tStart Then
   jml = 86400
Else
   jml = 0
End If
jml = jml + (dtr - dtl)
EH = (Int((Int((jml / 3600)) Mod 24))) * 60
EM = Int((Int((jml / 60)) Mod 60))
ElapsedMinutes = EH + EM
If ElapsedMinutes < 0 Then ElapsedMinutes = 1440
END_ELAPSEDTIME:
End Function
Public Function Table_Time(OpenTime As Date, OpenDate As Date, tn As Integer) As String
On Error GoTo ESub
Dim OpenHours As Single
Dim TablTime  As Long
Dim OpMins    As Long
Dim HRS       As String
Dim TableTime As Long
Dim OpenMins  As Long
 TableTime = DateDiff("n", OpenDate & " " & OpenTime, Now)
 
 'If TableTime > Val(Input_Option(2)) Then Tbl_Times_Up(tn) = -1
 If TableTime > 0 Then OpenHours = (TableTime / 60)
 If (OpenHours) < 1 Then
    OpenMins = (OpenHours * 60)
    OpenHours = 0
 Else
    OpMins = Fix(OpenHours)
    OpenMins = Abs(TableTime - (OpMins * 60))
    OpenHours = OpMins
 End If
 If OpenHours > 0 Then HRS = " Hrs "
 Table_Time = Format(OpenHours, "## ") & HRS & Format(OpenMins, "00") & " Mins"
ESub:
End Function

Public Function Get_Age(ByVal Birthdate As Date) As String

        'Don't set second parameter if you want Age as of today

        'Demo 1: get age of person born 2/11/1954
        'Dim objDate As New System.DateTime(1954, 2, 11)
        'Debug.WriteLine(GetAge(objDate))

        'Demo 1: get same person's age 10 years from now
        'Dim objDate As New System.DateTime(1954, 2, 11)
        'Dim objdate2 As System.DateTime
        'objdate2 = Now.AddYears(10)
        'Debug.WriteLine(GetAge(objDate, objdate2))

        Dim iMonths     As Integer
        Dim iYears      As Integer
        Dim dYears      As Single
        Dim lDayOfBirth As Long
        Dim lAsOf       As Long
        Dim iBirthMonth As Integer
        Dim iAsOFMonth  As Integer

        lDayOfBirth = DatePart("d", Birthdate)

        iBirthMonth = DatePart("m", Birthdate)
        iAsOFMonth = DatePart("m", Now)

        iYears = DateDiff("yyyy", Birthdate, Now)
        If iBirthMonth = iAsOFMonth Then
            If DatePart("d", Now) < lDayOfBirth Then
                iYears = iYears - 1
            End If
        Else
            If DatePart("M", Now) < iBirthMonth Then
                iYears = iYears - 1
            End If
        End If

        Get_Age = iYears
    End Function
Public Function ReturnAge() As Long
On Error GoTo ERH
Dim EnterAge    As String
Dim UserName    As String
Dim m           As String
Dim D           As String
Dim Y           As String
Dim s           As Integer
ReturnAge = 0
EnterAge = ""
EnterAge = ShowNumberPad("Enter Birth Date mm.dd.yy", False, 0, 0)
If EnterAge = "0" Or EnterAge = "" Then Exit Function
If InStr(EnterAge, ".") > 0 Then
ParseD:
    m = Left$(EnterAge, InStr(EnterAge, ".") - 1)
    s = InStr(EnterAge, ".")
    D = Mid$(EnterAge, s + 1, InStr(EnterAge, ".") - 1)
    s = InStr(s + 1, EnterAge, ".")
    Y = Mid$(EnterAge, s + 1, Len(EnterAge))
    EnterAge = m & "/" & D & "/" & Y
End If
If Not IsDate(EnterAge) Then
 Dim GetInfo() As String
 Dim ReturnData As String
 Dim GD As MSaleUtility.Utilities
 Set GD = New MSaleUtility.Utilities
     ReturnData = GD.ParseDriverLicense(EnterAge)
 Set GD = Nothing
 GetInfo = Split(ReturnData, "|")
 EnterAge = Format(GetInfo(0), "MM/DD/YYYY")
 UserName = "" 'GetInfo(1)
End If
If Not IsDate(EnterAge) Then
  See_Msg "Not A Valid Date", 5, "Must Enter MM.DD.YY"
  Exit Function
End If
Dim AllData$
If IsDate(EnterAge) Then
  Dim CAge As Long
  CAge = Get_Age(EnterAge)
  ReturnAge = CAge
  If See_Msg("Age Is " & CAge, 4, "Print Slip") = 6 Then
   GetSpoolPrinter (AllPrinter.Item("Receipt").PrinterName)
   EC(7) = vbCrLf
   AllData$ = EC(0)  'Int printer
   AllData$ = AllData$ & EC(1) & EC(3) & "Current Time " & Now & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & UserName & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & "Age Verified    : " & EnterAge & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & "Customer Age Is : " & CAge & "" & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & "Check By        : " & LoggedOnAs(0).EMPName & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & String(40, 45) & EC(2) & EC(7)
   AllData$ = AllData$ & EC(6) & EC(5)
   PrintChitSlip AllData$
  End If
End If
Exit Function
ERH:
End Function
Public Function ReturnDriverInfo(RData As DriverScanData) As String
On Error GoTo ERH
Dim LicenseData()   As String
Dim m               As String
Dim D               As String
Dim Y               As String
Dim s               As Integer
Dim EnterAge        As String
ReturnDriverInfo = ""
EnterAge = ShowNumberPad("Scan Drivers License", False, 0, 0)
If EnterAge = "" Then Exit Function
 LicenseData = Split(EnterAge, ",")
 RData.C_Name = LicenseData(0) & " " & LicenseData(2)
 RData.DLN = LicenseData(3)
 RData.DOB = LicenseData(4)
 If Not IsDate(RData.DOB) Then
   If See_Msg("Error Reading License", 4, "Enter By Hand?") = 6 Then
         RData.C_Name = ShowKeyBoard("Enter Name")
         RData.DLN = ShowKeyBoard("Enter Driver License No")
         EnterAge = ShowNumberPad("Enter Birth Date mm.dd.yy", False, 0, 0)
         If EnterAge = "0" Or EnterAge = "" Then Exit Function
         If InStr(EnterAge, ".") > 0 Then
            m = Left(EnterAge, InStr(EnterAge, ".") - 1)
            s = InStr(EnterAge, ".")
            D = Mid(EnterAge, s + 1, InStr(EnterAge, ".") - 1)
            s = InStr(s + 1, EnterAge, ".")
            Y = Mid(EnterAge, s + 1, Len(EnterAge))
            RData.DOB = m & "/" & D & "/" & Y
        End If
   Else
     Exit Function
   End If
 End If
Dovr:
 RData.RegNo = ShowKeyBoard("Enter Registration Number")
 If RData.RegNo = "" Then
    See_Msg "Required Item", 5, ""
    GoTo Dovr
 End If
 ReturnDriverInfo = "Success"
 Dim CAge As Long
 Dim AllData$
 CAge = Get_Age(RData.DOB)
 If See_Msg("Age Is " & CAge, 4, "Print Slip") = 6 Then
   GetSpoolPrinter (AllPrinter.Item("Receipt").PrinterName)
   EC(7) = vbCrLf
   AllData$ = EC(0)  'Int printer
   AllData$ = AllData$ & EC(1) & EC(3) & "Current Time " & Now & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & RData.C_Name & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & "Age Verified    : " & RData.DOB & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & "Customer Age Is : " & CAge & "" & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & "Check By        : " & LoggedOnAs(0).EMPName & EC(2) & EC(7)
   AllData$ = AllData$ & EC(1) & EC(3) & String(40, 45) & EC(2) & EC(7)
   AllData$ = AllData$ & EC(6) & EC(5)
   PrintChitSlip AllData$
End If
Exit Function
ERH:
End Function
Public Function ReturnQSRDept(LineNo As Integer, f As Integer) As Long
Dim Items() As String
   If InStr(Order.Ticket(0).TextMatrix(LineNo, f), "QSRDept") > 0 Or InStr(Order.Ticket(0).TextMatrix(LineNo, f), "QSRRoute") > 0 Then
      Items = Split(Order.Ticket(0).TextMatrix(LineNo, f), "-")
      ReturnQSRDept = Val(Items(1))
      Exit Function
    End If
End Function
Public Function ReturnQsrDept64(NIP As Integer, LineNo As Integer, f As Integer) As Long
 Dim Items() As String
 If InStr(Grid_array(NIP, LineNo, f), "QSRDept") > 0 Or InStr(Grid_array(NIP, LineNo, f), "QSRRoute") > 0 Then
      Items = Split(Grid_array(NIP, LineNo, f), "-")
      ReturnQsrDept64 = Val(Items(1))
      Exit Function
    End If
End Function
Public Function GetMonitor(MonitorName As String) As String
On Error GoTo ErrorHandler
If IsQSROn Then
  If InStr(MonitorName, "Monitor") > 0 Then
    MonitorName = "QSRDept-" & Right$(MonitorName, 1)
  End If
End If
If MonitorName = "" Or MonitorName = "None" Then
 MonitorName = "None"
 GetMonitor = MonitorName
 Exit Function
End If
If MonitorName = "Pass" Then
 GetMonitor = MonitorName
 Exit Function
End If
If RouteVideo Then
  If GetSetting("Micro$ale", "Video Routing", MonitorName) > "" Then
    If GetSetting("Micro$ale", "Video Routing", MonitorName) <> "None" Then
     GetMonitor = GetSetting("Micro$ale", "Video Routing", MonitorName)
     Exit Function
    End If
  End If
End If
If GetSetting("Micro$ale", "Printer Routing", MonitorName) > "" And Check_Option(45) = 0 Then
 If GetSetting("Micro$ale", "Printer Routing Time", MonitorName) > "" Then
     If OKTOROUTE(GetSetting("Micro$ale", "Printer Routing Time", MonitorName)) Then
       GetMonitor = GetSetting("Micro$ale", "Printer Routing", MonitorName) & ""
       Exit Function
     Else
       GetMonitor = MonitorName
     End If
 Else
  GetMonitor = GetSetting("Micro$ale", "Printer Routing", MonitorName) & ""
  Exit Function
 End If
Else
  GetMonitor = MonitorName
End If
If Check_Option(78) = 0 Then Exit Function
'// option 78 0 is disable kitchen Video 1 = Video On
If OrderIs = GetOrderType(GoSale) Then
  If Not AllPrinter.Item(MonitorName) Is Nothing Then
    If AllPrinter.Item(CStr(MonitorName)).GoOrder > "" And AllPrinter.Item(CStr(MonitorName)).GoOrder <> "None" Then
     GetMonitor = AllPrinter.Item(CStr(MonitorName)).GoOrder
     Exit Function
    End If
  End If
End If
If InStr(UCase(MonitorName), UCase("MONITOR")) > 0 Then
  If Not AllPrinter.Item(MonitorName) Is Nothing Then
    GetMonitor = AllPrinter.Item(CStr(MonitorName)).PrinterPort
  End If
End If
Exit Function
ErrorHandler:
GetMonitor = MonitorName & ""
Trans_Log Err.Description & " Get Monitor " & MonitorName
End Function
Public Function ReturnOTRate(RegRate As Currency) As Currency
Dim TW As Currency
Dim Rw As Currency
If Val(Input_Option(62)) = 0 Then   '// 62 /tip employee min wage
   ReturnOTRate = (RegRate * 1.5)
   Exit Function
End If
If Val(Input_Option(61)) = 0 Then   '// 61 non tip employee min wage
   Rw = 7.25
Else
   Rw = Val(Input_Option(61))
End If
If Val(Input_Option(62)) > 0 Then
   TW = Val(Input_Option(62))  'Tipped Wages
End If
If RegRate * 1.5 >= (Rw * 1.5) Then
  ReturnOTRate = (RegRate * 1.5)
Else
  ReturnOTRate = ((ServerOverTime(RegRate)))
End If
End Function
Public Property Get CalculateOT(OTHours As Single, RegRate As Currency) As Currency
Dim TW As Currency
Dim Rw As Currency
If OTHours = 0 Then
  CalculateOT = 0
  Exit Property
End If
If Val(Input_Option(61)) = 0 Then
   Rw = 7.25
Else
   Rw = Val(Input_Option(61))
End If
If Val(Input_Option(62)) = 0 Then
   TW = 2.13
Else
   TW = Val(Input_Option(62))  'Tipped Wages
End If
If RegRate * 1.5 >= (Rw * 1.5) Then
  CalculateOT = (OTHours * (RegRate * 0.5))
Else
  CalculateOT = OTHours * ((ServerOverTime(RegRate) - TW))
End If
End Property
Public Function CustomerIsWaiting(CustomerCheckNo As Long) As String
Dim CellNumber As String
Dim Response   As String
Dim Db         As ADODB.Connection
Dim RS         As ADODB.Recordset
If Not AllOrders Is Nothing Then
  CellNumber = AllOrders.Item(1).CallerPhone
Else
RedoEntry:
  CellNumber = ShowNumberPad("Enter Cell Number", True, 0, 0)
End If
If CellNumber = "" Then
  Call See_Msg("Skipped", 5, "no number saved..")
  Exit Function
End If
If See_Msg(CellNumber, 4, "Is This Correct?") <> 6 Then
   GoTo RedoEntry
End If
Response = SendTextMessage("Confirmation: Your Order Is In " & LocationName, CellNumber)
If Response = "Success" Then
  If ConnectADODB("MSaleData", Db) Then
    Sql = "Select * from [CustomerIsWaiting] where [CheckNumber] = " & CustomerCheckNo
   If ConnectADORS(Sql, Db, RS) = False Then
     See_Msg "Error Connecting To Database", 5, ""
     Exit Function
   End If
   If RS.EOF Then RS.AddNew
   RS.fields("CustomerPhone").Value = CellNumber
   RS.fields("DateTime").Value = Now
   RS.fields("CheckNumber").Value = CustomerCheckNo
   RS.Update
   DbClose Db, RS
  End If
Else
   'See_Msg Response, 5, "Text Failed"
End If
End Function
Public Sub NotifyCustomerWaiting(CustomerCheckNo As Long)
Dim CellNumber      As String
Dim Response        As String
'Dim CustomerCheckNo As Long
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
CellNumber = ""
If CustomerCheckNo = 0 Then
 CustomerCheckNo = Val(ShowNumberPad("Enter Check Number", True, 0, 0))
End If
If CustomerCheckNo = 0 Then Exit Sub
If ConnectADODB("MSaleData", Db) Then
    Sql = "Select * from [CustomerIsWaiting] where [CheckNumber] = " & CustomerCheckNo
   If ConnectADORS(Sql, Db, RS) = False Then
     See_Msg "Error Connecting To Database", 5, ""
     Exit Sub
   End If
   If Not RS.EOF Then
    CellNumber = RS.fields("CustomerPhone").Value
   Else
    See_Msg "Check Not Found", 5, ""
    DbClose Db, RS
    Exit Sub
   End If
   Response = SendTextMessage("Confirmation: Your Order Is Ready. Check No " & CustomerCheckNo & " " & LocationName, CellNumber)
   If Response = "Success" Then
     Sql = "delete from [CustomerIsWaiting] where [CheckNumber] = " & CustomerCheckNo
     Db.Execute Sql
   Else
   
   End If
    DbClose Db, RS
Else
   See_Msg "Failed to connect to database", 5, "Text Failed"
End If
End Sub

Public Function SendTextMessage(MessageTo As String, PHONENO As String) As String   ', MediaUrl As String
      If SiteLic.Twilio = 0 Then
         See_Msg "No License To Text", 5, "Contact dealer for texting support"
         Exit Function
      End If
      SendTextMessage = ""
      Dim TxtResponse As String
      Dim GL As GeneralLib.Utilities
      Set GL = New GeneralLib.Utilities
          TxtResponse = GL.SMSText(MessageTo & vbCrLf, PHONENO, "")  'MediaUrl
      Set GL = Nothing
      Dim xmls As MPxmlParse
      Set xmls = New MPxmlParse
      xmls.xml = TxtResponse
      TxtResponse = xmls.token("Sid")
      Set xmls = Nothing
      If TxtResponse > "" Then
         See_Msg "Success", 5, "Confirmed."
         SendTextMessage = "Success"
      Else
         See_Msg "Text Failed", 0, "Is Phone Correct " & PHONENO
      End If
End Function
Public Property Get ServerOverTime(OTRate As Currency) As Currency
Dim Rw As Currency  'Regular wage
Dim TW As Currency  'Tip Wage
Dim TC As Currency  'Tip Credit
If OTRate = 0 Then
  ServerOverTime = 0
  Exit Property
End If
If Val(Input_Option(61)) = 0 Then
   Rw = 7.25
Else
   Rw = Val(Input_Option(61))  'Regular wage
End If
If Val(Input_Option(62)) = 0 Then
   TW = 2.13
Else
   TW = Val(Input_Option(62))  'Tipped Wages
End If
TC = Rw - TW
ServerOverTime = IIf(OTRate > (Rw * 1.5) - TC, OTRate, (Rw * 1.5) - TC)
End Property
Public Function GetAutoDiscount(MItem As String, MCat As String, MDept As String) As String
GetAutoDiscount = ""
If DiscList Is Nothing Then Exit Function
If ADiscs.count = 0 Then Exit Function
Dim FoundCoupon As String
Dim LI          As Long
On Error GoTo Er
FoundCoupon = ""
For LI = 1 To ADiscs.count
 If Not ADiscs.Item(LI) Is Nothing Then
  If Not DiscList.Item(ADiscs.Item(LI).DiscountName & MItem) Is Nothing Then
     FoundCoupon = GetAutoCoupon(ADiscs.Item(LI).DiscountName)
  ElseIf Not DiscList.Item(ADiscs.Item(LI).DiscountName & MCat) Is Nothing Then
     FoundCoupon = GetAutoCoupon(ADiscs.Item(LI).DiscountName)
  ElseIf Not DiscList.Item(ADiscs.Item(LI).DiscountName & MDept) Is Nothing Then
     FoundCoupon = GetAutoCoupon(ADiscs.Item(LI).DiscountName)
  End If
  If FoundCoupon > "" Then
   GetAutoDiscount = FoundCoupon
   Exit Function
  End If
 End If
Next LI
Er:

End Function
Public Function GetAutoCoupon(CouponName As String) As String
GetAutoCoupon = ""
If ADiscs Is Nothing Then Exit Function
If Not ADiscs.Item(CouponName) Is Nothing Then
   Select Case DatePart("w", Date, vbMonday)
      Case Is = 1
          If ADiscs.Item(CouponName).Mon = 1 Then Exit Function
      Case Is = 2
          If ADiscs.Item(CouponName).Tue = 1 Then Exit Function
      Case Is = 3
          If ADiscs.Item(CouponName).Wed = 1 Then Exit Function
      Case Is = 4
          If ADiscs.Item(CouponName).Thur = 1 Then Exit Function
      Case Is = 5
          If ADiscs.Item(CouponName).Fri = 1 Then Exit Function
      Case Is = 6
          If ADiscs.Item(CouponName).Sat = 1 Then Exit Function
      Case Is = 7
          If ADiscs.Item(CouponName).Sun = 1 Then Exit Function
   End Select
   If UCase(ADiscs.Item(CouponName).StartTime) = "12:00 AM" And UCase(ADiscs.Item(CouponName).EndTime) = "12:00 AM" Then
     GetAutoCoupon = ADiscs.Item(CouponName).DiscountName
     Exit Function
   Else
     If IsDate(ADiscs.Item(CouponName).StartTime) And IsDate(ADiscs.Item(CouponName).EndTime) Then
       If IsValidTime(Time, CDate(ADiscs.Item(CouponName).StartTime), CDate(ADiscs.Item(CouponName).EndTime), -1) Then
          GetAutoCoupon = ADiscs.Item(CouponName).DiscountName
       Else
          GetAutoCoupon = ""
       End If
       Exit Function
     End If
   End If
Else
  GetAutoCoupon = ""
End If
End Function
Public Function IsValidCoupon(CouponName As String) As Boolean
IsValidCoupon = True
If ADiscs Is Nothing Then Exit Function
If Not ADiscs.Item(CouponName) Is Nothing Then
   Select Case DatePart("w", Date, vbMonday)
      Case Is = 1
          If ADiscs.Item(CouponName).Mon = -1 Then Exit Function
      Case Is = 2
          If ADiscs.Item(CouponName).Tue = -1 Then Exit Function
      Case Is = 3
          If ADiscs.Item(CouponName).Wed = -1 Then Exit Function
      Case Is = 4
          If ADiscs.Item(CouponName).Thur = -1 Then Exit Function
      Case Is = 5
          If ADiscs.Item(CouponName).Fri = -1 Then Exit Function
      Case Is = 6
          If ADiscs.Item(CouponName).Sat = -1 Then Exit Function
      Case Is = 7
          If ADiscs.Item(CouponName).Sun = -1 Then Exit Function
   End Select
   If UCase(ADiscs.Item(CouponName).StartTime) = "12:00 AM" And UCase(ADiscs.Item(CouponName).EndTime) = "12:00 AM" Then
     IsValidCoupon = True
     Exit Function
   Else
     If IsDate(ADiscs.Item(CouponName).StartTime) And IsDate(ADiscs.Item(CouponName).EndTime) Then
       If IsValidTime(Time, CDate(ADiscs.Item(CouponName).StartTime), CDate(ADiscs.Item(CouponName).EndTime), -1) Then
          IsValidCoupon = True
       Else
          IsValidCoupon = False
       End If
       Exit Function
     End If
   End If
Else
  IsValidCoupon = True
End If
End Function
Public Function PrePaidCharge() As Boolean
PrePaidCharge = False
Dim PrePaidAmount As Currency
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
  PrePaidAmount = GL.GetPrePaidAmount(LoggedOnAs(0).EMPName, Check_no)
Set GL = Nothing
If PrePaidAmount >= Chk_Sales_data(5) Then
  PrePaidCharge = True
End If
End Function
Public Function GetTipwitheld(ReportDate As String) As Currency
    Dim CCDB As ADODB.Connection
    Dim CCRS As ADODB.Recordset
    ConnectADODB "MSaleData", CCDB
    Sql = "Select Distinct [Date], Sum([Tips Withheld]) as TTips from [Tips Withheld]"
    Sql = Sql & " Where TRY_CAST([Date] AS date) = '" & CDate(ReportDate) & "' Group by [Date]"
    ConnectADORS Sql, CCDB, CCRS
    If Not CCRS.EOF Then
      GetTipwitheld = Val(CCRS!TTips & "")
    Else
      GetTipwitheld = 0
    End If
    DbClose CCDB, CCRS
End Function
Public Function GetAllPoolTips(InTime As String, EndTime As String, DayofWeek As Integer) As Currency
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
GetAllPoolTips = GL.GetAllPooledTips(InTime, EndTime, DayofWeek)
Set GL = Nothing
End Function
Public Function GetMissedTips(SvrName As String) As Integer
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
  GetMissedTips = GL.CountMissedTips(SvrName)
Set GL = Nothing
End Function
Public Function GetPDAUser() As String
On Error Resume Next
GetPDAUser = ""
Dim ObjNet As Object
Set ObjNet = CreateObject("WScript.Network")
GetPDAUser = ObjNet.UserName
Set ObjNet = Nothing
End Function
Public Function GetCommonPathName(InPath As String) As String
If InPath = "" Then
   GetCommonPathName = MyRegisterName
   Exit Function
End If
If InStr(InPath, "\\") > 0 Then
  GetCommonPathName = Mid(InPath, 3, InStr(3, InPath, "\") - 3)
  Exit Function
End If
GetCommonPathName = MyRegisterName
End Function
Public Function ReturnTxtFromFile(FileName As String) As Variant
On Error GoTo ERH
Dim FN As Object
Set FN = CreateObject("Scripting.FileSystemObject")
Dim RD As Variant
Dim D  As String
If FN.FileExists(FileName) = True Then
    Set RD = FN.OpenTextFile(FileName, 1, False, 0)
      ReturnTxtFromFile = RD.ReadLine
      If ReturnTxtFromFile > "" Then
        If Left(ReturnTxtFromFile, 1) = """" Then
          ReturnTxtFromFile = Mid(ReturnTxtFromFile, 2, Len(ReturnTxtFromFile) - 2)
        End If
      End If
    RD.Close
Else
    ReturnTxtFromFile = ""
End If
Set FN = Nothing
Set RD = Nothing
Exit Function
ERH:
ReturnTxtFromFile = ""
End Function
'*****************************************************************************************
' Function Name: GetSSN
' Description: Retrieves the Social Security Number (SSN) of an employee based on their first and last name.
' Parameters:
'   - Employee: String containing the full name of the employee in the format "First Name Last Name".
' Returns:
'   - String: Social Security Number (SSN) of the employee if found, otherwise returns "EMP Not Found".
'*****************************************************************************************
Public Function GetSSN(Employee As String) As String
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim Db As ADODB.Connection
    Dim Tb As ADODB.Recordset
    Dim FN As Integer
    Dim First_Name As String
    Dim Last_Name As String
    
    ' Find the position of the space in the Employee string
    FN = InStr(Employee, " ")
    
    ' Extract the first name and last name from the Employee string
    First_Name = Trim$(Left$(Employee, FN - 1))
    Last_Name = Trim$(Mid$(Employee, FN + 1, Len(Employee) - FN))
    
    ' Connect to the database
    ConnectADODB "MSaleData", Db
    
    ' Construct SQL query to retrieve employee data based on first and last name
    Sql = "Select * from [General data] where "
    Sql = Sql & " [First Name] = '" & First_Name & "'"
    Sql = Sql & " and [Last Name] = '" & Last_Name & "'"
    
    ' Execute SQL query
    ConnectADORS Sql, Db, Tb
    
    ' Check if record exists
    If Not Tb.EOF Then
        ' Return the Social Security Number (SSN) of the employee
        GetSSN = Tb.fields("Company Id").Value & ""
    Else
        ' Return "EMP Not Found" if employee not found
        GetSSN = "EMP Not Found"
    End If
    
    ' Close database connection
    DbClose Db, Tb
    
    ' Exit function
    Exit Function

ErrorHandler:
    ' Call See_Msg function to display error message
    See_Msg Err.Description, 5, "Cannot Get SSN"
End Function

Public Function ParseData(InData As String, ReturnValues() As Variant) As Boolean
Dim VarL1 As Long
Dim VarL2 As Integer
Dim Counter As Long
Dim PIS     As Long
Counter = -1
PIS = 1
ReDim ReturnValues(0)
If Len(InData) > 0 Then
  ParseData = True
  Do While PIS < Len(InData)
      Counter = Counter + 1
      ReDim Preserve ReturnValues(Counter)
      ReturnValues(Counter) = ""
      For VarL2 = PIS To Len(InData)
          If Mid(InData, VarL2, 1) = "," Or Asc(Mid(InData, VarL2, 1)) = 13 Then Exit For
          If Mid(InData, VarL2, 1) <> """" And Asc(Mid(InData, VarL2, 1)) <> 10 Then
             ReturnValues(Counter) = ReturnValues(Counter) & Mid(InData, VarL2, 1)
          End If
      Next VarL2
    PIS = VarL2 + 1
  Loop
Else
  ParseData = False
End If
End Function
Public Function ReturnAllFileData(FileName As String) As Variant
On Error GoTo ERH
Dim FN As Object
Dim RD As Object
Set FN = CreateObject("Scripting.FileSystemObject")
Dim D  As String
If FN.FileExists(FileName) = True Then
    Set RD = FN.OpenTextFile(FileName, 1, -2)
      ReturnAllFileData = RD.ReadAll
    RD.Close
Else
    ReturnAllFileData = ""
End If
Set FN = Nothing
Set RD = Nothing
Exit Function
ERH:
'ReturnTxtFromFile = ""
End Function
Public Function ReturnDNS(DbaseName As String) As String
On Error GoTo ERH
Dim MySql As String
Dim Cdb As ADODB.Connection
Dim CRS As ADODB.Recordset
ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Cdb
   MySql = "Select * from [DNS] where [DNS Label] = '" & Trim(DbaseName) & "'"
ConnectADORS MySql, Cdb, CRS
If Not CRS.EOF Then
     ReturnDNS = CRS![DNS Name] & ""
Else
     ReturnDNS = ""
End If
DbClose Cdb, CRS
Exit Function
ERH:
ReturnDNS = ""
End Function
Public Function GetUnPaidTips(ServerName As String, CashierName As String) As Currency
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
        ConnectADODB DebitDb, Db
        Sql = "Select * from [Debit] Where [Close Out day] = 0 and [Server Name] = '" & ServerName & "'"
        If Input_Option(29) = "Global" Then
        
        ElseIf Input_Option(29) = "By Cashier" Then
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
        Else
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
        End If
        Sql = Sql & " and [Tip Paid] = 0"
    
        ConnectADORS Sql, Db, RS
        GetUnPaidTips = 0
        Do While Not RS.EOF
          GetUnPaidTips = GetUnPaidTips + RS.fields("Tip").Value - CalChargeBack(RS.fields("Account Name").Value, RS.fields("Tip").Value)
         RS.MoveNext
        Loop
    DbClose Db, RS
    Sql = "Select Distinct [Server Name], Sum([Tip Amount]) as TotalTip from [DayPayments] Where [Server Name] = '" & ServerName & "'"
    
    If Input_Option(29) = "Global" Then
        
    ElseIf Input_Option(29) = "By Cashier" Then
      Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
    Else
      Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
    End If
    Sql = Sql & " and [Tip Paid] = 0"
    Sql = Sql & " Group by [Server Name] "
    If ConnectADODB(FinancialDb, Db) = False Then Exit Function
    If ConnectADORS(Sql, Db, RS) = False Then Exit Function
    If Not RS.EOF Then
     GetUnPaidTips = GetUnPaidTips + Val(RS![TotalTip] & "")
    End If
    DbClose Db, RS
End Function
Public Function GetServerTips(ServerName As String, ShiftDate As String, ShiftCheckNo As Long, ShiftTimeOn As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
  GetServerTips = 0
  ConnectADODB DebitDb, Db
   Sql = "Select Distinct [Server Name], Sum([tip]) as TotalTip from [Debit] Where [Close Out day] = 0 and [Server Name] = '" & ServerName & "'"
   Sql = Sql & " and [Check Number] >= " & ShiftCheckNo
     If IsDate(ShiftTimeOn) Then
       Sql = Sql & " and  TRY_CAST([Date] AS datetime) + TRY_CAST([time] AS datetime) >= '" & CDate(ShiftDate & " " & ShiftTimeOn) & "'"
     End If
   Sql = Sql & " Group By [Server Name]"
  ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
    GetServerTips = Val(RS![TotalTip] & "")
   End If
   DbClose Db, RS
   '// search for credit card tips
   Sql = "Select Distinct [Server Name], Sum([Tip Amount]) as TotalTip from [DayPayments] Where [Server Name] = '" & ServerName & "'"
   Sql = Sql & " and [Check Number] >= " & ShiftCheckNo
   If IsDate(ShiftTimeOn) Then
       Sql = Sql & " and TRY_CAST([Recorded Time] AS datetime) > '" & ShiftTimeOn & "'"
   End If
   Sql = Sql & " Group by [Server Name] "
  If ConnectADODB(FinancialDb, Db) = False Then Exit Function
  If ConnectADORS(Sql, Db, RS) = False Then Exit Function
  If Not RS.EOF Then
   GetServerTips = GetServerTips + Val(RS![TotalTip] & "")
  End If
  DbClose Db, RS
  '// search for non credit card tips -- house accounts -- gift card
  If GetServerTips = 0 Then ' And Not IsDate(ShiftTimeOn) Then
    ConnectADODB "MSaleData", Db
    Sql = "Select * from [Tip Income] where [Server Name] = '" & ServerName & "'"
    Sql = Sql & " and TRY_CAST([Date] AS date) = '" & CDate(ShiftDate) & "'"
    Sql = Sql & " and [Shift] = " & ShiftCheckNo
    ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
     GetServerTips = Val(RS![Charged Tips] & "")
    End If
    DbClose Db, RS
  End If
End Function
Public Function GetPaidTips(ServerName As String, CashierName As String) As Currency
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
        ConnectADODB DebitDb, Db
        Sql = "Select * from [Debit] Where [Close Out day] = 0 and [Server Name] = '" & ServerName & "'"

        If Input_Option(29) = "Global" Then
        
        ElseIf Input_Option(29) = "By Cashier" Then
           Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
        Else
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
        End If
        Sql = Sql & " and [Tip Paid] = 1"
        ConnectADORS Sql, Db, RS
          Do While Not RS.EOF
            GetPaidTips = GetPaidTips + (RS.fields("Tip").Value - CalChargeBack(RS.fields("Account Name").Value, RS.fields("Tip").Value))
            RS.MoveNext
          Loop
    DbClose Db, RS
    
    Sql = "Select Distinct [Server Name], Sum([Tip Amount]) as TotalTip from [DayPayments] Where  [Server Name] = '" & ServerName & "'"
    If Input_Option(29) = "Global" Then
        
    ElseIf Input_Option(29) = "By Cashier" Then
      Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
    Else
      Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
    End If
    Sql = Sql & " and [Tip Paid] = 1"
    Sql = Sql & " Group By [Server Name]"
    If ConnectADODB(FinancialDb, Db) = False Then Exit Function
    If ConnectADORS(Sql, Db, RS) = False Then Exit Function
    If Not RS.EOF Then
     GetPaidTips = GetPaidTips + Val(RS![TotalTip] & "")
    End If
    DbClose Db, RS
End Function
Public Sub AssignCashier(ServerName As String, CashierName As String, CheckNo As Long) '// ADDED FOR LOUIS 01/24/2024
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    ConnectADODB DebitDb, Db
       Sql = "Update [Debit] Set [Cashier Name] = '" & CashierName & "' Where [Check Number] = " & CheckNo & " and [Close Out day] = 0 and [Server Name] = '" & ServerName & "'"
    Db.Execute Sql
    DbClose Db, RS
    Sql = "Update [DayPayments] set [Cashier Name] = '" & CashierName & "' Where [Check Number] = " & CheckNo & " and  [Server Name] = '" & ServerName & "'"
    If ConnectADODB(FinancialDb, Db) = False Then Exit Sub
    Db.Execute Sql
    DbClose Db, RS
    Trans_Log "Assign Cashier " & CashierName & " " & CheckNo
End Sub
Public Sub MarkTipPaid(ServerName As String, CashierName As String)
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
        ConnectADODB DebitDb, Db
        Sql = "Update [Debit] Set [Tip Paid] = 1 Where [Close Out day] = 0 and [Server Name] = '" & ServerName & "'"
        If Input_Option(29) = "Global" Then
        
        ElseIf Input_Option(29) = "By Cashier" Then
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
        Else
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
        End If
        Sql = Sql & " and [Tip Paid] = 0"
    Db.Execute Sql
    DbClose Db, RS
    
    Sql = "Update [Payments] set [Tip Paid] = 1 Where  [Server Name] = '" & ServerName & "'"
    If Input_Option(29) = "Global" Then
    
    ElseIf Input_Option(29) = "By Cashier" Then
      Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
    Else
      Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
    End If
    Sql = Sql & " and [Tip Paid] = 0"
    If ConnectADODB(FinancialDb, Db) = False Then Exit Sub
    Db.Execute Sql
    DbClose Db, RS
End Sub
Public Sub CloseAllMarked(CashierName As String)
On Error GoTo ERH
    Dim Db As ADODB.Connection
     ConnectADODB DebitDb, Db
        Sql = "Update [Debit] Set [Tip Paid] = 1 Where [Close Out day] = 0 "
        
        If Input_Option(29) = "Global" Then
        
        ElseIf Input_Option(29) = "By Cashier" Then
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
        Else
          Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
        End If
        Sql = Sql & " and [Tip Paid] = 0"
    Db.Execute Sql
    SetNothing Db
    If ConnectADODB(FinancialDb, Db) = False Then Exit Sub
        Sql = "Update [Payments] set [Tip Paid] = 1 Where [Tip Paid] = 0"
        If Input_Option(29) = "Global" Then
        
        ElseIf Input_Option(29) = "By Cashier" Then
         Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
        Else
         Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
        End If
    Db.Execute Sql
    SetNothing Db
Exit Sub
ERH:
Trans_Log "Error " & Err.Description & " Closeallmarkd"
Err.Clear
End Sub
Public Sub CloseEMVCharges(CashierName As String)
    On Error GoTo ERH
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    ConnectADODB DebitDb, Db
      If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" And Input_Option(29) = "By Cashier" Then
        Sql = "Update [Debit] Set [Cashier Name] = 'Closed' Where [Close Out day] = 0"
        Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
      ElseIf Check_Option(60) = 0 Then
        Sql = "Update [Debit] Set [Cashier Name] = 'Closed' Where [Close Out day] = 0"
        Sql = Sql & " and [Cashier Name] = '" & CashierName & "' and [Register Name] = '" & MyRegisterName & "'"
      Else
        Sql = "Update [Debit] Set [Cashier Name] = 'Closed' Where [Close Out day] = 0"
        Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
      End If
    Db.Execute Sql
    DbClose Db, RS
    Exit Sub
ERH:
    Err.Clear
End Sub
Public Function ReturnPOSName() As String
   ReturnPOSName = "Micro$ale"
End Function
Public Function ReturnOEMName() As String
 ReturnOEMName = "Micro$ale POS "
 ReturnOEMName = ReturnOEMName & " " & GetPDAUser & " Users " & NetWork_Data + 1 & " Lic " & SiteLic.Users
End Function
Public Function GetSupportId() As String
Dim SiteId As String
Dim TID As MicrosaleLicense.GetPosLicense
Set TID = New MicrosaleLicense.GetPosLicense
SiteId = Trim$(TID.LicenseHasExpired(DBFilePath & GetLicFile))
GetSupportId = " Site Id " & SiteId
Set TID = Nothing
SaveSetting "ProductID", "License File", "Customer Site Id", SiteId

End Function
Public Function GetStarting(AccountName As String) As Currency
Dim sDb As ADODB.Connection
Dim SRS As ADODB.Recordset
If ConnectADODB(HouseAcct, sDb) = False Then Exit Function
Sql = "Select * from [Starting Balance] Where [Account Name] = '" & AccountName & "'"
If ConnectADORS(Sql, sDb, SRS) = False Then Exit Function
If Not SRS.EOF Then
    GetStarting = Val(SRS![Beginning Balance] & "")
Else
    GetStarting = 0
End If
DbClose sDb, SRS
End Function
Public Function ReturnComputerName() As TerminalId
Dim TName As String
Dim TData As String
Dim tn As GeneralLib.Utilities
Set tn = New GeneralLib.Utilities

TName = tn.GetKeyData("TerminalName")
Set tn = Nothing
Dim GL As GeneralLib.TerminalSetup
Set GL = New GeneralLib.TerminalSetup
TData = GL.RegisterTerminalVB6(TName)
Set GL = Nothing
ReturnComputerName = ParseTerminalInfoValues(TData)

End Function
Public Function ParseTerminalInfoValues(ByVal TerminalData As String) As TerminalId
    Dim parts() As String
    On Error GoTo ErrHandler
    parts = Split(TerminalData, "|")
    If UBound(parts) < 5 Then
        ParseTerminalInfoValues.TerminalNo = -1
        Exit Function
    End If
    With ParseTerminalInfoValues
    .TerminalNo = CLng(Val(parts(0)))
    .TerminalGuid = Trim$(parts(1))
    .TerminalName = Trim$(parts(2))
    .MachineName = Trim$(parts(3))
    .WindowsUser = Trim$(parts(4))
    .IpAddress = Trim$(parts(5))
    End With
    Exit Function
ErrHandler:
    ParseTerminalInfoValues.TerminalNo = -1
End Function
Public Function GetSvrName(DataIn As String) As String
On Error GoTo ERH
 
 'If Left(DataIn, 2) = "\\" Then
 ' Dim TP As SetPaths
 ' Set TP = New SetPaths
 ' GetSvrName = TP.ReturnCorrectTerminalName(DataIn)
 ' Set TP = Nothing
 ' Exit Function
 'End If

'If InStr(UCase(DataIn), UCase(SvrPath.TerminalName)) > 0 Then
'   GetSvrName = SvrPath.TerminalName
'   Exit Function
'End If
'If SvrPath.sTerminalAddress > "" And UCase(SvrPath.TerminalName) <> UCase(SvrPath.sTerminalAddress) Then
'  GetSvrName = SvrPath.TerminalIpAddress
'Else
  GetSvrName = SvrPath.TerminalName
'End If
Exit Function
ERH:
GetSvrName = SvrPath.TerminalName
End Function
Public Function Encrypt(DataIn As String, PWord As String) As String
If Trim$(DataIn) = "" Then
   Encrypt = ""
   Exit Function
End If
Encrypt = DataIn
If Left(DataIn, 4) = "XXXX" Then
  Exit Function
End If
Dim AES As MicrosaleLicense.GetPosLicense
Set AES = New MicrosaleLicense.GetPosLicense
Encrypt = AES.AES_Encrypt(DataIn, PWord)
Set AES = Nothing
End Function
Public Function Decrypt(DataIn As String, PWord As String) As String
If DataIn = "" Then
   Decrypt = ""
   Exit Function
End If
Decrypt = DataIn
If Left(DataIn, 4) = "XXXX" Then
  Exit Function
End If
Dim AES As MicrosaleLicense.GetPosLicense
Set AES = New MicrosaleLicense.GetPosLicense
Decrypt = AES.AES_Decrypt(DataIn, PWord)
Set AES = Nothing
End Function
Public Function TestConnection(SvrName As String, UN As String, PW As String) As String
Dim Catalog    As String
Catalog = "Initial Catalog=Master;"
TestConnection = "Provider=SQLOLEDB;Data Source=" & SvrName & ";" & Catalog & ";User Id= '" & UN & "';Password= '" & PW & "';"
End Function

Public Function ConnectADODB(ByVal DatabaseName As String, ByRef Db As ADODB.Connection) As Boolean
    On Error GoTo EH
    Dim DBName As String
    ConnectADODB = False
    DBName = GetDatabaseNameFromPath(DatabaseName)
    ' Optional pre-check you already have
    'If (GetSqlHouseAccount(DBName) Or DBName = "Frequent DinerSql") Then
        'If Val(Check_Option(124)) = 0 Then
        '    Dim MyTName As String
        '    If InStr(CPath.TerminalIpAddress, "\") > 0 Then
        '        MyTName = Left$(CPath.TerminalIpAddress, InStr(CPath.TerminalIpAddress, "\") - 1)
        '    Else
        '        MyTName = CPath.TerminalIpAddress
        '    End If''

        '    If TerminalIsPresent(MyTName) = False Then
        '        See_Msg "No Connection To Gift/House Account Server", 5, ""
        '        Exit Function
        '    End If
        'End If
    'End If

    ' Make sure any prior connection is not reused accidentally
    On Error Resume Next
    If Not Db Is Nothing Then
        If Db.State <> 0 Then Db.Close
    End If
    Set Db = Nothing
    On Error GoTo EH

    Set Db = New ADODB.Connection

    ' Server-side cursors are usually faster for SQL Server
    Db.CursorLocation = adUseServer
    Db.CommandTimeout = 30
    Db.ConnectionTimeout = 15
    Db.Mode = adModeReadWrite
    Db.IsolationLevel = adXactReadCommitted

    Db.ConnectionString = ReturnAdoConnection(DatabaseName)

    ' --- Simple retry (3 tries) ---
    Dim tries As Integer
    For tries = 1 To 3
        On Error GoTo TryFail
        Db.Open
        ConnectADODB = True
        Exit Function

TryFail:
        ' brief backoff
        On Error Resume Next
        Err.Clear
        Sleep 250
        On Error GoTo EH
    Next tries

    Exit Function

EH:
    Error_Log Err.Number, Err.Description, DBName, "None"
    See_Msg "Connection Failed...", 5, ""

    ' Clean up connection object
    On Error Resume Next
    If Not Db Is Nothing Then
        If Db.State <> 0 Then Db.Close
    End If
    Set Db = Nothing

    ' Do NOT End the program here
    ConnectADODB = False
End Function
Public Function AttachToStation(DataPathwDbaseName As String) As Boolean
 On Error GoTo ErrorHandler
  If TerminalFailed(DataPathwDbaseName) Then
    AttachToStation = False
    See_Msg "Register Failed " & GetSvrName(DataPathwDbaseName), 5, "Problem with NetWork or Station"
    Trans_Log "Register Failed " & DataPathwDbaseName
  Else
    AttachToStation = True
  End If
 Exit Function
ErrorHandler:
If DataPathwDbaseName = "" Then Exit Function
AttachToStation = False
See_Msg "Register Update Failed " & GetSvrName(DataPathwDbaseName), 5, "Problem with NetWork...."
'SETvariables
End Function
Public Function ReturnServerName(DatabaseName As String, DbaseName As String) As String
ReturnServerName = ""
If InStr(UCase(DatabaseName), ".MDB") = 0 Then
        
        Select Case DbaseName
           'Case Is = "House Accounts", "Frequent DinerSql", "BadChecksSql"
           '   ReturnServerName = CPath.TerminalIpAddress
           
           Case Else
              ReturnServerName = SvrPath.TerminalIpAddress
         End Select
End If
End Function
Public Function GetDatabaseNameFromPath(ByVal FullPath As String) As String
    On Error GoTo ERH

    Dim s As String
    Dim pos As Long

    s = Trim$(FullPath)

    ' Remove trailing "\" if present
    If Right$(s, 1) = "\" Then
        s = Left$(s, Len(s) - 1)
    End If

    ' Find last backslash
    pos = InStrRev(s, "\")

    If pos > 0 Then
        GetDatabaseNameFromPath = Mid$(s, pos + 1)
    Else
        GetDatabaseNameFromPath = s
    End If

    Exit Function

ERH:
    GetDatabaseNameFromPath = ""
End Function

Public Function ReturnAdoConnection(ByVal DatabaseName As String) As String

    Dim catalogPart As String
    Dim Server As String
    Dim User As String
    Dim PW As String
    Dim DbaseName As String
    If InStr(1, UCase$(DatabaseName), ".MDB", vbTextCompare) = 0 Then
        ' SQL Server
        DbaseName = GetDatabaseNameFromPath(DatabaseName)
        catalogPart = "Initial Catalog=" & DbaseName & ";"
        Select Case DbaseName
            'Case "House Accounts", "Frequent DinerSql"
            '    Server = CPath.TerminalSqlInstance
            '    User = CPath.ServerUser
            '    PW = CPath.ServerPW
            Case Else
                Server = SvrPath.TerminalSqlInstance
                User = SvrPath.ServerUser
                PW = SvrPath.ServerPW
        End Select

        ReturnAdoConnection = _
            "Provider=SQLOLEDB;" & _
            "Data Source=" & Server & ";" & _
            catalogPart & _
            "User ID=" & User & ";" & _
            "Password=" & PW & ";" & _
            "Persist Security Info=False;" & _
            "Connect Timeout=15;" & _
            "Application Name=MicroSale;"

    Else
        ' Access (Jet)
        ReturnAdoConnection = _
            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & DatabaseName & ";" & _
            "Persist Security Info=False;"
    End If

End Function


Public Function AdoConnection(ByVal DatabaseName As String, ByVal DbaseName As String) As ADODB.Connection
    On Error GoTo Fail

    Dim cn As ADODB.Connection
    Dim PathtoFile As String
    Dim dsn As String

    Set cn = New ADODB.Connection
    ' Tune these as you like:
    cn.ConnectionTimeout = 15
    cn.CommandTimeout = 30
    cn.CursorLocation = adUseClient   ' or adUseServer if you prefer server-side cursors

    ' Decide SQL Server vs. Access by file extension presence (your original logic)
    If InStr(1, UCase$(DatabaseName), ".MDB") = 0 And InStr(1, UCase$(DatabaseName), ".ACCDB") = 0 Then
        ' ---- SQL Server (SQLOLEDB) ----
        PathtoFile = "Initial Catalog=" & DbaseName

        Select Case DbaseName
            'Case "House Accounts", "Frequent DinerSql"
               ' dsn = "Provider=SQLOLEDB;" & _
                      "Data Source=" & CPath.TerminalIpAddress & ";" & _
                      PathtoFile & ";" & _
                      "Persist Security Info=False;" & _
                      "User ID=" & CPath.ServerUser & ";" & _
                      "Password=" & CPath.ServerPW & ";"
            Case Else
                dsn = "Provider=SQLOLEDB;" & _
                      "Data Source=" & SvrPath.TerminalIpAddress & ";" & _
                      PathtoFile & ";" & _
                      "Persist Security Info=False;" & _
                      "User ID=" & SvrPath.ServerUser & ";" & _
                      "Password=" & SvrPath.ServerPW & ";"
        End Select

        ' (Optional) If you’ve installed the modern provider, swap SQLOLEDB -> MSOLEDBSQL
        ' dsn = Replace$(dsn, "Provider=SQLOLEDB", "Provider=MSOLEDBSQL")

    Else
        ' ---- Access (JET / ACE) ----
        If Right$(UCase$(DatabaseName), 6) = ".ACCDB" Then
            dsn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=" & DatabaseName & ";" & _
                  "Persist Security Info=False"
        Else
            dsn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & DatabaseName & ";" & _
                  "Persist Security Info=False"
        End If
    End If

    cn.Open dsn
    Set AdoConnection = cn
    Exit Function

Fail:
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State <> adStateClosed Then cn.Close
    End If
    Set AdoConnection = Nothing
    ' Optional: surface error to your logger
    ' Trans_Log "ReturnAdoConnection failed: " & Err.Number & " - " & Err.Description
End Function

Public Function GetPayment(ByVal ChkNO As Long) As PType
    On Error GoTo ERH

    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim Sql As String

    ' ---- Defaults (always return something) ----
    GetPayment.TenderName = "Cash"
    GetPayment.TenderAmt = 0
    GetPayment.TenderTip = 0

    ' =========================
    ' 1) Try Debit (open day)
    ' =========================
    If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) Then

        Sql = "SELECT TOP 1 [Tip],[Amount],[Account Name] " & _
              "FROM [Debit] " & _
              "WHERE [Close Out day] = 0 AND [Check Number] = " & ChkNO & " " & _
              "ORDER BY [Time] DESC"

        If ConnectADORS(Sql, Db, RS) Then
            If Not RS.EOF Then
                GetPayment.TenderTip = NzCur(RS.fields("Tip").Value)
                GetPayment.TenderAmt = NzCur(RS.fields("Amount").Value)
                GetPayment.TenderName = NzStr(RS.fields("Account Name").Value)
                DbClose Db, RS
                Exit Function
            End If
        End If

        DbClose Db, RS
    End If

    ' =========================
    ' 2) Fall back to Payments
    ' =========================
    If ConnectADODB(SvrPath.TerminalPath & FinancialDb, Db) Then

        Sql = "SELECT TOP 1 [Payment Type],[Payment Amount],[Tip Amount] " & _
              "FROM [dayPayments] " & _
              "WHERE [Check Number] = " & ChkNO & " " & _
              "ORDER BY [Recorded Time] DESC"

        If ConnectADORS(Sql, Db, RS) Then
            If Not RS.EOF Then
                GetPayment.TenderName = NzStr(RS.fields("Payment Type").Value)
                GetPayment.TenderAmt = NzCur(RS.fields("Payment Amount").Value)
                GetPayment.TenderTip = NzCur(RS.fields("Tip Amount").Value)
            End If
        End If

        DbClose Db, RS
    End If

    Exit Function

ERH:
    See_Msg Err.Description, 5, "Error Loading Payments"
    On Error Resume Next
    DbClose Db, RS
End Function


Public Function GetPaymentAmt(ByVal CheckNo As Long) As Currency
    On Error GoTo EH

    Dim RS As ADODB.Recordset
    Dim GRS As GeneralLib.Utilities
    Dim Sql As String

    Set GRS = New GeneralLib.Utilities

    Sql = "SELECT TOP 1 [Amount] " & _
          "FROM [Debit] " & _
          "WHERE [Close Out Day] = 0 " & _
          "  AND [Check Number] = " & CheckNo

    Set RS = GRS.ReturnRecordset(Sql, _
                ReturnAdoConnection(SvrPath.TerminalPath & DebitDb), _
                "Debit")

    If Not RS.EOF Then
        If Not IsNull(RS!Amount) Then
            GetPaymentAmt = RS!Amount
        Else
            GetPaymentAmt = 0
        End If
    Else
        GetPaymentAmt = 0
    End If

CleanUp:
    DbClose Nothing, RS
    Set GRS = Nothing
    Exit Function

EH:
    Error_Log Err.Description, "Error Get Payment", "", ""
    GetPaymentAmt = 0
    Resume CleanUp
End Function
Public Function IsPaid(CheckNo As Long, Amount As Currency) As Boolean
IsPaid = False
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then Exit Function
  Sql = "Select top 1 1 from [Debit] where [Close out day] = 0 "
  Sql = Sql & " and [Check Number] = " & CheckNo
  Sql = Sql & " and [Amount] = " & Amount
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
  If Not Tb.EOF Then
   IsPaid = True
  End If
DbClose Db, Tb
End Function

Public Function IsCredit(ByVal CheckNo As Long) As Boolean

    IsCredit = False
    
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim Sql As String
    
    If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then Exit Function
    
    Sql = "SELECT TOP 1 1 " & _
          "FROM [Debit] " & _
          "WHERE [Close out day] = 0 " & _
          "  AND [Check Number] = " & CheckNo & " " & _
          "  AND [Tender] = 'EMV'"
    
    If ConnectADORS(Sql, Db, RS) = False Then GoTo CleanUp
    
    IsCredit = Not RS.EOF

CleanUp:
    DbClose Db, RS

End Function
Public Function isGiftCard(ByVal CheckNo As Long) As String
    On Error GoTo EH

    isGiftCard = ""

    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim Sql As String

    If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then Exit Function

    Sql = "SELECT TOP 1 [XMLString], [Account Number], [Time] " & _
          "FROM [Debit] " & _
          "WHERE [Close out day] = 0 " & _
          "  AND [Check Number] = " & CheckNo & " " & _
          "  AND [Tender] NOT IN ('EMV','QRPayment') " & _
          "ORDER BY [Time] DESC"

    If ConnectADORS(Sql, Db, RS) = False Then GoTo CleanUp

    If Not RS.EOF Then
        isGiftCard = Nz(RS.fields("XMLString").Value, "") & _
                    "<AccountNumber>" & Nz(RS.fields("Account Number").Value, "") & "</AccountNumber>" & _
                    "<Time>" & Nz(RS.fields("Time").Value, "") & "</Time>"
    End If

CleanUp:
    DbClose Db, RS
    Exit Function

EH:
    ' log if you want
    On Error Resume Next
    DbClose Db, RS
End Function
Public Function IsHouseAccount(ByVal CheckNo As Long) As Boolean
    On Error GoTo EH

    IsHouseAccount = False

    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim Sql As String
    Dim loc As String

    Sql = "SELECT TOP 1 1 " & _
          "FROM [HouseCharges] " & _
          "WHERE [Day of Charge] = 0 " & _
          "  AND [Voided] = 0 " & _
          "  AND [Check Number] = " & CheckNo

    ' Only allow void from location check was entered
    If Check_Option(165) = 1 Then
        loc = Replace(LocationReset, "'", "''") ' escape quotes
        Sql = Sql & " AND [Location Name] = '" & loc & "'"
    End If

    If ConnectADODB(HouseAcct, Db) = False Then Exit Function
    If ConnectADORS(Sql, Db, RS) = False Then GoTo CleanUp

    IsHouseAccount = Not RS.EOF

CleanUp:
    DbClose Db, RS
    Exit Function

EH:
    On Error Resume Next
    DbClose Db, RS
End Function
Public Function CheckOnHold(CNo As Long) As Boolean
'//option 71 is hold menu items
If Check_Option(71) = 0 Then CheckOnHold = False: Exit Function
On Error GoTo ErrorHandler
Dim Cdb As ADODB.Connection, CRS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Cdb
Sql = "Select top 1 [Check Number] from [Check Detail] where [On Hold] > 0"      'and [Check Number]= " & TryCheck
Sql = Sql & " and [Check Number] = " & CNo
Sql = Sql & " and [Close Out Day] = 0"
ConnectADORS Sql, Cdb, CRS
CheckOnHold = False
If Not CRS.EOF Then
 CheckOnHold = True
 See_Msg "Release Items On Hold", 0, "Release Item and Press Done"
End If
DbClose Cdb, CRS
Exit Function
ErrorHandler:
End Function
Public Function GetPrePaidPhoneNotCredit(MyCheck As Long, CheckTender As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetPrePaidPhoneNotCredit = 0
If ConnectADODB(SvrPath.TerminalPath & FinancialDb, Db) = False Then
  Exit Function
End If
Sql = "Select * from [daypayments] where [Payment Type] = '" & CheckTender & "' and [check Number] = " & Val(MyCheck)
Call ConnectADORS(Sql, Db, RS)
Do While Not RS.EOF
 GetPrePaidPhoneNotCredit = GetPrePaidPhoneNotCredit + RS.fields("Payment Amount").Value
 RS.MoveNext
Loop
Call DbClose(Db, RS)
GetPrePaidPhoneNotCredit = GetPrePaidPhoneNotCredit + PhoneOrderPaidAmount(MyCheck)
End Function
Public Function GetCardPayment(MyCheck As Long) As String
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
GetCardPayment = ""
   If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then Exit Function
    Sql = "Select * from [Debit] where [Tender] = 'EMV' And [Close out Day] = 0 and [Check Number] = " & MyCheck
   If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
   If Not Tb.EOF Then
     GetCardPayment = "Credit Card"  'Tb![Account Name] & ""
   End If
ERH:
DbClose Db, Tb
End Function
Public Function GetPrePayment(MyCheck As Long) As Currency
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
GetPrePayment = 0
   If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then Exit Function
    Sql = "Select * from [Debit] where [Tender] = 'EMV' And [Close out Day] = 0 and [Check Number] = " & MyCheck
   If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
   Do While Not Tb.EOF
     GetPrePayment = GetPrePayment + Tb![Amount]
    Tb.MoveNext
   Loop
ERH:
DbClose Db, Tb
End Function
Public Function PhoneOrderIsPaid(PaidCheckNo As Long) As Currency
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
PhoneOrderIsPaid = 0
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
 Sql = "Select * from [Open Checks] where [Check Number] = " & PaidCheckNo
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
 If Not Tb.EOF Then
    PhoneOrderIsPaid = Val(Tb![Payments] & "")
 End If
ERH:
DbClose Db, Tb
End Function
Public Function PhoneOrderPaidAmount(PaidCheckNo As Long) As Currency
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
PhoneOrderPaidAmount = 0
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
If PaidCheckNo < 1000 Then '// check numbers under 1000 have not been assigned to as driver
 Sql = "Select * from [PrePaid] where [PhoneOrderNumber] = " & PaidCheckNo
Else
 Sql = "Select * from [PrePaid] where [OrderNumber] = " & PaidCheckNo
End If
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
 If Not Tb.EOF Then
    PhoneOrderPaidAmount = Val(Tb.fields("PaymentAmount").Value & "")
    If AllOrders Is Nothing Then
       AllOrders.add 1
    End If
    If Tb.fields("PrePaidTender").Value & "" = "" Then
      AllOrders.Item(1).OrderPayment = "PrePaidTender"  '// added to make sure is not prepaid is found that it will still report a tender.
    Else
      AllOrders.Item(1).OrderPayment = Tb.fields("PrePaidTender").Value & ""
    End If
    AllOrders.Item(1).OrderPaymentAmt = Val(Tb.fields("PaymentAmount").Value & "")
 End If
ERH:
DbClose Db, Tb
End Function

Public Function PhoneOrderIsPaidHow(PaidCheckNo As Long) As String
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
PhoneOrderIsPaidHow = ""
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
If PaidCheckNo < 1000 Then '// check numbers under 1000 have not been assigned to as driver
 Sql = "Select * from [PrePaid] where [PhoneOrderNumber] = " & PaidCheckNo
Else
 Sql = "Select * from [PrePaid] where [OrderNumber] = " & PaidCheckNo
End If
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
 If Not Tb.EOF Then
    'PhoneOrderIsPaidHow = Tb.fields("PaymentAmount").Value & ""
    PhoneOrderIsPaidHow = Tb.fields("PrePaidTender").Value & ""
 End If
ERH:
DbClose Db, Tb
End Function
Public Function DeliveryPrePaidTender(PaidCheckNo As Long) As String
On Error GoTo ERH
DeliveryPrePaidTender = GetCardPayment(PaidCheckNo)
If DeliveryPrePaidTender > "" Then Exit Function
DeliveryPrePaidTender = ""
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
 Sql = "Select * from [Open Checks] where [Check Number] = " & PaidCheckNo
   If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
   If Not Tb.EOF Then
     If Val(Tb![Payments] & "") > 0 Then
       DeliveryPrePaidTender = PhoneOrderIsPaidHow(PaidCheckNo)
     ElseIf Val(Tb![Deposit] & "") > 0 Then
        DeliveryPrePaidTender = "Deposit"
     Else
       DeliveryPrePaidTender = ""
     End If
   End If
ERH:
DbClose Db, Tb
End Function
Public Function MovePrePayment(ByVal MyCheck As Long, ByVal NewCheckNo As Long, ByVal newServer As String) As Boolean
    On Error GoTo EH

    Dim Db As ADODB.Connection
    Dim rowsDebit As Long
    Dim rowsPay As Long
    Dim Sql As String

    MovePrePayment = False

    '--- 1) Update Debit DB (set-based)
    If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then GoTo CleanUp

    Sql = "UPDATE [Debit] " & _
          "SET [Check Number] = " & NewCheckNo & ", " & _
          "    [Server Name] = '" & SqlQuote(newServer) & "', " & _
          "    [Cashier Name] = '" & SqlQuote(newServer) & "' " & _
          "WHERE [Tender] = 'EMV' " & _
          "  AND [Close out Day] = 0 " & _
          "  AND [Check Number] = " & MyCheck & ";"

    Db.Execute Sql, rowsDebit, adExecuteNoRecords

    Db.Close
    Set Db = Nothing

    If rowsDebit <= 0 Then
        ' nothing to move
        GoTo CleanUp
    End If

    '--- 2) Update Payments in CheckStat DB
    If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then GoTo CleanUp

    Sql = "UPDATE [Payments] " & _
          "SET [Check Number] = " & NewCheckNo & " " & _
          "WHERE [Check Number] = " & MyCheck & " " & _
          "  AND [Close out day] = 0;"

    Db.Execute Sql, rowsPay, adExecuteNoRecords

    MovePrePayment = True
    Trans_Log "Moving Payment from " & MyCheck & " to " & NewCheckNo & _
              " (Debit rows=" & rowsDebit & ", Payments rows=" & rowsPay & ")"

CleanUp:
    On Error Resume Next
    If Not Db Is Nothing Then
        If Db.State <> adStateClosed Then Db.Close
        Set Db = Nothing
    End If
    Exit Function

EH:
    Trans_Log "MovePrePayment error " & Err.Number & ": " & Err.Description & _
           " (From " & MyCheck & " To " & NewCheckNo & ")"
    MovePrePayment = False
    Resume CleanUp
End Function

Private Function SqlQuote(ByVal s As String) As String
    SqlQuote = Replace$(s & "", "'", "''")
End Function

Public Function RecoverSuspendedPayment(MyCheck As Long, NewCheckNo As Long, newServer As String) As Boolean
On Error GoTo ERH
RecoverSuspendedPayment = False
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
   If ConnectADODB(SvrPath.TerminalPath & DebitDb, Db) = False Then Exit Function
    Sql = "Select * from [Debit] where [Tender] = 'EMV' And [Close out Day] = -1 and [Check Number] = " & MyCheck
   If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
    If Not Tb.EOF Then
     RecoverSuspendedPayment = True
     Do While Not Tb.EOF
       Tb.fields("Check Number").Value = NewCheckNo
       Tb.fields("Server Name").Value = newServer
       Tb.fields("Cashier Name").Value = newServer
       Tb.fields("Close Out Day").Value = 0
       Tb.Update
      Tb.MoveNext
     Loop
    End If
DbClose Db, Tb
ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
Sql = "Update [Pre Auth] Set [Check Number] = " & NewCheckNo & ", [Server Name] = '" & newServer & "',[Auth Amount] = 0 where [Auth Amount] = -1 and [Check Number] = " & MyCheck
Db.Execute Sql
Db.Close
ERH:
End Function
Public Function FindShiftPrefix(CategoryName As String, ShiftNo As Integer) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB SvrPath.LocalTerminalPath & "CRMENU.MDB", Db
    Sql = "Select * from [Shift Keys] Where [Category Name] = '" & CategoryName & "'"
    Sql = Sql & " and [Shift Level] = " & Val(ShiftNo)
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  FindShiftPrefix = RS![Key Name] & ""
Else
  FindShiftPrefix = ""
End If
DbClose Db, RS
End Function
Public Function GetModPrefix(PFX As String, ItemIn As String) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim RmvPfx() As String
ConnectADODB SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
    Sql = "Select [Prefix] from [Price Levels] Where [Modifier Title] = '" & Trim$(ItemIn) & "'"
    Sql = Sql & " and [PreFix] = '" & Trim$(PFX) & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  GetModPrefix = Trim$(RS![PreFix]) & ItemIn
Else
  GetModPrefix = ItemIn
End If
DbClose Db, RS
End Function
Public Function GetAltPricing(PFX As String, ItemIn As String, ShiftLevel As Integer) As Currency
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim AltName As String
GetAltPricing = -1
ConnectADODB SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
    Sql = "Select * from [Modifiers] Where [Modifier Category] = '" & ItemIn & "'"
    Sql = Sql & " and [Modifier Name] = '" & PFX & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  AltName = RS.fields("Alt Modifier").Value & ""
End If
If AltName <> "None" And AltName > "" Then
  SetNothing RS
    Sql = "Select * from [Price Levels] Where [Modifier Title] = '" & ItemIn & "'"
    Sql = Sql & " and [Modifier Category] = '" & AltName & "'"
    Sql = Sql & " and [Level] = " & ShiftLevel - 1
  ConnectADORS Sql, Db, RS
  If Not RS.EOF Then
    GetAltPricing = Format(SetDualPrice(Val(RS.fields("Price").Value & "")), "0.00")
  End If
End If
DbClose Db, RS
End Function
Public Function GetModPrefixShift(PFX As String, ItemIn As String) As Long
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim RmvPfx() As String
ConnectADODB SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
    Sql = "Select * from [Modifiers] Where [Modifier Category] = '" & Trim$("!" & PFX) & "'"
    Sql = Sql & " and [Modifier Name] = '" & ItemIn & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  GetModPrefixShift = Val(RS.fields("Price Level").Value & "")
Else
  GetModPrefixShift = 0
End If
DbClose Db, RS
End Function
Public Function GetKeyPrefix(ItemIn As String) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim RmvPfx() As String
ConnectADODB SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
    Sql = "Select [Prefix] from [Price Levels] Where [Modifier Title] = '" & Trim$(ItemIn) & "'"
    Sql = Sql & " and [Level] = " & (MatrixPriceLevel - 1)
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  GetKeyPrefix = RS![PreFix] & ItemIn
Else
  GetKeyPrefix = ItemIn
End If
DbClose Db, RS
End Function
Public Function GetCashTenders(ByVal InDate As String, ByVal EDate As String) As BalanceSheetValues
On Error GoTo ErrorHandler
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
GetCashTenders.DayCashTenders = 0: GetCashTenders.DayNoCashTips = 0
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db

 Sql = "Select Distinct sum([Amount]) as TCash  from [CheckPayments] "
 Sql = Sql & " Where [Media] = 'Cash'"
 Sql = Sql & " and TRY_CAST([Date] AS datetime) Between '" & CDate(InDate & " " & GetStartTime) & "'"
 Sql = Sql & " and '" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
   GetCashTenders.DayCashTenders = Val(RS.fields("TCash").Value & "")
End If
SetNothing RS
 Sql = "Select sum([Tip]) as TTip from [CheckPayments] "
 Sql = Sql & " Where [Media] <> 'Cash'"
 Sql = Sql & " and TRY_CAST([Date] AS datetime) Between '" & CDate(InDate & " " & GetStartTime) & "'"
 Sql = Sql & " and '" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "'"
 
ConnectADORS Sql, Db, RS
'// used this return date string when we supported both access and sql
If Not RS.EOF Then
   GetCashTenders.DayNoCashTips = Val(RS.fields("TTip").Value & "")
End If
DbClose Db, RS

Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "DSR Balance Sheet ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function UsingCountsAsCustomer(SRC As ADODB.Connection, Optional ByVal forceRefresh As Boolean = False) As Integer
    Static initialized As Boolean
    Static cachedValue As Integer
    '// added this function due to old system with incomplatable databases when you merge
    
    If Not initialized Or forceRefresh Then
        'Dim SRC As adodb.Connection
        Dim RS As ADODB.Recordset
        Dim Sql As String

        On Error GoTo CleanFail

        Sql = "SELECT [Menu Item Name] " & _
              "FROM [MSaleData].[dbo].[Menu Items] " & _
              "WHERE [Item Info] = 'CountasCustomer'"

        Set SRC = New ADODB.Connection
        SRC.Open ReturnAdoConnection("MSaleData")

        Set RS = New ADODB.Recordset
        RS.CursorLocation = adUseClient
        RS.Open Sql, SRC, adOpenForwardOnly, adLockReadOnly

        cachedValue = IIf(Not RS.EOF, 1, 0)
        initialized = True

cleanexit:
        On Error Resume Next
        If Not RS Is Nothing Then
            If RS.State <> adStateClosed Then RS.Close
            Set RS = Nothing
        End If
        If Not SRC Is Nothing Then
            If SRC.State <> adStateClosed Then SRC.Close
            Set SRC = Nothing
        End If

        UsingCountsAsCustomer = cachedValue
        Exit Function

CleanFail:
        ' If the first call fails, we still "initialize" to a safe default (0)
        cachedValue = 0
        initialized = True
        Resume cleanexit
    End If

    ' Fast path: just return the cached value
    UsingCountsAsCustomer = cachedValue
End Function
Public Function ReturnMealCount(StartDate As String, EndDate As String) As MealCounts
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim MealCount  As Integer
Dim MealSales  As Currency
If UsingCountsAsCustomer(Nothing) = 0 Then
  ReturnMealCount.MealCount = 0
  ReturnMealCount.MealSale = 0
  Exit Function
  '// if mealcountsascutomer is not in use return a 0
End If
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
 Sql = "Select Distinct [Menu Item], Count([Quantity]) as Customer ,Sum([Item Price]) as MealPrice from [Detail]"
 Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & CDate(StartDate & " " & GetStartTime) & "'"
 Sql = Sql & " and '" & CDate(SetEndingDate(EndDate) & " " & SetEndingTime) & "'"
 Sql = Sql & " and [menu item] in (select [Menu Item Name] from [MSaleData].[dbo].[Menu Items] Where [Item Info] = 'CountasCustomer') "
 Sql = Sql & " Group by [Menu Item]"
ConnectADORS Sql, Db, RS
 MealCount = 0: MealSales = 0
 Do While Not RS.EOF
   MealCount = MealCount + RS.fields("Customer").Value
   MealSales = MealSales + RS.fields("MealPrice").Value
  RS.MoveNext
 Loop
 DbClose Db, RS
 Trans_Log "Meal Count " & MealCount & " Sales " & MealSales
 ReturnMealCount.MealCount = MealCount
 ReturnMealCount.MealSale = MealSales
End Function
Public Function CheckFingerPrint() As String
 CheckFingerPrint = ""
 FingerPrint.Show 1
 If FPIDFound = "" Then
   See_Msg "Employee Finger Print Not Found", 5, ""
 ElseIf FPIDFound = "Exit" Then
    FPIDFound = ""
    Exit Function
 End If
 CheckFingerPrint = FPIDFound
End Function
Public Function Nz(ByVal Value As Variant, Optional ByVal Default As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = Default
    Else
        Nz = Value
    End If
End Function
Public Function ReturnDailyMealCount() As Long
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim MealCount  As Integer
If UsingCountsAsCustomer(Nothing) = 0 Then
  ReturnDailyMealCount = 0
  Exit Function
  '// if mealcountsascutomer is not in use return a 0
End If
MealCount = 0
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
 Sql = "Select Distinct [Menu Item], Count([Quantity]) as Customer from [Check Detail]"
 Sql = Sql & " Where [Voided] = 0 and [close out day] = 0 and [menu item] in (select [Menu Item Name] from [MSaleData].[dbo].[Menu Items] Where [Item Info] = 'CountasCustomer') "
 Sql = Sql & " Group by [Menu Item]"
ConnectADORS Sql, Db, RS
 Do While Not RS.EOF
   MealCount = MealCount + RS.fields("Customer").Value
  RS.MoveNext
 Loop
 DbClose Db, RS
 Trans_Log "Meal Count " & MealCount
 ReturnDailyMealCount = MealCount
Exit Function
ErrorHandler:
Trans_Log "Error with countsascustomer " & Err.Description
Err.Clear
End Function
Public Function SendNotification(GuidId As String) As String
On Error GoTo ERH
Dim R  As String
Dim SR As BCFiscalPost.BCSales
Set SR = New BCFiscalPost.BCSales
   R = SR.SendOnLineOrderReady(GuidId)
Set SR = Nothing
If InStr(R, "success") > 0 Then
   See_Msg "Customer Notified", 5, "Their Order Is Ready"
End If
Trans_Log "Message Returned from Send Order Ready Notification " & R
Exit Function
ERH:
Trans_Log "Error On Notify Customer " & Err.Description
'// added for champs chicken 08/20/2020
End Function
Public Function IsLastBag() As Currency
IsLastBag = 0
Dim FB As Integer
For FB = 0 To Number_In_Party - 1
 IsLastBag = IsLastBag + Seat_Sales_Data(FB, 5)
Next FB
End Function
Public Function ReleaseDrink(ByVal ItemPLU As String) As Boolean
    On Error GoTo ErrorHandler

    Dim FP As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim Sql As String
    Dim PourID As Long

    ReleaseDrink = False
    ItemPLU = Trim$(ItemPLU)

    If Check_Option(127) = 0 Then
        ReleaseDrink = True
        Exit Function
    End If

    If Len(ItemPLU) = 0 Then Exit Function

    If ConnectADODB(FinancialDb, FP) = False Then Exit Function

    Sql = "SELECT TOP 1 * " & _
          "FROM [Pours] " & _
          "WHERE [PLU No] = '" & Replace(ItemPLU, "'", "''") & "' " & _
          "AND [Open] = 1 " & _
          "AND [Closed] = 0 " & _
          "AND [Check No] > 0 " & _
          "AND [Register Name] = '" & Replace(MyRegisterName, "'", "''") & "' " & _
          "ORDER BY [PLU No]"

    If ConnectADORS(Sql, FP, RS) = False Then GoTo CleanUp

    If RS.EOF Then
        ReleaseDrink = False
    Else
        PourID = Val(RS.fields("PLU NO").Value & "")
        RS.Close
        Set RS = Nothing

        Sql = "UPDATE [Pours] " & _
              "SET [Open] = 0, [Closed] = 1 " & _
              "WHERE [PLU NO] = " & CStr(PourID)

        FP.Execute Sql
        ReleaseDrink = True
    End If

CleanUp:
    DbClose FP, RS
    Exit Function

ErrorHandler:
    ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Berg Free Pour")
    Select Case ErrorRespose
        Case "Abort"
            Resume CleanUp
        Case "Resume"
            Resume
        Case "Resume Next"
            Resume Next
        Case Else
            Resume CleanUp
    End Select
End Function
Function RoundPriceTo5or9(Price As Double) As Double
    If Check_Option(106) = 0 Then
       RoundPriceTo5or9 = Price
       Exit Function '// makes rounding optional for dual pricing
    End If
    Dim intPart As Integer
    Dim decimalPart As Double
    Dim cents As Integer
    
    intPart = Int(Price)
    decimalPart = Price - intPart
    cents = CInt(decimalPart * 100)

    Select Case cents Mod 10
        Case 0 To 2
            cents = cents - (cents Mod 10) ' round down to nearest 0
        Case 3 To 6
            cents = cents - (cents Mod 10) + 5 ' nearest 5
        Case 7 To 9
            cents = cents - (cents Mod 10) + 9 ' nearest 9
    End Select

    RoundPriceTo5or9 = intPart + (cents / 100)
End Function
Public Function OpenRSFast(ByVal cn As ADODB.Connection, ByVal sqlText As String) As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset

    RS.CursorLocation = adUseClient
    RS.Open sqlText, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Set OpenRSFast = RS
End Function

Public Function ConnectADORS(SqlString As String, Db As ADODB.Connection, RS As ADODB.Recordset) As Boolean
On Error GoTo ErrorHandler
Set RS = New ADODB.Recordset
RS.CursorLocation = adUseClient
RS.CursorType = adOpenStatic
RS.LockType = adLockOptimistic
If RS.State = 1 Then
 SetNothing RS
End If
RS.Open SqlString, Db
ConnectADORS = True
Exit Function
ErrorHandler:
ConnectADORS = False
If Err.Description = "Connection failure" Then
    Error_Log Err.Number, Err.Description & " " & SqlString, "connted rs", ""
    Resume Next
End If
If Err.Number = -2147467259 Then
   See_Msg "Network Connection Error", 5, ""
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "ADODB Record Set " & SqlString)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ReturnAuditGroup() As String
  Dim Db As ADODB.Connection
  Dim RS As ADODB.Recordset
  ReturnAuditGroup = ""
  ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
   Sql = "Select * from [Check Numbers] Where [Audit Status] = 0 and [Register Group] = '" & SvrPath.TerminalGroupName & "'"
   ConnectADORS Sql, Db, RS
   If RecordsCount(RS) > 1 Then
     DbClose Db, RS
     Load_What = "GroupAudits"
     Find_What = "Select Audit Group"
     Select_What.Show 1
     ReturnAuditGroup = GetSelected
   Else
     If Not RS.EOF Then
       ReturnAuditGroup = RS.fields("Work Area").Value & ""
     End If
    DbClose Db, RS
   End If
   GetSelected = ""
End Function
Public Function GetPrintRouting(PrtName As String, TBLName As String) As String
On Error GoTo ERH
GetPrintRouting = PrtName
If Check_Option(0) = 1 Then Exit Function
If TBLName = "" Then Exit Function
If Not TBPrinter.Item(TBLName & PrtName) Is Nothing Then
  GetPrintRouting = TBPrinter.Item(TBLName & PrtName).PrinterName
End If
Exit Function
ERH:
GetPrintRouting = PrtName
Err.Clear
End Function
Public Function FormatTxt(DataIn As String) As String
On Error GoTo ERH
 Dim astrWords()    As String
 Dim C              As Long
 Dim CountWords     As Long
 Dim Lines          As String
 astrWords = Split(DataIn)
 CountWords = UBound(astrWords) - LBound(astrWords) + 1
 Lines = "": FormatTxt = ""
 For C = 0 To CountWords - 1
  If Len(Trim$(astrWords(C))) + Len(Lines) < 40 Then
   Lines = Lines & astrWords(C) & " "
  Else
    FormatTxt = FormatTxt & Lines & vbCrLf
    Lines = ""
    C = C - 1
    If C < 0 Then Exit For
  End If
 Next C
 FormatTxt = FormatTxt & Lines & vbCrLf
 Exit Function
ERH:
End Function
Public Function GetDeliveryCount(SvrName As String, AuditStartTime As String) As Long
On Error GoTo ERH
Dim Db As Object
Dim OP As ADODB.Recordset
Dim Reg As Integer
ConnectADODB CPath.TerminalPath & CallInOrderDb, Db
 Sql = "Select distinct [Order Destination], count([Order Destination]) as TotalDel from [Open Checks] where [Server Name] = '" & SvrName & "' and [Order Destination] = 'Delivery'"
 Sql = Sql & "Group by [Order Destination]"
ConnectADORS Sql, Db, OP
GetDeliveryCount = 0
  If Not OP.EOF Then
   GetDeliveryCount = Val(OP![TotalDel] & "")
  End If
  DbClose Db, OP
  ConnectADODB CPath.TerminalPath & CheckStatDb, Db
  Sql = "Select * from [Closed Checks] where [Server Name] = '" & SvrName & "'"
  Sql = Sql & " and TRY_CAST([Order Date] AS datetime) + TRY_CAST([Close time] AS datetime) >= '" & AuditStartTime & "'"
  Sql = Sql & " and [Check Total] > 0"
  Sql = Sql & " and [Order Destination] = 'Delivery' and [close out day] =0"
  ConnectADORS Sql, Db, OP
  Do While Not OP.EOF
    GetDeliveryCount = GetDeliveryCount + 1
    OP.MoveNext
  Loop
DbClose Db, OP
Exit Function
ERH:
Err.Clear
GetDeliveryCount = 0
End Function
Public Function GetDriverSales(DriversName As String, AuditStart As String) As Currency
On Error GoTo ERH
Dim Db              As Object
Dim OP              As ADODB.Recordset
GetDriverSales = 0
ConnectADODB CPath.TerminalPath & CheckStatDb, Db
  Sql = "Select * from [Closed Checks] where [Server Name] = '" & DriversName & "'"
  Sql = Sql & " and TRY_CAST([Order Date] AS datetime) + TRY_CAST([Close Time] AS datetime) >= '" & AuditStart & "'"
  Sql = Sql & " and [Order Destination] = 'Delivery'  and [Close Out Day] = 0"
ConnectADORS Sql, Db, OP
    Do While Not OP.EOF
       GetDriverSales = GetDriverSales + Format(Val(OP![Check Total] & ""), "0.00")
       OP.MoveNext
    Loop
DbClose Db, OP
Exit Function
ERH:
See_Msg Err.Description, 5, "error in delivery sales"
Error_Log Err.Number, Err.Description, "Delivery Sales", ""
Err.Clear
End Function
Public Function CurrentDeptSales(DeptReports() As DeptSale) As Integer
On Error GoTo ErrorHandler
CurrentDeptSales = -1
If TotalNetSales = 0 Then Exit Function
Dim Cdb As ADODB.Connection, Ctb As ADODB.Recordset
Dim catdb As ADODB.Connection
Dim CatTb As ADODB.Recordset
Dim pos As Double
Dim RecordNo As Integer, TSperc As Double
Dim DeptRecords As Integer
Dim CatList(0 To 3) As New Collection
Dim CatReport() As DeptSale
Dim DiscReport() As DeptSale
Dim LL As Integer
Dim SoldQty As Long
      If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Cdb) Then
            Sql = "Select Distinct [Cateory Name] from [Check Detail] "
            Sql = Sql & " Where [Cateory Name] > ''"
            Sql = Sql & " order by [Cateory Name] asc "
        ConnectADORS Sql, Cdb, Ctb
           Do While Not Ctb.EOF
             On Error Resume Next
             CatList(0).add CStr(Ctb![Cateory Name] & ""), CStr(Ctb![Cateory Name] & "")
            Ctb.MoveNext
           Loop
           SetNothing Ctb
            Sql = "Select Distinct [Department Report] from [Check Detail] "
            Sql = Sql & " Where [Department Report] > ''"
            Sql = Sql & " order by [Department Report] asc "
        ConnectADORS Sql, Cdb, Ctb
           Do While Not Ctb.EOF
                On Error Resume Next
                      CatList(1).add CStr(Ctb![Department Report] & ""), CStr(Ctb![Department Report] & "")
            Ctb.MoveNext
           Loop
        DbClose Cdb, Ctb
      End If
On Error GoTo ErrorHandler
    ReDim CatReport(CatList(0).count) As DeptSale
    ReDim DeptReports(CatList(1).count) As DeptSale
    For LL = 0 To CatList(1).count - 1
       DeptReports(LL).Department = CatList(1).Item(LL + 1)
    Next LL
    RecordNo = CatList(0).count - 1
    DeptRecords = CatList(1).count - 1
Set CatList(0) = Nothing
Set CatList(1) = Nothing
'depart ment totals
        Sql = "Select Distinct [Department Report] , sum([Quantity]) as Qtysold, sum([price]) as TotalSold"
        Sql = Sql & " , sum([Cost Amount]) as TotalCost"
        Sql = Sql & " , sum([Inclusive Amount]) as TotalInclusive"
        Sql = Sql & " from [Check detail]"
        Sql = Sql & " where [Voided] = 0 "
        Sql = Sql & " and [Department Report] > '' "
        Sql = Sql & " and [close out day] = 0"
        Sql = Sql & " group by [Department Report]"
        Sql = Sql & " Order by [Department Report] asc"
    If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, catdb) Then
    ConnectADORS Sql, catdb, CatTb
        Dim GetCat As Integer
        Do While Not CatTb.EOF
           For GetCat = 0 To DeptRecords
           If Trim(UCase(DeptReports(GetCat).Department)) = Trim(UCase(CatTb![Department Report] & "")) Then
                SoldQty = Val(CatTb!QtySold & "")
                If SoldQty = 0 Then SoldQty = 1
                DeptReports(GetCat).Quantity = DeptReports(GetCat).Quantity + SoldQty
                DeptReports(GetCat).DollarsSold = DeptReports(GetCat).DollarsSold + (Val(Check_Null(CatTb![Totalsold]))) - (Val(Check_Null(CatTb!TotalInclusive)))
                DeptReports(GetCat).DollarsCost = DeptReports(GetCat).DollarsCost + (Val(Check_Null(CatTb![totalcost])))
            Exit For
           End If
           Next GetCat
          CatTb.MoveNext
        Loop
    DbClose catdb, CatTb
  End If
 
'***** added  ******
        Sql = "Select Distinct [Department Report], sum([Quantity]) as Qtysold, sum([price]) as TotalSold"
        Sql = Sql & " , sum([Cost Amount]) as TotalCost"
        Sql = Sql & " , sum([Inclusive Amount]) as TotalInclusive" ''
        Sql = Sql & " from [Check detail]"
        Sql = Sql & " where [Voided] = 0 and [Price] < 0"
        Sql = Sql & " and [Department Report] > '' "
        Sql = Sql & " and [Department Report] <> 'Non Sales Discount'"
        Sql = Sql & " and [close out day] = 0"
        Sql = Sql & " group by [Department Report]"
        Sql = Sql & " Order by [Department Report] asc"
    If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, catdb) Then
    ConnectADORS Sql, catdb, CatTb
    Do While Not CatTb.EOF
       For GetCat = 0 To DeptRecords
       If Trim(UCase(DeptReports(GetCat).Department)) = Trim(UCase(CatTb![Department Report])) Then
        SoldQty = (Val(Check_Null(CatTb!QtySold)))
        If SoldQty = 0 Then SoldQty = 1
        DeptReports(GetCat).QuantityDisc = DeptReports(GetCat).QuantityDisc + SoldQty
        DeptReports(GetCat).DollarsDisc = DeptReports(GetCat).DollarsDisc + (Val(Check_Null(CatTb![Totalsold]))) - (Val(Check_Null(CatTb!TotalInclusive)))
        Exit For
       End If
       Next GetCat
      CatTb.MoveNext
    Loop
    DbClose catdb, CatTb
  End If
 
CurrentDeptSales = UBound(DeptReports, 1)
Set CatList(0) = Nothing
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function GetOrderDestNo(OrderIn As String) As Integer
  Select Case OrderIn
    Case Is = "Go Order"
      GetOrderDestNo = SalesDest.GoSale
    Case Is = "Drive Thru"
      GetOrderDestNo = SalesDest.DTSale
    Case Is = "Pick Up"
      GetOrderDestNo = SalesDest.PKUPSale
    Case Is = "Delivery"
      GetOrderDestNo = SalesDest.DelSale
    Case Is = "Catering"
      GetOrderDestNo = SalesDest.catering
    Case "Dine In", "Here Order"
     GetOrderDestNo = SalesDest.HereSale
    Case Else
     GetOrderDestNo = -1
  End Select
End Function
Public Function GetOrderType(OrderSel As SalesDest) As String
   GetOrderType = OrderFor(OrderSel).Destination
End Function
Public Function MyOrderStatus(StatType As ChkStat)
    MyOrderStatus = StatType
End Function
Public Function CheckStatus(StatType As ChkStat) As String
     Select Case StatType
        Case Is = ChkStat.AddToCheck
          CheckStatus = "Add To Check"
        Case Is = ChkStat.CallIncheck
          CheckStatus = "Call In Check"
        Case Is = ChkStat.CancelCheck
          CheckStatus = "Cancel Check"
        Case Is = ChkStat.ChangeCheck
          CheckStatus = "Change Check"
        Case Is = ChkStat.CloseCheck
          CheckStatus = "Close Check"
        Case Is = ChkStat.DriveThruCheck
          CheckStatus = "Drive Thru Check"
        Case Is = ChkStat.EditCheck
          CheckStatus = "Edit Check"
        Case Is = ChkStat.EditHold
          CheckStatus = "Edit Hold"
        Case Is = ChkStat.ExpressCheck
          CheckStatus = "Express Check"
        Case Is = ChkStat.GoCheck
          CheckStatus = "Go Order"
        Case Is = ChkStat.HereCheck
          CheckStatus = "Here Check"
        Case Is = ChkStat.HoldCharge
          CheckStatus = "Hold Charge"
        Case Is = ChkStat.LoadCallInCheck
          CheckStatus = "Load CallIn Check"
        Case Is = ChkStat.MergeCheck
          CheckStatus = "Merge Check"
        Case Is = ChkStat.NewCheck
          CheckStatus = "New Check"
        Case Is = ChkStat.ParkCheck
          CheckStatus = "Park Check"
        Case Is = ChkStat.PAYATTABLE
          CheckStatus = "Pay at table"
        Case Is = ChkStat.PreAuthCheck
          CheckStatus = "Pre Auth Check"
        Case Is = ChkStat.PresentCheck
          CheckStatus = "Present Check"
        Case Is = ChkStat.PrintCheck
          CheckStatus = "Print Check"
        Case Is = ChkStat.ReleaseCallIn
          CheckStatus = "Release CallIn"
        Case Is = ChkStat.SplitCheck
          CheckStatus = "Split Check"
        Case Is = ChkStat.SplitPayCheck
          CheckStatus = "Split PayCheck"
        Case Is = ChkStat.SplitPresentcheck
          CheckStatus = "Split PresentCheck"
        Case Is = ChkStat.TranferPhone
          CheckStatus = "Transfer Phone"
        Case Is = ChkStat.UnDoSplit
          CheckStatus = "UnDo Split"
     End Select
        
End Function
Public Function GetRemote(Rmt_Printer As String, DeptName As String) As String
If Rmt_Printer = "None" Then
 GetRemote = Rmt_Printer
 Exit Function
End If
If Check_Option(189) = 1 And Check_Option(78) = 1 Then
 If InStr(Rmt_Printer, "Monitor") > 0 Then
  GetRemote = GetMonitor(Rmt_Printer)
  Exit Function
 End If
End If
If Check_Option(189) = 1 Then
  If Rmt_Printer = "" And DeptName > "" Then
   GetRemote = GetDeptPrinter(DeptName)
  Else
   GetRemote = Rmt_Printer
  End If
  Exit Function
End If
If Check_Option(78) = 1 And InStr(Rmt_Printer, "Monitor") > 0 Then
  GetRemote = GetMonitor(Rmt_Printer)
Else
  GetRemote = Rmt_Printer
End If
End Function
Public Function Set86Item(ItemName As String) As Integer
On Error GoTo ERH
Set86Item = -1
If ItemName = "" Then Exit Function
Dim DPName  As String
Dim QP      As Integer
Dim OnHand  As Integer
Dim InvOH   As Integer
DPName = "None"
Find_What = "Select Function"
Load_What = "Item Availability"
Select_What.Show 1
If GetSelected = "" Then Exit Function
If GetSelected = "86 Menu Item" Then
 QP = 0: OnHand = 0: InvOH = 0
Else
  OnHand = Val(ShowNumberPad("Enter Qty", False, 0, 0))
  QP = 1: InvOH = OnHand
End If
If Not AllMenuItems(ItemName) Is Nothing Then
  DPName = AllMenuItems(ItemName).DepartmentName
End If
Dim Menu     As ADODB.Connection
Dim MRecords As ADODB.Recordset
Dim SetCD As VER11SystemMaintenance.MaintenanceMenus
Set SetCD = New VER11SystemMaintenance.MaintenanceMenus
    Call SetCD.SetCountDown(ItemName, DPName, QP, OnHand, InvOH)
Set SetCD = Nothing
ConnectADODB SvrPath.TerminalPath & "MSaleData", Menu
 Sql = "Update [Menu Items] set [Quick Button] = " & 0 & " ,[Inventory Amount] = " & OnHand & " Where [menu item name] = '" & ItemName & "'"
 Menu.Execute Sql
'ConnectADORS Sql, Menu, MRecords
'If Not MRecords.EOF Then
'    MRecords![Quick Button] = QP
'    MRecords![Inventory Amount] = OnHand
'    MRecords.Update
    Set86Item = OnHand
    See_Msg "Saved....", 5, ItemName
'Else
'   See_Msg "No Update", 5, "Item Not Found"
'End If
DbClose Menu, Nothing 'MRecords
SendUpDate "CRMENU.MDB"
L_MenuMem ""
Exit Function
ERH:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "86 Item ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ReturnKVSTerminalId() As Long
ReturnKVSTerminalId = Mid(KDSName, InStr(KDSName, "-") + 1, Len(KDSName))
End Function
Public Function RequestsValue(RequestName As String, ItemName As String) As Currency
On Error GoTo ermsg
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
RequestsValue = 0
ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Db
  Sql = "Select * from [Special Requests] Where [Remote 1] = '" & ItemName & "' AND [Request Name] = '" & RequestName & "'"
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
  RequestsValue = 0
Else
  RequestsValue = Format(Val(Tb![Amount] & ""), " 0.00")
End If
DbClose Db, Tb
Exit Function
ermsg:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Request ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function InventoryDeduction(RequestName As String) As Integer
On Error GoTo ermsg
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
InventoryDeduction = 1
ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Db
  Sql = "Select * from [Void Reasons] Where [Void Name] = '" & RequestName & "'"
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
  InventoryDeduction = 1
Else
  InventoryDeduction = Val(Tb![Inventory] & "") + 1
End If  ' 0 or 1 add 1 to set value
DbClose Db, Tb
If InventoryDeduction = 0 Then InventoryDeduction = 1
Exit Function
ermsg:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Voids ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function Delete_Payment_Record(VoidCheckNo As Long, S_Name As String, VoidAuth As Boolean) As Boolean
On Error GoTo ErrorHandler
Delete_Payment_Record = False
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim ChargeType  As String
Dim DeviceIp    As String
Dim GetAccount  As Boolean
Dim FeeAmount   As Currency
Dim DepositAmt  As Currency
Dim AllData$
If EMVINTERFACE = "" Then
  Delete_Payment_Record = True
End If
    Dim xml As MPxmlParse
    Dim DCAC As ver11DebitCards.Accounting
    Set DCAC = New ver11DebitCards.Accounting
    Dim AccountFound        As Boolean
    Dim Accounts            As Integer
    Dim AcctData()          As ver11DebitCards.AccountInfo
    ReDim AcctData(0)
    Dim ScoInfo             As ver11DebitCards.ReturnedResponse
    AcctData(0).CheckNo = VoidCheckNo
    AcctData(0).AccountName = "Void"   '// send over to only void current days work
    GetAccount = DCAC.GetAccountChgData(AcctData(), SvrPath.TerminalPath)
If GetAccount Then
     For Accounts = 0 To UBound(AcctData, 1)
       Select Case AcctData(Accounts).TenderName
          Case Is = "InterCard"
            Dim IC As ver11DebitCards.InterCard
            Set IC = New ver11DebitCards.InterCard
             If AcctData(Accounts).AccountName = "Sell Card" Then
               Call IC.InterCardTransaction(AcctData(Accounts).AuthCode, TransType.DebitAccount, Val(AcctData(Accounts).ckTotal), -1, -1, -1)
             Else
               Call IC.InterCardTransaction(AcctData(Accounts).AuthCode, TransType.CreditAccount, Val(AcctData(Accounts).ckTotal), -1, -1, -1)
             End If
            Set IC = Nothing
          Case Is = "EMV"
            If VoidAuth Then GoTo VoidPreAuth
            If See_Msg("Void charge ", 4, "Continue ?") <> 6 Then
              Exit Function
            End If
            If See_Msg("Void Card the charge will be lost ", 4, "Are You Sure?") <> 6 Then
              Exit Function
            End If
            
            If Can_Continue("Void Check") = False Then
             Delete_Payment_Record = False
             Exit Function
            End If
VoidPreAuth:
            DeviceIp = ""
            If IsPAXInterface(EMVINTERFACE) Then
              Set xml = New MPxmlParse
              xml.xml = AcctData(Accounts).XMLDataString
              If xml.token("SeqNum") > "" And xml.token("IPAddress") > "" Then
                DeviceIp = xml.token("IPAddress")
              ElseIf xml.token("PaxIp") > "" Then
                DeviceIp = xml.token("PaxIp")
              Else
               If AcctData(Accounts).RegisterName <> MyRegisterName Then
                 DeviceIp = GetPaxIpAddress(AcctData(Accounts).RegisterName)
                 If DeviceIp = "" Then
                   See_Msg "Cannot void from this work station", 0, "Must Void from " & AcctData(Accounts).RegisterName
                   Exit Function
                 End If
               End If
              End If
              Set xml = Nothing
            End If
            Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
            If LoadEMVDevice(AcctData(Accounts).RegisterName, EmvEquip) = "Failed" Then
               See_Msg "No Device Found", 5, "Check Device Setup"
               Exit Function
            End If
            If DeviceIp > "" Then EmvEquip.Item(1).PinPadIpAddress = DeviceIp
            Dim Answer As String
            Dim SIFO As EMVProcessLib.Process
            Dim Sales As AllEMVSales
            Set Sales = New AllEMVSales
            If Sales.Item(CStr(AcctData(Accounts).CheckNo)) Is Nothing Then
              Sales.add CStr(AcctData(Accounts).CheckNo)
            End If
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).Cashier = LoggedOnAs(0).EMPName
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).Term = AcctData(Accounts).RegisterName
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).IpAddress = DeviceIp
            If PAXBARTABPREAUTH = 1 Or EMVSUPPORTPREAUTH = 1 Then
             Sales.Item(CStr(AcctData(Accounts).CheckNo)).SalesAmt = Format(AcctData(Accounts).ckTotal, "0.00")
            Else
             Sales.Item(CStr(AcctData(Accounts).CheckNo)).SalesAmt = Format(AcctData(Accounts).ckTotal + AcctData(Accounts).TipAmt, "0.00")
            End If
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).TipAmt = 0  ''Format(AcctData(Accounts).TipAmt, "0.00")
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).TaxAmt = 0
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).XMLString = AcctData(Accounts).XMLDataString
            Select Case AcctData(Accounts).ChargeType
              Case Is = 1
                Sales.Item(CStr(AcctData(Accounts).CheckNo)).TranType = "Debit"
              Case Is = 2
                Sales.Item(CStr(AcctData(Accounts).CheckNo)).TranType = "EBT"
              Case Else
                Sales.Item(CStr(AcctData(Accounts).CheckNo)).TranType = "Credit"
            End Select
            Sales.Item(CStr(AcctData(Accounts).CheckNo)).TransNo = CStr(VoidCheckNo)
            If IsPAXInterface(EMVINTERFACE) Then
              See_Msg "Searching for transaction....", 5, ""
            Else
            
            End If
            Dim PMG As PAXMSG
            Set PMG = New PAXMSG
            Set SIFO = New Process
            If IsDataCapInterface(EMVINTERFACE) Then
              Answer = SIFO.ProcessSale(Sales, DevSaleType_VoidSaleByRecordNo, PMG, EmvEquip)
            ElseIf IsDejavooInterface(EMVINTERFACE) Then
              Answer = SIFO.ProcessSale(Sales, DevSaleType_djvCapture, PMG, EmvEquip)
              Trans_Log "Void Capture " & Answer
              Answer = SIFO.ProcessSale(Sales, DevSaleType_djvVoid, PMG, EmvEquip)
              Trans_Log "Void Sale " & Answer
            Else
              Answer = SIFO.ProcessSale(Sales, DevSaleType_XVoidbyRecordNo, PMG, EmvEquip)
            End If
            Set Sales = Nothing
            PMG.Item(1).ResultTxt = VoidApprovals(PMG.Item(1).ResultTxt)
            
            If UCase(Trim$(PMG.Item(1).Message)) = "ALREADY VOIDED" Or UCase(Trim$(PMG.Item(1).Message)) = "APPROVED OFFLINE" Or UCase(Trim$(PMG.Item(1).Message)) = "ALREADY REVERSED" Then
               PMG.Item(1).ResultTxt = "APPROVAL"
            End If
            If PMG.Item(1).ResultTxt = "DECLINED" Then
               Delete_Payment_Record = True
               PMG.Item(1).ResultTxt = "APPROVAL"
               Trans_Log "Void Charge Declined Removed from Batch... Check No " & CStr(Check_no) & " auth code " & PMG.Item(1).AuthCode
               Trans_Log "Response from the processor " & Answer
            End If
            
            Trans_Log "Void Message " & PMG.Item(1).Message & " Results Returned " & PMG.Item(1).ResultTxt
            
            If IsApproved(PMG.Item(1).ResultTxt) Then
              Delete_Payment_Record = True
              If VoidAuth = False Then  'no print for void preauth bar tabs
                Set xml = New MPxmlParse
                xml.xml = PMG.Item(1).ExtraData
                Emv_Back.r_ReturnResult = PMG.Item(1).ExtraData
                If xml.token("PrintData") > "" Then
                  PrintChitSlip FormatEMVReceipt(xml.token("PrintData"))
                Else
                 Dim CCD As CCData
                    With CCD
                      .HeaderData = "Void"
                      .HeaderReason = "Void"
                      .AuthNo = PMG.Item(1).AuthCode & ""
                      .SaleAmt = Format(AcctData(Accounts).ckTotal, "0.00")
                      .TipAmt = 0
                      .AccNo = PMG.Item(1).Last4
                      .CustName = ""
                      .CType = PMG.Item(1).CardType
                      .ServerName = LoggedOnAs(0).EMPName
                      .TBLName = CStr(Check_no)
                      .ChkNO = CStr(Check_no)
                      .ProcessorResponse = "Approved"
                      .Datetime = Time
                      .ckTotal = Format(AcctData(Accounts).ckTotal, "0.00")
                    End With
                    PrintSlip CCD
                 Set xml = Nothing
               End If
              End If
              See_Msg "Payment Voided", 5, ""
            Else
               Delete_Payment_Record = False
               Trans_Log "Void Failed " & PMG.Item(1).Message & " Results Returned " & PMG.Item(1).ResultTxt
               See_Msg "Payment Voided Failed", 5, PMG.Item(1).Message
               Set SIFO = Nothing
               Set PMG = Nothing
               Set DCAC = Nothing
               Exit Function
            End If
            
            Set SIFO = Nothing
            Set PMG = Nothing
         Case Is = "QRPayment"
            If See_Msg("Void charge ", 4, "Continue ?") <> 6 Then
              Exit Function
            End If
            If See_Msg("Void Card the charge will be lost ", 4, "Are You Sure?") <> 6 Then
              Exit Function
            End If
            If Can_Continue("Managers Menu") = False Then
             Delete_Payment_Record = False
             Exit Function
            End If
            If Check_Option(246) = 1 Then
               Dim Rst As String
               See_Msg "Searching for transaction....", 5, ""
               Dim VB As GeneralLib.Utilities
               Set VB = New GeneralLib.Utilities
               Rst = VB.VoidBirdPay(CStr(VoidCheckNo))
               Trans_Log "Void QR Payment " & Rst
               Set VB = Nothing
            End If
         Case Is = "Mercury Gift", "Gift EPay New Issue", "Gift EPay ReLoad", "GiftCard", "Gift Card"
             Dim MGCD As ver11DebitCards.GiftChargeData
             Dim MGR  As ver11DebitCards.GiftCardResponse
                MGCD.AccountNumber = AcctData(Accounts).AuthCode
                MGCD.AccountRef = AcctData(Accounts).AccountName
                MGCD.PathtoFile = SvrPath.TerminalPath
                MGCD.OperatorName = AcctData(Accounts).ServerName
                MGCD.RegisterName = MyRegisterName
                MGCD.ChargeAmount = AcctData(Accounts).ckTotal
                MGCD.CheckNumber = AcctData(Accounts).CheckNo
                MGCD.AccountRef = AcctData(Accounts).CardNumber
                '// account ref added for gift local to void a gift card
                MGCD.XMLString = AcctData(Accounts).XMLDataString
                If AcctData(Accounts).AccountName = "Sell Card" Then
                  MGCD.TransactionType = VoidReturn
tryagain:
                  Delete_Payment_Record = GiftCardPayments(MGCD, MGR)
                Else
                  If AcctData(Accounts).TenderName = "Gift EPay New Issue" Then
                    MGCD.TransactionType = VoidIssue
                  Else
                    MGCD.TransactionType = VoidSale
                  End If
                  Delete_Payment_Record = GiftCardPayments(MGCD, MGR)
                End If
                If InStr(MGR.ApprovalStr, "VoidIssue") > 0 Then
                  MGCD.TransactionType = VoidIssue
                  GoTo tryagain
                End If
       End Select
     Next Accounts
      'DCAC.VoidCharge VoidCheckNo, SvrPath.TerminalPath
 End If

  ConnectADODB SvrPath.TerminalPath & DebitDb, Db
   If Check_Option(246) = 1 Then '\\ QRCode option
    Sql = "Delete from [InProcess] where [CheckNo] = " & VoidCheckNo
    Db.Execute Sql
   End If
    Sql = "Delete from [Debit] Where [Check Number] = " & VoidCheckNo
    Sql = Sql & " and [Close Out Day] = 0 "
   Db.Execute Sql
   SetNothing Db
 
ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
  If OrderIs = "Delivery" Or OrderIs = "Pick Up" Then
    Sql = "Select * from [dayPayments] where [Payment Type] = 'Deposit' and [Check Number] = " & VoidCheckNo
   ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
       DepositAmt = Val(Tb.fields("Payment Amount").Value & "")
    End If
    Tb.Close
  End If
  Sql = "Delete from [dayPayments] where [Check Number] = " & VoidCheckNo
  Db.Execute Sql
  Sql = "Delete from [Pre Auth] where [Check Number] = " & VoidCheckNo
  Db.Execute Sql
SetNothing Db
 
If OrderIs = "Delivery" Or OrderIs = "Pick Up" Then  ', [check total] = 0
  If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
   Sql = "Update [Open Checks] Set [Deposit] = 0, [Payments] = 0 Where [Check Number] = " & VoidCheckNo
  Db.Execute Sql
   Sql = "Delete [PrePaid] where [OrderNumber] = " & VoidCheckNo
  Db.Execute Sql
   Sql = "Delete [PrePaid] where [PhoneOrderNumber] = " & VoidCheckNo
  Db.Execute Sql
  SetNothing Db
  If ConnectADODB(SvrPath.TerminalPath & FinancialDb, Db) = False Then Exit Function
  Sql = "update [Payments] Set [Payment Amount] = 0 where [Check Number] = " & VoidCheckNo
  Db.Execute Sql
  If Not AllOrders Is Nothing Then
    If Not AllOrders.Item(1) Is Nothing Then
     AllOrders.Item(1).OrderPaymentAmt = 0  '// need this value to make sure preauth total dont update.
     AllOrders.Item(1).Deposits = 0
    End If
  End If
End If
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Delete from [Payments] where [Check Number] = " & VoidCheckNo
  Sql = Sql & " and [Close out day] = 0"
  Db.Execute Sql
  SetNothing Db
  Delete_Payment_Record = True
  Trans_Log "Voiding payment " & VoidCheckNo
Exit Function
ErrorHandler:
Trans_Log "Error Voiding " & Err.Description
Err.Clear
End Function
Public Function GetPaxIpAddress(TermName As String) As String
Dim Cdb As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
    Sql = "Select * from [EMVDeviceList] where [DeviceName] = '" & TermName & "'"
ConnectADORS Sql, Cdb, Tb
If Not Tb.EOF Then
  GetPaxIpAddress = Tb.fields("PinPadIpAddress").Value & ""
Else
  GetPaxIpAddress = ""
End If
DbClose Cdb, Tb
End Function
Public Function SetPaymentToOpen(Check As Long) As String
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim RC As Integer
SetPaymentToOpen = "Failed"
ConnectADODB SvrPath.TerminalPath & DebitDb, Db
    Sql = "Select * from [Debit] where [Check Number] = " & Check
ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
   Tb.fields("Server Name").Value = "OpenRecord"
   Tb.Update
   Tb.MoveNext
Loop
DbClose Db, Tb
SetPaymentToOpen = "Success"
End Function
Public Function GetSigCapture(Getcheck As Long) As String
            GetSigCapture = ""
            If IsDataCapInterface(EMVINTERFACE) = False Then Exit Function
            Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
            If LoadEMVDevice(MyRegisterName, EmvEquip) = "Failed" Then
               See_Msg "No Device Found", 5, "Check Device Setup"
               Exit Function
            End If
            'If DeviceIp > "" Then EmvEquip.Item(1).PinPadIpAddress = DeviceIp
            Dim Answer As String
            Dim SIFO As EMVProcessLib.Process
            Dim Sales As AllEMVSales
            Set Sales = New AllEMVSales
            If Sales.Item(CStr(Getcheck)) Is Nothing Then
              Sales.add CStr(Getcheck)
            End If
            Sales.Item(CStr(Getcheck)).Cashier = LoggedOnAs(0).EMPName
            Sales.Item(CStr(Getcheck)).Term = MyRegisterName
            Sales.Item(CStr(Getcheck)).IpAddress = EmvEquip.Item(1).PinPadIpAddress
            Sales.Item(CStr(Getcheck)).SalesAmt = 0
            Sales.Item(CStr(Getcheck)).TipAmt = 0
            Sales.Item(CStr(Getcheck)).TaxAmt = 0
            Sales.Item(CStr(Getcheck)).XMLString = ""
            Sales.Item(CStr(Getcheck)).TranType = "Credit"
            Sales.Item(CStr(Getcheck)).TransNo = CStr(Getcheck)
            Dim PMG As PAXMSG
            Set PMG = New PAXMSG
            Set SIFO = New Process
              GetSigCapture = SIFO.ProcessSale(Sales, DevSaleType_Getsignature, PMG, EmvEquip)
            Set Sales = Nothing
            Set SIFO = Nothing
            Set PMG = Nothing
End Function
Public Function DepositsCollected() As Currency
On Error GoTo ErrorHandler
Dim fdb     As ADODB.Connection
Dim FRs     As ADODB.Recordset
Dim Sql     As String
ConnectADODB SvrPath.TerminalPath & FinancialDb, fdb
Sql = "Select Distinct [Payment Type], Sum([Amount]) as TotalCollected from [Paid Ins] Where [Close Out Day] = 0"
Sql = Sql & " Group by [Payment Type]"
ConnectADORS Sql, fdb, FRs
If Not FRs.EOF Then
     DepositsCollected = Val(FRs![TotalCollected] & "")
Else
     DepositsCollected = 0
End If
DbClose fdb, FRs
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function CreditCardDeposits(CashierName As String) As Currency
On Error GoTo ErrorHandler
Dim CCDB As ADODB.Connection
Dim CCRS As ADODB.Recordset
Dim RN   As Integer
CreditCardDeposits = 0
   ConnectADODB SvrPath.TerminalPath & DebitDb, CCDB
    Sql = "Select Distinct [Check Number], sum([Amount]) as Ckamount "
    Sql = Sql & " From [Debit] where [Close Out day] = 0 "
    If CashierName > "" Then
     Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
    End If
    Sql = Sql & " and [check Number] = -9999"
    Sql = Sql & " Group by [check Number]"

ConnectADORS Sql, CCDB, CCRS
  If Not CCRS.EOF Then
    CreditCardDeposits = Format(CCRS![Ckamount], "0.00")
  Else
    CreditCardDeposits = 0
  End If
DbClose CCDB, CCRS
Exit Function
ErrorHandler:
DbClose CCDB, CCRS
End Function
Public Function QROrderStatus(MyCheckNo As String) As String
Dim BPResponse As String
QROrderStatus = ""
    Dim CIP As GeneralLib.Utilities
    Set CIP = New GeneralLib.Utilities
    BPResponse = CIP.CheckForBirdPayments(MyCheckNo)
    Set CIP = Nothing
    If BPResponse = "No Orders Pending" Then
      BPResponse = WaitingOnPayment(Val(MyCheckNo))
    End If
    If BPResponse = "Pending" Then
      If See_Msg("Order Pending", 4, "Try and Cancel Order?") = 6 Then
        DeleteWaiting Val(MyCheckNo)
      End If
    End If
    If BPResponse = "Paid" Then
     QROrderStatus = BPResponse
    End If
End Function
Public Function WaitingOnPayment(CheckNo As Long) As String
On Error GoTo ErrorHandler
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
WaitingOnPayment = GL.IsQRPaided(CStr(CheckNo))
Set GL = Nothing
Exit Function
ErrorHandler:
See_Msg Err.Description, 5, "Error connecting to database"
Error_Log Err.Number, Err.Description, "Error Pay Check", ""
End Function
Public Function DeleteWaiting(CheckNo As Long) As Boolean
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Rsp As String
Dim SSID As String
Dim SS As GeneralLib.Utilities
Set SS = New GeneralLib.Utilities
ConnectADODB SvrPath.TerminalPath & DebitDb, Db
  Sql = "Select * From [InProcess] where [Status] = 'Pending' and [CheckNo] = " & CheckNo
ConnectADORS Sql, Db, RS

  If Check_Option(246) = 1 Then
   Rsp = SS.CancelQRCodeWaiting(CStr(CheckNo))
   Trans_Log "Cancel QR Code Pay at the table " & Rsp
  End If
  Sql = "delete From [InProcess] where [Status] = 'Pending' and [CheckNo] = " & CheckNo
  If UCase(Rsp) = "ERROR" Then
   If See_Msg("Cancel Order Failed", 4, "Cancel Anyway?") = 6 Then
      Trans_Log "Cancel Failed Server Canceled Anyway "
   Else
    Sql = ""
   End If
  End If
Set SS = Nothing
If Sql > "" Then Db.Execute Sql
DbClose Db, RS
Exit Function
ErrorHandler:
See_Msg Err.Description, 5, "Error connecting to database"
Error_Log Err.Number, Err.Description, "Error Pay Check", ""
DbClose Db, RS
End Function
Public Function StoreTableSeqNo(SeqNo As String, CheckNo As Long, DataReturnStr As String, EMVHW As EMVProcessLib.EmvDeviceHardWare) As String   'EmvDeviceHardWare
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & DebitDb, Db
  Sql = "Select * From [InProcess] where [PayAtTheTableId] = '" & SeqNo & "'"
ConnectADORS Sql, Db, RS
  If RS.EOF Then RS.AddNew
  RS.fields("StartTime").Value = Now
  RS.fields("CheckNo").Value = CheckNo
  RS.fields("ServerName").Value = LoggedOnAs(0).EMPName
  RS.fields("TerminalId").Value = MyRegisterName
  RS.fields("PayAtTheTableId").Value = SeqNo
  RS.fields("PinPadIpAddress").Value = EMVHW.Item(1).PinPadIpAddress
  RS.fields("PinPadIpPort").Value = EMVHW.Item(1).PinPadIpPort
  RS.fields("PinPadMacAddress").Value = ""
  RS.fields("Response").Value = EMVHW.Item(1).deviceId
  RS.fields("Status").Value = "Pending"
  RS.fields("Processor").Value = "SlingShot"
  RS.fields("CheckAmount").Value = 0
  RS.fields("XMLString").Value = ""
  RS.Update
DbClose Db, RS
See_Msg "PinPad Success", 5, "Transaction Sent for " & CStr(CheckNo)
Exit Function
ErrorHandler:
See_Msg Err.Description, 5, "Error connecting to database"
Error_Log Err.Number, Err.Description, "Error On Store Pay at The Table", ""
DbClose Db, RS
End Function
Public Function CashDeposits(CashierName As String) As Currency
On Error Resume Next
Dim Cr  As Integer
Dim fdb As ADODB.Connection
Dim FRs As ADODB.Recordset
CashDeposits = 0  ': NonCashDeposit = 0
     ConnectADODB SvrPath.TerminalPath & FinancialDb, fdb
      Sql = "Select distinct sum([Payment Amount]) as PayMentTotal from [dayPayments] where [check number] = -9999 and [Payment Type] = 'Cash'"
      If Check_Option(60) = 0 Or Check_Option(34) = 1 Then
        Sql = Sql & " and [Cashier Name] = '" & CashierName & "'"
      End If
      Sql = Sql & " and [Cashier Name]  <> 'Closed'"
      Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
     ConnectADORS Sql, fdb, FRs
      If Not FRs.EOF Then CashDeposits = CashDeposits + Val(FRs![PaymentTotal] & "")
   DbClose fdb, FRs
End Function
Public Function FindTare(TareName As String) As Single
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
Sql = "Select * from [Tare] Where [Tare Name] = '" & TareName & "'"
ConnectADORS Sql, Db, RS
  If Not RS.EOF Then
     FindTare = Val(RS(1) & "")
  Else
     FindTare = 0
  End If
DbClose Db, RS
End Function
Public Function OCDSleep() As Integer
OCDSleep = Val(GetConfigInfo("Misc", "OCD Sleep Seconds", 500, "Functions.ini"))
End Function
Public Function ModifierCategoryPreFix(ByCategory As String) As String
ModifierCategoryPreFix = GetConfigInfo("Modifier Prefix", ByCategory, "", "Functions.ini")
End Function
Public Function GetBarCodeLen() As Integer
GetBarCodeLen = Val(GetConfigInfo("Bar Codes", "Bar Code Length", 6, "Functions.ini"))
If GetBarCodeLen = 0 Then GetBarCodeLen = 6
End Function
Public Function GetScaleLen() As Integer
GetScaleLen = Val(GetConfigInfo("NCI", "Scale Data Length", 16, "Functions.ini"))
If GetScaleLen = 0 Then GetScaleLen = 16
End Function
Public Function GetScaleModel() As String
GetScaleModel = GetConfigInfo("NCI", "Scale Model No", "NCI", "Functions.ini")
End Function
Public Function GetCheckDigitLen() As Integer
GetCheckDigitLen = Val(GetConfigInfo("Bar Codes", "Check Digit Length", 1, "Functions.ini"))
If GetCheckDigitLen = 0 Then GetCheckDigitLen = 1
End Function
Public Function GetPricePos() As Integer
GetPricePos = Val(GetConfigInfo("Bar Codes", "Type 2 Bar Code Price Position", 6, "Functions.ini"))
If GetPricePos = 0 Then GetPricePos = 6
End Function
Public Function IsValueTax(ValueIn As Integer) As Integer
'Inclusive Tax by Timed Event=
If GetConfigInfo("Quick Bar", "Inclusive Tax by Timed Event", "", "Functions.ini") > "" Then
  If ValueIn = 1 Then
   If CheckTimedEvent(GetConfigInfo("Quick Bar", "Inclusive Tax by Timed Event", "", "Functions.ini")) Then
      IsValueTax = 1
   Else
      IsValueTax = 0
   End If
  Else
    IsValueTax = ValueIn
  End If
Else
    If Check_Option(163) = 1 Then  'Disable Inclusive Tax
       IsValueTax = 0
    Else
       IsValueTax = ValueIn
    End If
End If
End Function
Public Function GetBergLen() As Integer
GetBergLen = Val(GetConfigInfo("Misc", "Berg PLU Code Length", 0, "Functions.ini"))
'GetBergLen = 4
End Function
Public Function GetOrderDest(ButtonName As String) As String
GetOrderDest = GetConfigInfo("Order Destination By Button", ButtonName, "", "Functions.ini")
End Function
Public Function DiscountTotal(TypeDisc As String, TQty As Integer, WTD As Integer) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Reg As Integer
DiscountTotal = 0: TQty = 0
        ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
        Sql = "Select distinct [Department Report], Sum(Price) as CouponTTl,"
        Sql = Sql & " Count([Department Report]) as CouponCt"
        Sql = Sql & " From [check detail] where  Voided = 0 "                     ' [Check Status] = 0"
        If WTD = 0 Then
          Sql = Sql & " and [Close out day] = 0"
        End If
        If TypeDisc = "Coupon" Then
           Sql = Sql & " and [Department Report] = 'Coupon'"
        ElseIf TypeDisc = "Frequent Diner" Then
           Sql = Sql & " and [Department Report] = 'Frequent Diner'"
        Else
           Sql = Sql & " and [Department Report] = 'Discount'"
           Sql = Sql & " Or [Department Report] = 'Combo'"
        End If
        Sql = Sql & " Group by [Department Report]"
        ConnectADORS Sql, Db, RS
             If Not RS.EOF Then
                DiscountTotal = DiscountTotal + Abs(RS!CouponTTl)
                TQty = TQty + RS!Couponct
             End If
        DbClose Db, RS
   
 
End Function
Public Function GetTabNameThisCheck(CheckNo As Long) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
  GetTabNameThisCheck = "NoOne"
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
        Sql = "Select * from [Open Checks] Where [Close out day] = 0 and [Check Number] = " & CheckNo
  ConnectADORS Sql, Db, RS
     If Not RS.EOF Then
        GetTabNameThisCheck = RS.fields("Table Number").Value & ""
     End If
DbClose Db, RS
End Function
Public Function GetTheCheckOwner(Operator As String, CurrentCheck As Long) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
  GetTheCheckOwner = Operator
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
        Sql = "Select * from [Open Checks] Where [Close out day] = 0 and [Check Number] = " & CurrentCheck
  ConnectADORS Sql, Db, RS
     If Not RS.EOF Then
        GetTheCheckOwner = RS.fields("Server Name").Value
     End If
DbClose Db, RS

End Function
Public Function GetWhoOwnsThisCheck(CheckNo As Long) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
  GetWhoOwnsThisCheck = "NoOne"
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
        Sql = "Select * from [Open Checks] Where [Close out day] = 0 and [Check Number] = " & CheckNo
  ConnectADORS Sql, Db, RS
     If Not RS.EOF Then
        GetWhoOwnsThisCheck = RS.fields("Server Name").Value
     End If
DbClose Db, RS
End Function
Public Function GetWhoOwnsTheCheck(WhoIsRunningTheTerminal As String, CheckNo As Long) As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
  GetWhoOwnsTheCheck = WhoIsRunningTheTerminal
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
        Sql = "Select * from [Closed Checks] Where [Close out day] = 0 and [Check Number] = " & CheckNo
  ConnectADORS Sql, Db, RS
     If Not RS.EOF Then
        GetWhoOwnsTheCheck = RS.fields("Server Name").Value
     End If
DbClose Db, RS
End Function
Public Function GetWhoClosedCheck() As String
On Error Resume Next
  If Check_Option(213) = 1 Then
    GetWhoClosedCheck = MyRegisterName
  Else
    'GetWhoClosedCheck = LoggedOnAs(0).EMPName
    GetWhoClosedCheck = LoggedOnAs(0).EmpLoggedOnAs
  End If
'07/24/2001 allows tracking of who closed servers checks
'Charged 750.00 for this function Viking Sales (Duane sales man)
End Function
Public Function TabOver(ItemName As Variant, TPos As Long) As Long
  TabOver = TPos - Len(ItemName)
  If TabOver < 0 Then TabOver = Len(ItemName)
  If Val(Input_Option(30)) < 33 Then
    If TabOver - 16 < 0 Then TabOver = 16
    'TabOver = TabOver + 5
  Else
    TabOver = TabOver + 16
  End If
End Function
Public Function PMS_GetPostAll() As String
If FileCheck(SvrPath.TerminalPath & "INI Files\" & "PMS_PostAll.INI") Then
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
Dim RData()       As String
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "PMS_PostAll.INI"), TXTFileInfo()) Then
    For Wl = 0 To UBound(TXTFileInfo, 1)
      RData = Split(TXTFileInfo(Wl), "=")
      PMS_PostAllSales = RData(1)
      If UCase(PMS_PostAllSales) = "ON" Then
         PMS_PostAllSales = "ENABLED"
      Else
         PMS_PostAllSales = "DISABLED"
      End If
    Next Wl
 End If
End If
'PMS_DeptSale
End Function
Public Function PMS_GetVoidInq() As String
If FileCheck(SvrPath.TerminalPath & "INI Files\" & "PMS_VoidChargeInq.INI") Then
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
Dim RData()       As String
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "PMS_VoidChargeInq.INI"), TXTFileInfo()) Then
    For Wl = 0 To UBound(TXTFileInfo, 1)
      RData = Split(TXTFileInfo(Wl), "=")
      PMS_SendVoidInq = RData(1)
    Next Wl
 End If
End If
End Function
Public Function PMS_GetRevCenter(CenterName As String) As String
PMS_GetRevCenter = ""
If FileCheck(SvrPath.TerminalPath & "INI Files\" & "PMS_RevenueCenter.INI") Then
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
Dim RData()       As String
 If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "PMS_RevenueCenter.INI"), TXTFileInfo()) Then
    For Wl = 0 To UBound(TXTFileInfo, 1)
      If InStr(TXTFileInfo(Wl), CenterName) > 0 Then
         RData = Split(TXTFileInfo(Wl), "=")
         PMS_GetRevCenter = RData(1)
      End If
    Next Wl
 End If
End If
End Function
Public Function PMS_GetDeptSales() As String
If FileCheck(SvrPath.TerminalPath & "INI Files\" & "PMS_DepartmentSales.INI") Then
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
Dim RData()       As String
 If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "PMS_DepartmentSales.INI"), TXTFileInfo()) Then
    For Wl = 0 To UBound(TXTFileInfo, 1)
         If InStr(TXTFileInfo(Wl), "=") > 0 Then
          RData = Split(TXTFileInfo(Wl), "=")
          PMS_DeptSales(Val(RData(1)) - 1).DeptName = RData(0)
          PMS_DeptSales(Val(RData(1)) - 1).Dept = 0
        End If
    Next Wl
 End If
End If
'PMS_DeptSale
End Function
Public Function PMS_GetTender() As String
If FileCheck(SvrPath.TerminalPath & "INI Files\" & "PMS_Tenders.INI") Then
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
Dim RData()       As String
Dim W             As Integer
 
 If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "PMS_Tenders.INI"), TXTFileInfo()) Then
    W = -1
    For Wl = 0 To UBound(TXTFileInfo, 1)
     If InStr(TXTFileInfo(Wl), "=") > 0 Then
      RData = Split(TXTFileInfo(Wl), "=")
      W = W + 1
      PMS_Tender(W).DeptName = Trim$(RData(0))
      PMS_Tender(W).Dept = 0
      PMS_Tender(W).TenderNo = RData(1)
     End If
    Next Wl
 End If
End If
'PMS_DeptSale
End Function
Public Function PMS_GetPostTender() As String
If FileCheck(SvrPath.TerminalPath & "INI Files\" & "PMS_PostTenders.INI") Then
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
Dim RData()       As String
Dim W             As Integer
 If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "PMS_PostTenders.INI"), TXTFileInfo()) Then
    W = -1
    For Wl = 0 To UBound(TXTFileInfo, 1)
     If InStr(TXTFileInfo(Wl), "=") > 0 Then
      W = W + 1
      RData = Split(TXTFileInfo(Wl), "=")
      PMS_PostTender(W).DeptName = Trim$(RData(0))
      PMS_PostTender(W).Dept = 0
      PMS_PostTender(W).TenderNo = RData(1)
     End If
    Next Wl
 End If
End If
'PMS_DeptSale
End Function
Public Function ReturnLabel(IncommingText As String) As String
  SaveLanguageText IncommingText
  If CustOption.Translator = "Off" Then
    ReturnLabel = IncommingText
    Exit Function
  End If
 On Error GoTo ERH
    If AllLanguage.Item(Trim$(IncommingText)) Is Nothing Then
       ReturnLabel = IncommingText
    Else
       ReturnLabel = AllLanguage.Item(Trim$(IncommingText)).TranslatedTxt
    End If
    If ReturnLabel = "" Then ReturnLabel = IncommingText
Exit Function
ERH:
ReturnLabel = IncommingText
'this function added 08/09/2001 to handler multi language
'incomming = english and return what ever.
End Function
Public Sub LoadTranslations()
    If CustOption.Translator = "Off" Then Exit Sub
    
    Set AllLanguage = Nothing
    Set AllLanguage = New AllLanguages
    Dim Db As ADODB.Connection, RS As ADODB.Recordset
    OpenDb SvrPath.LocalTerminalPath & "MSaleData", Db
      Sql = "Select * from [Language] "
    ConnectADORS Sql, Db, RS
     Do While Not RS.EOF
       If AllLanguage.Item(RS![English] & "") Is Nothing Then
         AllLanguage.add(RS![English] & "").TranslatedTxt = RS![Return Text] & ""
       End If
      RS.MoveNext
     Loop
    DbClose Db, RS
End Sub
Public Function SeatOrTray() As SeatName
      If OrderIs = "Pick Up" Or OrderIs = "Go Order" Or OrderIs = "Delivery" Then
         SeatOrTray.SeatName = "Bag  "
         SeatOrTray.AddName = ""
         Exit Function
      End If
      If Check_Option(0) = 0 Then
       SeatOrTray.SeatName = "Seat "
       SeatOrTray.AddName = "Joiner "
      Else
       SeatOrTray.SeatName = "Tray "
       SeatOrTray.AddName = "Tray "
      End If
End Function
Public Sub SaveLanguageText(InTx As String)
    If ExportToDBase = 0 Then Exit Sub
    If InTx = "" Then Exit Sub
    Dim Db As ADODB.Connection, RS As ADODB.Recordset
    OpenDb SvrPath.LocalTerminalPath & "MSaleData", Db
      Sql = "Select * from [Language] Where [English] = '" & InTx & "'"
    ConnectADORS Sql, Db, RS
    If RS.EOF Then
       RS.AddNew
       RS![Return Text] = ""
       RS![English] = InTx & ""
       RS.Update
    End If
    DbClose Db, RS
End Sub
Public Function GetChargeType(ByRef ChargeType As SelectChargeType) As String
   Select Case ChargeType
      Case Is = 1
          GetChargeType = "Debit"
      Case Is = 2
          GetChargeType = "Credit"
      Case Is = 3
          GetChargeType = "House Account"
      Case Is = 4
          GetChargeType = "Gift Card"
   End Select
End Function
Public Function Is_Cash(TenderType As String) As Integer
 On Error Resume Next
 Dim SC As Integer
 If TenderType = "Deposit" Then Is_Cash = 1: Exit Function
 If TenderType = "Cash" Then Is_Cash = 1: Exit Function
 For SC = 0 To UBound(CreditCards, 1) - 1
     If TenderType = CreditCards(SC).CrdName Then
       Is_Cash = CreditCards(SC).CashTender
       Exit Function
     End If
  Next SC
End Function
Public Function OK_ToOpenDrawer(TenderType As String) As Integer
 On Error Resume Next
 Dim SC As Integer
 If TenderType = "Deposit" Then OK_ToOpenDrawer = 1: Exit Function
 'If TenderType = "Cash" Then OK_ToOpenDrawer = 1: Exit Function
 For SC = 0 To UBound(CreditCards, 1) - 1
     If TenderType = CreditCards(SC).CrdName Then
       OK_ToOpenDrawer = CreditCards(SC).OpensDrawer
       Exit Function
     End If
  Next SC
End Function
Public Function Round2Money(ByVal x As Variant) As Currency
    Dim v As Double
    v = Val(CStr(x))
    Round2Money = CCur(Sgn(v) * Int(Abs(v) * 100# + 0.5) / 100#)
End Function


Public Function PlaceDec(ByVal SaleAmt As Variant) As Variant
    On Error GoTo Fail

    Dim s As String, sign As String
    Dim i As Long, ch As String
    Dim Cleaned As String, hasDot As Boolean
    Dim out As String

    If IsNull(SaleAmt) Then
        PlaceDec = "0.00": Exit Function
    End If

    s = Trim$(CStr(SaleAmt))
    If LenB(s) = 0 Then
        PlaceDec = "0.00": Exit Function
    End If

    ' Preserve a leading minus
    If Left$(s, 1) = "-" Then
        sign = "-"
        s = Mid$(s, 2)
    End If

    ' Keep only digits and the first decimal point (strip $, commas, spaces, etc.)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            Cleaned = Cleaned & ch
        ElseIf ch = "." And Not hasDot Then
            If Len(Cleaned) = 0 Then Cleaned = "0"
            Cleaned = Cleaned & "."
            hasDot = True
        End If
    Next

    If LenB(Cleaned) = 0 Then
        PlaceDec = "0.00": Exit Function
    End If

    If hasDot Then
        ' Normalize any decimal input to 2 places (rounded)
        If Right$(Cleaned, 1) = "." Then Cleaned = Cleaned & "00"
        PlaceDec = Format$(CCur(sign & Cleaned), "0.00")
    Else
        ' No decimal typed: follow your original behavior:
        ' 1–2 digits => treat as whole dollars; 3+ => insert before last two
        Select Case Len(Cleaned)
            Case 1, 2
                out = Cleaned & ".00"
            Case Else
                out = Left$(Cleaned, Len(Cleaned) - 2) & "." & Right$(Cleaned, 2)
        End Select
        If sign = "-" Then out = "-" & out
        PlaceDec = out
    End If
    Exit Function

Fail:
    PlaceDec = "0.00"
End Function

Public Function GroupCat() As String
'On Error Resume Next
GroupCat = "None"
If CPath.TerminalFailed Then Exit Function
If TerminalFailed(SvrPath.TerminalPath) Then
 CPath.TerminalFailed = True
 Exit Function
Else
 CPath.TerminalFailed = False
End If
Dim TR   As ADODB.Connection
Dim Trs  As ADODB.Recordset
OpenDb SvrPath.TerminalPath & TimeRecordsDb, TR
 Sql = "Select * from [Employee Records] Where [Employee Name] = '" & LoggedOnAs(0).EMPName & "' and [Time In] > '' and [Time Out] = ''"
ConnectADORS Sql, TR, Trs
  If Not Trs.EOF Then
     GroupCat = Trs![Department Name]
  Else
     GroupCat = "None"
  End If
DbClose TR, Trs
End Function
Public Function GroupCatOff(Db As ADODB.Connection) As Boolean
GroupCatOff = True
Dim RS   As ADODB.Recordset
Sql = "select * from [Group Category] Where [Button Level] > -1 and [Group Name] > ''"
ConnectADORS Sql, Db, RS
  If Not RS.EOF Then
     GroupCatOff = False
  Else
     GroupCatOff = True
  End If
  SetNothing RS
End Function

Public Function LogOnCode() As String
On Error GoTo ErrorHandler
LogOnCode = "None"
If TerminalIsPresent(SvrPath.TerminalName) = False Then
 CPath.TerminalFailed = True
 Exit Function
End If
Dim Db As ADODB.Connection, RS As ADODB.Recordset
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
Sql = "Select * from [Employee Records]"
Sql = Sql & " Where [Employee Name] = '" & LoggedOnAs(0).EMPName & "' and [Time in] > '' and [Time Out] = ''"
ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
      LogOnCode = Check_Null(RS![Department Name])
   End If
DbClose Db, RS
Exit Function
ErrorHandler:
  Select Case Err
     Case Is = 3040, 3043, 3044
      If CheckPaths(SvrPath.TerminalPath) = False Then
       'AutoSwitchCommon True, TimeRecordsDb
       SwitchToBackup Common
      End If
      'SecondaryServer.SVRDown.Caption = "Cannot Connect To Time Records"
      'SecondaryServer.Show 1
      'Resume
  End Select
End Function
Public Function CheckPaths(PathToTerminal As String) As Boolean
Dim ttpath      As SetPaths
Set ttpath = New SetPaths
    PathToTerminal = ttpath.ReturnCorrectPath(PathToTerminal, "")
    If ttpath.RegisterOn(PathToTerminal) Then
       CheckPaths = True
    Else
       CheckPaths = False
    End If
Set ttpath = Nothing
End Function
Public Function CheckTax(SeatNo As Integer) As Boolean
On Error GoTo ErrorHandler
Dim TR          As Integer
Dim ResetSTax   As Boolean
Dim LT          As Currency
ResetSTax = False
   With Order.Ticket(0)
     For TR = 0 To NextRow(SeatNo) - 1
       If .TextMatrix(TR, 26) <> "Service Charge" Then  '// added for alitos 10.09.2025
       '// not sure about this function for service charge
        If Val(.TextMatrix(TR, 32)) = 0 And Val(.TextMatrix(TR, 2)) > 0 Then
         If Val(.TextMatrix(TR, 10)) > 0 Or Val(.TextMatrix(TR, 11)) > 0 Then   'Regular sales tax and special tax
           ResetSTax = True
           Exit For
         End If
       End If
      End If
     Next TR
   End With
If ResetSTax = False Then
    Exit Function
Else
   If Check_Option(70) = 0 Then
   With Order.Ticket(0)
     For TR = 0 To NextRow(SeatNo) - 1
       .TextMatrix(TR, 32) = 0
     Next TR
   End With
   End If
End If
Exit Function
ErrorHandler:
CheckTax = False
End Function
Public Sub ResetBarTax(SeatNo As Integer)
On Error GoTo ErrorHandler
Dim TR As Integer
Dim LT As Currency
   With Order.Ticket(0)
     For TR = 0 To NextRow(SeatNo) - 1
       .TextMatrix(TR, 32) = 0
     Next TR
   End With
Exit Sub
ErrorHandler:
End Sub
Public Function GetComboMatch(ElementNo As Long) As Long
Dim C As Long
Dim R As Long
R = -1
For C = 0 To UBound(ComboMatch, 2)
  If ComboMatch(ElementNo, C).MatchNo = ElementNo Then
    R = R + 1
  End If
Next C
GetComboMatch = R
End Function
Public Function GetExpandPrint(GMenuItem As Variant) As Long
On Error GoTo ErrorHandler
GetExpandPrint = AllMenuItems(GMenuItem).ExpandPrint
Exit Function
ErrorHandler:
End Function
Public Function Match(TargetItem As String, FindDisc As String, PrimayItem As String, LowestPrice As Currency, MatchLines() As CouponMatches) As Boolean
 With Order.Ticket(0)
    Dim M1 As Integer
    Dim MM As Integer
    Dim MustMatch As Long
    Dim AllGroups As Long
    Dim ML As Integer
    Match = False
    Dim m  As Integer
    MustMatch = GetMatchNo(FindDisc)
    Dim Matches As Long
    ReDim MatchLines(MustMatch)
    Dim GroupMatched() As Integer
    ReDim GroupMatched(6)
    Dim MDisc  As StoreDiscounts
    Dim ItemD  As ItemDiscounts
    Set MDisc = New StoreDiscounts
    Set ItemD = New ItemDiscounts
    If MDisc.Item(TargetItem & PrimayItem) Is Nothing Then
       MDisc.add TargetItem & PrimayItem
    End If 'allows item to match with same item
    Matches = 0
     For MM = 0 To NextRow(SeatNumber) - 1
       If Val(.TextMatrix(MM, 2)) >= 0 Then
         If .TextMatrix(MM, 17) = "None" Then   'blocked for combo meals
          For AllGroups = 1 To 6
           For M1 = 0 To GetComboMatch(AllGroups)
              If Trim(FindDisc) = Trim(ComboMatch(AllGroups, M1).DiscountName) Then
               If MDisc.Item(.TextMatrix(MM, 1) & .TextMatrix(MM, 14)) Is Nothing Then
                If .TextMatrix(MM, 1) = ComboMatch(AllGroups, M1).MatchWith And .TextMatrix(MM, 22) = "" And GroupMatched(AllGroups) = 0 Then
                 For m = 0 To Matches
                   If .TextMatrix(MM, 1) = MatchLines(m).ItemName Then GoTo NM1
                   DiscountGroup = DiscountGroup + Val(.TextMatrix(MM, 2))
                   MatchLines(Matches).LineNo = MM
                   MatchLines(Matches).ItemName = .TextMatrix(MM, 1)
                   MatchLines(Matches).ItemNo = .TextMatrix(MM, 14)
                   GroupMatched(AllGroups) = AllGroups
                   Matches = Matches + 1
                   If Matches >= MustMatch Then
                     For ML = 0 To Matches - 1
                     .TextMatrix(MatchLines(ML).LineNo, 22) = TargetItem
                     If LowestPrice > .TextMatrix(MatchLines(ML).LineNo, 2) Then
                       LowestPrice = .TextMatrix(MatchLines(ML).LineNo, 2)
                     End If
                     Next ML
                     Match = True
                     Set MDisc = Nothing
                     Set ItemD = Nothing
                     Exit Function
                   End If
                 Next m
                End If
              End If
           End If
NM1:
         Next M1
        Next AllGroups
      End If
    End If
    Next MM
End With
Set MDisc = Nothing
Set ItemD = Nothing
End Function
Public Function ODCIsConnected(Disconnect As Boolean) As Boolean
On Error GoTo ERH
ODCIsConnected = False
If Check_Option(50) = 1 Then Exit Function
Dim Db As ADODB.Connection
Dim RS  As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
If Disconnect Then
  Sql = "Delete from [LoggedOnWhere] where [Server Name] = 'Connected To OCD'"
  Db.Execute Sql
  SetNothing Db
  Exit Function
Else
  Sql = "Select * from [LoggedOnWhere] where [Server Name] = 'Connected To OCD'"
End If
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
If Not RS.EOF Then
     If RS.fields("Register Name").Value <> MyRegisterName Then
      DbClose Db, RS
      ODCIsConnected = True
      Exit Function
    End If
Else
    RS.AddNew
    RS![Server Name] = "Connected To OCD"
    RS![Register Name] = MyRegisterName
    RS.Update
End If
DbClose Db, RS
Exit Function
ERH:
Err.Clear
End Function
Public Function MatchCombo(DiscName As String, Combodisc As Currency, PrimaryNumber As String, LowestPrice As Currency, Taxes As CouponTaxes) As Boolean
   Dim FindMatch   As String
   Dim SR          As Integer
   Dim MC          As Integer
   Dim M1          As Integer
   Dim M2          As Integer
   Dim m           As Integer
   Dim FindDisc    As String
   Dim MatchLine() As CouponMatches
   DiscName = "": Combodisc = 0
   DiscountGroup = 0  'used modgroupprice recycle var for use here.
     With Order.Ticket(0)
    For SR = 0 To NextRow(SeatNumber) - 1

            If Val(.TextMatrix(SR, 2)) >= 0 Then  'i am here working down
            PrimaryNumber = .TextMatrix(SR, 14)
            If Trim(.TextMatrix(SR, 17)) = "None" Then
            FindMatch = .TextMatrix(SR, 1)
            For MC = 0 To UBound(ComboDiscounts(), 1) - 1
                DiscountGroup = 0  'Need to reset for out of order search
                If Trim(FindMatch) = ComboDiscounts(MC).DiscountMatch Then ' Or InStr(Trim(FindMatch), ComboDiscounts(MC).DiscountMatch) > 0 Then
                  FindDisc = ComboDiscounts(MC).DiscountName
                  DiscountGroup = DiscountGroup + Val(.TextMatrix(SR, 2))
                  If LowestPrice > Val(.TextMatrix(SR, 2)) Or LowestPrice = 0 Then
                     LowestPrice = Val(.TextMatrix(SR, 2))
                  End If
                  If ValidDay(FindDisc) = False Then GoTo Nxt
                  If IsValidCoupon(FindDisc) = False Then GoTo Nxt 'Exit Function
                  If CountCoupons(FindDisc, ComboDiscounts(MC).DicountLimit) = False Then GoTo Nxt  ' Exit Function
                  ''// need to find next auto discount if you exit function stop search
                  '// fails when using day qualified discounts  05.10.2016
                    
                    
                    If Match(FindMatch, FindDisc, PrimaryNumber, LowestPrice, MatchLine()) = True Then
                        If Trim(ComboDiscounts(MC).DiscountType) <> "Combo Percentage" Then LowestPrice = 0
                        If Trim(ComboDiscounts(MC).DiscountType) = "Combo Set Amount" Then
                         Combodisc = Abs(Val(ComboDiscounts(MC).DiscountAmt) - DiscountGroup)
                         LowestPrice = 0
                         DiscountGroup = 0
                        Else
                         Combodisc = Val(ComboDiscounts(MC).DiscountAmt)
                        End If
                        DiscName = ComboDiscounts(MC).DiscountName
                        CouponPLU = ComboDiscounts(MC).PLU
                        CouponDepartment = ComboDiscounts(MC).DiscountDept
                        Taxes.Tax1 = ComboDiscounts(MC).DiscTax1 * -1
                        Taxes.Tax2 = ComboDiscounts(MC).DiscTax2 * -1
                        Taxes.Tax3 = ComboDiscounts(MC).DiscTax3 * -1
                        .TextMatrix(SR, 17) = DiscName & "-" & PrimaryNumber
                        MatchCombo = True
                        For m = 0 To NextRow(SeatNumber) - 1
                          If .TextMatrix(m, 22) = FindMatch Then
                            .TextMatrix(m, 17) = DiscName & "-" & PrimaryNumber
                            .TextMatrix(m, 22) = ""
                          End If
                        Next m
                        If ComboDiscounts(MC).ComboType = "Combo Item" Then
                         Make_It_a_Combo MatchLine(), Val(PrimaryNumber), Combodisc
                         MatchCombo = False
                        End If
                       Exit Function
                    End If
                  End If
            Next MC
          End If
        End If
Nxt:
     Next SR
    End With
End Function
Public Sub Make_It_a_Combo(Matched() As CouponMatches, ItemNo As Integer, PriceAdj As Currency)
On Error GoTo ERH
Dim m       As Integer
Dim f       As Integer
Dim SN      As Integer
Dim AL      As Integer
Dim L       As Integer
Dim PcUp    As Currency
Dim FL      As Integer
Dim SL      As Integer
PcUp = 0
Store_Grid Order.Ticket(0)
With Order.Ticket(0)
  For AL = 0 To (UBound(Matched, 1)) - 1
     For m = 0 To NextRow(SeatNumber) - 1
       If Val(.TextMatrix(m, 14)) = Matched(AL).ItemNo Then
         .RemoveItem m
       End If
     Next m
   Next AL
   For m = 0 To NextRow(SeatNumber) - 1
     If .TextMatrix(m, 14) = ItemNo Then
       FL = m
       SL = Val(.TextMatrix(m, 15))
       For f = m To NextRow(SeatNumber) - 1
         If Val(.TextMatrix(f, 14)) <> ItemNo Then
            For AL = 0 To UBound(Matched, 1) - 1
              .AddItem " ", f
            Next AL
            L = -1
            For AL = f To (f + UBound(Matched, 1)) - 1
                L = L + 1
                For SN = 0 To TTL_Col
                  .TextMatrix(AL, SN) = Grid_array(SeatNumber, Matched(L).LineNo, SN)
                  If SN = 14 Then .TextMatrix(AL, 14) = ItemNo
                  If SN = 15 Then .TextMatrix(AL, 15) = SL
                  
                  If SN = 2 Then
                    PcUp = PcUp + Val(Grid_array(SeatNumber, Matched(L).LineNo, SN))
                    .TextMatrix(AL, 2) = ""
                    .TextMatrix(AL, 0) = ""
                  End If
                Next SN
            Next AL
            Exit For
         End If
       Next f
       Exit For
     End If
   Next m
   .TextMatrix(FL, 1) = .TextMatrix(FL, 1) & " Combo"
   .TextMatrix(FL, 2) = Format((Val(.TextMatrix(FL, 2)) + PcUp) - PriceAdj, "0.00")
End With
Sleep 0
Store_Grid Order.Ticket(0)
Order.Load_Grid
Order.Amount_Tend_Update
Exit Sub
ERH:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Auto Combo ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub CheckSoftwareLicense()
   On Error GoTo ErrorHandler
   Dim LicAgreement As String
   Dim Cl As GeneralLib.Utilities
   Set Cl = New GeneralLib.Utilities
    LicAgreement = Cl.GetKeyData("EulaAccepted")
   Set Cl = Nothing
    If LicAgreement = "false" Then
      See_Msg "No License agreement -- Program will end", 0, "License Agreement Required To Operate."
      End
    End If
Exit Sub
ErrorHandler:
Trans_Log "Problem with license agreement " & Err.Description
End Sub

Public Function GetMatchNo(DiscountName As String) As Long
   On Error GoTo ErrorHandler
   Dim sDb As ADODB.Connection
   Dim Srn As ADODB.Recordset
   ConnectADODB SvrPath.LocalTerminalPath & "MSaleData", sDb
   Sql = "Select top 1 [Match No] From [Match] where [Discount Name] = '" & DiscountName & "' order by [Match No] Desc"
   ConnectADORS Sql, sDb, Srn
   If Not Srn.EOF Then
      GetMatchNo = Srn![Match No]
   Else
      GetMatchNo = 1
   End If
   DbClose sDb, Srn
Exit Function
ErrorHandler:
GetMatchNo = 1
End Function
Public Function ValidDay(DiscountName As String) As Boolean
On Error Resume Next
ValidDay = True
 If Not CouponRestriction.Item(DiscountName & (DatePart("w", Date, 2, 1) - 1)) Is Nothing Then
   ValidDay = False
 Else
  ValidDay = True
 End If
End Function
Public Sub SaveCharge(CheckNo As Long, SrvName As String, ChargeAmt As Currency, TipAmt As Currency)
On Error GoTo ERH
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Sub
          Sql = "Select * from [payments] where [Close Out Day] = 0 and [check Number] = " & CheckNo
          Sql = Sql & " and [Payment Type] = 'Paid Credit'"
          Sql = Sql & " and [Payment Amount] = " & ChargeAmt
          If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
          If RS.EOF Then
             RS.AddNew
             RS![Check Number] = CheckNo
             RS![Payment Type] = "Paid Credit"
             RS![Payment Amount] = Format(ChargeAmt, "0.00")
             RS![Date] = Format(Date, "MM/DD/YY")
             RS![Time] = Format(Time, "HH:MM:SS ampm")
             RS![Close out day] = 0
           End If
           RS![Tip Amount] = 0
           RS.Update
          DbClose Db, RS
Exit Sub
ERH:
Err.Clear
DbClose Db, RS
End Sub
Public Function GetAuthAmt(fdb As ADODB.Connection, Chk As Long) As Currency
On Error Resume Next
Dim FRs As ADODB.Recordset
If fdb Is Nothing Then Exit Function
    Sql = "Select Distinct sum([Amount]) as TAuth from [Batch] where [Close Out day] = 0 and [check number] = '" & CStr(Chk) & "'"
    Sql = Sql & " and [Pre Auth] = 1 and [Status] <> 'Void' Group by [check number]"
   ConnectADORS Sql, fdb, FRs
  If Not FRs.EOF Then
     GetAuthAmt = Format(FRs![TAuth], "0.00")
  Else
     GetAuthAmt = 0
  End If
SetNothing FRs
End Function
Public Function GetStartDate(WhichDay As Integer) As String
    On Error Resume Next
    Dim Db As ADODB.Connection, Tb As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    If WhichDay = 7 Then
      Sql = "Select * from [Check Detail] Where [Close Out Day] = 0 order by TRY_CAST([check Date] AS datetime) asc"
    Else
      Sql = "Select * from [Check Detail] order by TRY_CAST([check Date] AS datetime) asc"
    End If
    ConnectADORS Sql, Db, Tb
            If Tb.EOF Then
               GetStartDate = CStr(Format(Date, "MM/DD/YYYY"))
            Else
               GetStartDate = Tb![Check Date]
            End If
DbClose Db, Tb
End Function
Public Function GetLocationName(AccountName As String) As Long
On Error GoTo ErrorHandler
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
GetLocationName = 0
If AccountName = "" Then Exit Function
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & AccountName & "' or [Scan Name] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & AccountName & "' Or [Spc Interest 2] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & AccountName & "' Or [Spc Interest 4] = '" & AccountName & "'"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
If Not Tb Is Nothing Then
   If Not Tb.BOF Then Tb.MoveFirst
Else
  See_Msg "Cannot Connect to House Account", 5, ""
  Exit Function
End If
If Tb.EOF Then
    Call See_Msg("Information not found", 5, "")
    DbClose Nothing, Tb
    Set GRS = Nothing
    Exit Function
Else
    GetLocationName = Val(Tb![Meal Plan] & "")
End If
DbClose Nothing, Tb
Set GRS = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Location Name ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function

Public Function GetCreditBalance(AccountName As String) As Integer
On Error GoTo ErrorHandler
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
GetCreditBalance = 0
If AccountName = "" Then Exit Function
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & AccountName & "' or [Scan Name] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & AccountName & "' Or [Spc Interest 2] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & AccountName & "' Or [Spc Interest 4] = '" & AccountName & "'"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
If Not Tb Is Nothing Then
   If Not Tb.BOF Then Tb.MoveFirst
Else
  See_Msg "Cannot Connect to House Account", 5, ""
  Exit Function
End If
If Tb.EOF Then
    Call See_Msg("Information not found", 5, "")
    DbClose Nothing, Tb
    Set GRS = Nothing
    Exit Function
Else
    GetCreditBalance = Tb![Credit Balance]
End If
DbClose Nothing, Tb
Set GRS = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Credit Balance ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetAccountBalance(AccountName As String) As Currency
On Error GoTo ErrorHandler
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim BegBal As Currency
If AccountName = "" Then Exit Function
Dim GCard As String
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & AccountName & "' or [Scan Name] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & AccountName & "' Or [Spc Interest 2] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & AccountName & "' Or [Spc Interest 4] = '" & AccountName & "'"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
If Not Tb Is Nothing Then
   If Not Tb.BOF Then Tb.MoveFirst
Else
  See_Msg "Cannot Connect to House Account", 5, ""
  Exit Function
End If

If Tb.EOF Then
    Call See_Msg("Information not found", 5, "")
    DbClose Nothing, Tb
    Set GRS = Nothing
    GetAccountBalance = 0
    Exit Function
Else
    AccountName = Tb![scan Name]
    GCard = Tb![Account Type] & ""
    
    SetNothing Tb
    Sql = "Select Distinct sum([beginning Balance]) as StartBal from [Starting Balance]"
    Sql = Sql & " Where [Account Name] = '" & AccountName & "'"
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "beginning Balance")
    If Not Tb Is Nothing Then
       If Not Tb.BOF Then Tb.MoveFirst
    Else
      See_Msg "Cannot Connect to House Account", 5, ""
      Exit Function
    End If
    If Not Tb.EOF Then
        BegBal = Val(Tb!StartBal & "")
    Else
        BegBal = 0
    End If
    SetNothing Tb
    Dim PaymentsMade As Currency, BalanceOwed As Currency
    PaymentsMade = 0: BalanceOwed = 0
    Sql = "Select Distinct sum([Adj total]) as BalanceDue, Sum([Tip Amount]) as TTip "
    Sql = Sql & "From [HouseCharges] where [charged to] = '" & AccountName & "'"
    Sql = Sql & " and [Voided] = 0"
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Charge Data")
    If Not Tb Is Nothing Then
       If Not Tb.BOF Then Tb.MoveFirst
    Else
      See_Msg "Cannot Connect to House Account", 5, ""
      Exit Function
    End If
    If Not Tb.EOF Then BalanceOwed = Format(Val(Check_Null(Tb(0))), "0.00")   '+ Val(TB![TTip] & "")
    SetNothing Tb

    Sql = "Select Distinct sum([Amount]) as PaymentDue "
    Sql = Sql & "From [HousePayments] where [Account Name] = '" & AccountName & "'"
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Payments")
    If Not Tb Is Nothing Then
       If Not Tb.BOF Then Tb.MoveFirst
    Else
      See_Msg "Cannot Connect to House Account", 5, ""
      Exit Function
    End If
    
    If Not Tb.EOF Then PaymentsMade = Format(Val(Check_Null(Tb(0))), "0.00")
    DbClose Nothing, Tb
    GetAccountBalance = (PaymentsMade + (BegBal * -1)) - BalanceOwed
    Set GRS = Nothing
End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Account Balance ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetExpirationDate(AccountName As String) As String
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim BegBal As Currency
If AccountName = "" Then Exit Function
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & AccountName & "' or [Scan Name] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & AccountName & "' Or [Spc Interest 2] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & AccountName & "' Or [Spc Interest 4] = '" & AccountName & "'"

If ConnectADODB(HouseAcct, Db) = False Then
  Exit Function
End If

If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
        If Not Tb.EOF Then
            GetExpirationDate = Tb![Account Expires On] & ""
        End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Expiration Date ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ReturnAllAccountCharges(AccountName As String) As Currency
On Error GoTo ErrorHandler
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Tb  As ADODB.Recordset
Dim BegBal As Currency
ReturnAllAccountCharges = 0
If AccountName = "" Then Exit Function
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & AccountName & "' or [Scan Name] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & AccountName & "' Or [Spc Interest 2] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & AccountName & "' Or [Spc Interest 4] = '" & AccountName & "'"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
    If Not Tb Is Nothing Then
       If Not Tb.BOF Then Tb.MoveFirst
     Else
       See_Msg "error connection", 5, ""
     End If
     If Tb.EOF Then
            Call See_Msg("Information not found", 5, "")
            ReturnAllAccountCharges = 0
            DbClose Nothing, Tb
            Set GRS = Nothing
            Exit Function
        Else
            AccountName = Tb![scan Name]
            SetNothing Tb
            Sql = "Select Distinct sum([beginning Balance]) as StartBal from [Starting Balance]"
            Sql = Sql & " Where [Account Name] = '" & AccountName & "'"
            'If ConnectADORS(Sql, DB, Tb) = False Then Exit Function
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Starting Balance")
            If Not Tb.EOF Then
                BegBal = Val(Check_Null(Tb!StartBal))
            Else
                BegBal = 0
            End If
            SetNothing Tb
            Dim PaymentsMade As Currency, BalanceOwed As Currency
            PaymentsMade = 0: BalanceOwed = 0
            Sql = "Select Distinct sum([Adj total]) as BalanceDue, Sum([Tip Amount]) as TTip "
            Sql = Sql & "From [HouseCharges] where [charged to] = '" & AccountName & "'"
            Sql = Sql & " and [Voided] = 0"
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Charge Data")

            If Not Tb.EOF Then BalanceOwed = Format(Val(Check_Null(Tb(0))), "0.00")   '+ Val(TB![TTip] & "")
            SetNothing Tb
        
            Sql = "Select Distinct sum([Amount]) as PaymentDue "
            Sql = Sql & "From [HousePayments] where [Account Name] = '" & AccountName & "'"
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Payments")
            If Not Tb.EOF Then PaymentsMade = Format(Val(Check_Null(Tb(0))), "0.00")
            DbClose Nothing, Tb
            ReturnAllAccountCharges = (PaymentsMade + Abs(BegBal)) - BalanceOwed
        End If
   Set GRS = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Return All Account Charges ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ReturnAccountCharges(AccountName As String, StartDate As String, EndDate As String) As Currency
On Error GoTo ErrorHandler
Dim BegBal As Currency
Dim AccountNo As String
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Tb As ADODB.Recordset
If InStr(AccountName, "'") > 0 Then Exit Function
If AccountName = "" Then Exit Function
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & AccountName & "' or [Scan Name] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & AccountName & "' Or [Spc Interest 2] = '" & AccountName & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & AccountName & "' Or [Spc Interest 4] = '" & AccountName & "'"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
        If Tb.EOF Then
            Call See_Msg("Information not found", 5, "")
            ReturnAccountCharges = 0
            DbClose Nothing, Tb
            Set GRS = Nothing
            Exit Function
        Else
            AccountName = Tb![scan Name]
            AccountNo = Tb![Acct Number] & ""
            
            Sql = "Select Distinct sum([beginning Balance]) as StartBal from [Starting Balance]"
            Sql = Sql & " Where [Account Name] = '" & AccountName & "'"
            'If Input_Option(50) = "Account #" Then
            ' Sql = Sql & " And [Account Number] = '" & AccountNo & "'"
            'End If
            'If ConnectADORS(Sql, DB, Tb) = False Then Exit Function
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Starting Balance")

            If Not Tb.EOF Then
                BegBal = Val(Check_Null(Tb!StartBal))
            Else
                BegBal = 0
            End If
            'If Not Tb Is Nothing Then SetNothing Tb: Set Tb = Nothing
            Dim PaymentsMade As Currency, BalanceOwed As Currency
            PaymentsMade = 0: BalanceOwed = 0
            Sql = "Select Distinct sum([Adj total]) as BalanceDue, Sum([Tip Amount]) as TTip "
            Sql = Sql & "From [HouseCharges] where [charged to] = '" & AccountName & "'"
            'If Input_Option(50) = "Account #" Then
            '  Sql = Sql & " and [Scan Number] = '" & AccountNo & "'"
            'End If
            Sql = Sql & " and [Voided] = 0"
            Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & CDate(StartDate) & "' and '" & CDate(EndDate) & "'"
            'If ConnectADORS(Sql, DB, Tb) = False Then Exit Function
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Charge Data")

            If Not Tb.EOF Then BalanceOwed = Format(Val(Check_Null(Tb(0))), "0.00")   '+ Val(TB![TTip] & "")
            'If Not Tb Is Nothing Then SetNothing Tb
        
            Sql = "Select Distinct sum([Amount]) as PaymentDue "
            Sql = Sql & " From [HousePayments] where [Account Name] = '" & AccountName & "'"
            'If Input_Option(50) = "Account #" Then
            '  Sql = Sql & " and [Account Number] = '" & AccountNo & "'"
            'End If
            Sql = Sql & " and TRY_CAST([Date] AS datetime) between '" & CDate(StartDate) & "' and '" & CDate(EndDate) & "'"
            'If ConnectADORS(Sql, DB, Tb) = False Then Exit Function
            Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Payments")
            If Not Tb.EOF Then PaymentsMade = Format(Val(Check_Null(Tb(0))), "0.00")
            DbClose Nothing, Tb
            ReturnAccountCharges = (BalanceOwed - PaymentsMade) + BegBal
            Set GRS = Nothing
        End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Return Account Charges " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetNonTaxSales(WtDNonTaxedSales As Currency, TaxedSales As Currency) As Currency
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim TAvg        As Integer
Dim NOR         As Integer
Dim RN          As Integer
Dim FD          As Integer
GetNonTaxSales = 0: TaxedSales = 0

    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "select Distinct Sum([Price]) as NtSale, sum([Inclusive Amount]) as Tax from [Check Detail] where [Voided] = 0 and [Tax Table] = 0"
    Sql = Sql & " and [Special tax] = 0 And [Close out day] = 0"
    Sql = Sql & " and [Liquor tax] = 0 And [Close out day] = 0"
    Sql = Sql & " and [Deleted] = 0 "  '// fixed non taxed sales 10.24.2023  -- added liquor tax to the mix
    
    For FD = 1 To AllExemptSales.count
      Sql = Sql & " and [Department Report] <> '" & AllExemptSales.Item(FD).ExemptItem & "'"
    Next FD
    ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then GetNonTaxSales = GetNonTaxSales + (Val(Tb![NtSale] & "") - Val(Tb.fields("Tax").Value & ""))
    SetNothing Tb
    ' taxed sales
    Sql = "select Distinct Sum([Price]) as NtSale, sum([Inclusive Amount]) as Tax from [Check Detail] where [Voided] = 0 and [Tax Table] <> 0"
    
    Sql = Sql & " and [Close out day] = 0"
    Sql = Sql & " and [Deleted] = 0 "
    For FD = 1 To AllExemptSales.count
      Sql = Sql & " and [Department Report] <> '" & AllExemptSales.Item(FD).ExemptItem & "'"
    Next FD
    Sql = Sql & " Or [Voided] = 0 and [Special tax] <> 0"
    Sql = Sql & " and [Close out day] = 0"
    Sql = Sql & " and [Deleted] = 0 and [Department Report] <> 'Gratuity' "
    ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then TaxedSales = TaxedSales + (Val(Tb![NtSale] & "") - Val(Tb![Tax] & ""))
    SetNothing Tb
    Sql = "select Distinct Sum([price]) as WTDNtSale,Sum([Inclusive Amount]) as Tax from [Check Detail] where [Voided] = 0 and [Tax Table] = 0"
    Sql = Sql & "And [Special tax] = 0 "
    Sql = Sql & "And [liquor tax] = 0 "
    Sql = Sql & " and [Deleted] = 0 "
    Sql = Sql & " and [Price] > 0 "
    For FD = 1 To AllExemptSales.count
      Sql = Sql & " and [Department Report] <> '" & AllExemptSales.Item(FD).ExemptItem & "'"
    Next FD
    ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then WtDNonTaxedSales = WtDNonTaxedSales + (Val(Tb![WTDNtSale] & "") - Val(Tb![Tax] & ""))
DbClose Db, Tb
Exit Function
ErrorHandler:
See_Msg "error " & Err.Description, 5, ""
Exit Function
End Function
Public Function ServerTipSales(SvrName As String, ShiftDate As String, ShiftCheckNo As Long) As Currency
On Error GoTo ErrorHandler
ServerTipSales = 0
'*************
Dim sDb As ADODB.Connection
Dim SRS As ADODB.Recordset
     ConnectADODB SvrPath.TerminalPath & CheckStatDb, sDb
       Sql = "Select Distinct sum([Check total]) as SSales, Sum([Gross Exempt]) as GE, Sum([Voided Exempt]) as VE, "
       Sql = Sql & " sum([Gross Sales Tax]) as TTX, Sum([Voided Sales Tax]) as VSTX ,"
       Sql = Sql & " sum([Gross Special Tax]) as STTX, Sum([Voided Special Tax]) as VSSTX "
       Sql = Sql & " From [Closed Checks]"
       Sql = Sql & " Where [Server Name] = '" & SvrName & "'"
       Sql = Sql & " and [Check Number] >= " & ShiftCheckNo
       Sql = Sql & " and [Close Out Day] = 0"
       Sql = Sql & " group by [Server Name]"
    ConnectADORS Sql, sDb, SRS
     If Not SRS.EOF Then
        ServerTipSales = ServerTipSales + Val(Check_Null(SRS!SSales)) - (Val(SRS![ge] & "") - Val(SRS![ve] & ""))
        ServerTipSales = ServerTipSales - Val(SRS!TTx & "") '- Val(SRS![VSTx] & ""))
        ServerTipSales = ServerTipSales - Val(SRS!STTX & "") '- Val(SRS![VSSTX] & ""))
     End If  'make sure sales tax is not reported in server tip sales
   DbClose sDb, SRS
   If ServerTipSales < 0 Then ServerTipSales = 0
  
If ServerTipSales = 0 Then
  ConnectADODB SvrPath.TerminalPath & "MSaleData", sDb
    Sql = "Select * from [Tip Income] where [Server Name] = '" & SvrName & "'"
    Sql = Sql & " and TRY_CAST([Date] AS date) = '" & CDate(ShiftDate) & "'"
  ConnectADORS Sql, sDb, SRS
    If Not SRS.EOF Then
     ServerTipSales = Val(SRS![Tip Sales] & "")
    End If
    DbClose sDb, SRS
    If ServerTipSales < 0 Then ServerTipSales = 0
End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Tip Sales " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ReturnCreditCardSales(SvrName As String, ShiftCheckNo As Long) As String
On Error GoTo ErrorHandler
'*************
Dim sDb As ADODB.Connection
Dim SRS As ADODB.Recordset
ReturnCreditCardSales = 0
'If EMVDeviceIsPresent Then
  ConnectADODB SvrPath.TerminalPath & DebitDb, sDb
    Sql = "Select Distinct [Server Name], sum(Amount) as SSales, Sum([Tip]) as TAmt From [Debit]"
    Sql = Sql & " Where [Server Name] = '" & SvrName & "'"
    Sql = Sql & " and [check Number] >= " & ShiftCheckNo
    Sql = Sql & " and [Close Out Day] = 0"
    Sql = Sql & " and [Tender] = 'EMV'"
    Sql = Sql & " Or [Server Name] = '" & SvrName & "'"
    Sql = Sql & " and [check Number] >= " & ShiftCheckNo
    Sql = Sql & " and [Close Out Day] = 0"
    Sql = Sql & " and [Tender] = 'QRPayment'"
    
    Sql = Sql & " Group By [Server Name]"
 ConnectADORS Sql, sDb, SRS
    If Not SRS.EOF Then
       ReturnCreditCardSales = Val(Check_Null(SRS!SSales))
    Else
       ReturnCreditCardSales = 0
    End If
'End If
DbClose sDb, SRS
Exit Function
ErrorHandler:
DbClose sDb, SRS
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Return Credit Card Sales " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetShiftNo(SvrName As String) As Long
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
 ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
  Sql = "Select * From [Check Numbers] where [Servers Name] = '" & SvrName & "' and [Audit Status] = 0"
 ConnectADORS Sql, Db, RS
   If RS.EOF Then
     GetShiftNo = 0
   Else
     GetShiftNo = Val(RS![Check No] & "")
   End If
 DbClose Db, RS
End Function
Public Sub RemoveNullDates()
    On Error GoTo ErrorHandler

    Dim SL As ADODB.Connection
    OpenDb SvrPath.TerminalPath & "MSaleData", SL

    Dim Sql As String
    Sql = "DELETE FROM [Tip Income] WHERE [Date] IS NULL"

    SL.Execute Sql
    DbClose SL, Nothing
    Exit Sub

ErrorHandler:
    ' Handle gracefully
    MsgBox "Error cleaning Tip Income dates: " & Err.Description, vbExclamation
    DbClose SL, Nothing
End Sub
Public Function ReturnTipsPaidOut(SvrName As String, BeginDate As String, EndDate As String) As Currency
On Error GoTo ErrorHandler
'*************
If BeginDate = "" Or EndDate = "" Then Exit Function
'RemoveNullDates
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
OpenDb SvrPath.TerminalPath & "MSaleData", SL
    Sql = "Select Distinct [Server Name],"
    Sql = Sql & " Sum([Tips Paid Out]) as TipsPaidOut"
    Sql = Sql & " from [Tip Income] where [Server Name] = '" & SvrName & "'"
    Sql = Sql & " and TRY_CAST([Date] AS date)  Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
 '   Sql = Sql & " and Not [date] Isnull"
    Sql = Sql & " Group By [Server Name]"
ConnectADORS Sql, SL, SRd
     If Not SRd.EOF Then
       ReturnTipsPaidOut = Val(SRd![TipsPaidOut] & "")
     Else
       ReturnTipsPaidOut = 0
     End If
DbClose SL, SRd
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Return Tips Paid Out " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ReturnGratuityCollected(SvrName As String, BeginDate As String, EndDate As String) As Currency
On Error GoTo ErrorHandler
'*************
If BeginDate = "" Or EndDate = "" Then Exit Function
'RemoveNullDates
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & "MSaleData", SL
    Sql = "Select Distinct [Server Name],Sum([Declared Tips]) as TDecl,"
    Sql = Sql & " Sum([Tips Paid Out]) as TChrg, Sum([Gratuity]) as TGrat"
    Sql = Sql & " from [Tip Income] where [Server Name] = '" & SvrName & "'"
    Sql = Sql & " and TRY_CAST([Date] AS date)  Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
    'Sql = Sql & " and Not Isnull([date])"
    Sql = Sql & " Group By [Server Name]"
ConnectADORS Sql, SL, SRd
     If Not SRd.EOF Then
       ReturnGratuityCollected = Val(SRd![TGrat] & "")
     Else
       ReturnGratuityCollected = 0
     End If
DbClose SL, SRd
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Return Gratuity Collected " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub GetPayRollInfo(ServerSalesData As ServerPayRollData, ServerName As String, BeginDate As String, EndDate As String, ShiftNo As Long)
On Error GoTo ErrorHandler
'*************
ServerSalesData.TipSales = 0
'ServerSalesData.DeclaredTips = 0
ServerSalesData.GratutityCollected = 0
ServerSalesData.TipsPaidOut = 0
'ServerSalesData.CreditCardTips = 0
If BeginDate = "" Or EndDate = "" Then Exit Sub
Dim Ar As Integer
Dim SL As ADODB.Connection, SRd As ADODB.Recordset
 ConnectADODB SvrPath.TerminalPath & "MSaleData", SL
    Sql = "Select Distinct [Server Name],Sum([Declared Tips]) as TDecl,"
    Sql = Sql & " Sum([Tip Sales]) as TSales,"
    Sql = Sql & " Sum([Charged Tips]) as TipsCharged,"
    Sql = Sql & " Sum([Tips Paid Out]) as TPaidOut, Sum([Gratuity]) as TGrat"
    Sql = Sql & " from [Tip Income] where [Server Name] = '" & ServerName & "'"
    Sql = Sql & " and TRY_CAST([Date] AS date)  Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
    'If Format(BeginDate, "MM/DD/YYYY") = Format(EndDate, "MM/DD/YYYY") Then
    Sql = Sql & " and [Shift] = " & ShiftNo
    'End If
    Sql = Sql & " Group By [Server Name]"
  ConnectADORS Sql, SL, SRd
     If Not SRd.EOF Then
       ServerSalesData.TipSales = Val(SRd!TSales & "")
       ServerSalesData.DeclaredTips = Val(SRd!TDecl & "")
       ServerSalesData.GratutityCollected = Val(SRd!TGrat & "")
       ServerSalesData.TipsPaidOut = Val(SRd!TPaidout & "")
       ServerSalesData.CreditCardTips = Val(SRd!TipsCharged & "")
     End If
DbClose SL, SRd
'**********
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Sales " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function GetServerGratuity(ServerName As String, BeginDate As String, EndDate As String, ShiftSales As Currency) As Currency
On Error GoTo ErrorHandler
'*************
GetServerGratuity = 0
'If InStr(ServerName, "Yu") > 0 Then Stop
If BeginDate = "" Or EndDate = "" Then Exit Function
Dim Ar As Integer
Dim SL As ADODB.Connection, SRd As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & "MSaleData", SL
    Sql = "Select Distinct [Server Name],Sum([Declared Tips]) as TDecl,"
    Sql = Sql & " Sum([Tip Sales]) as TSales,"
    Sql = Sql & " Sum([Charged Tips]) as TipsCharged,"
    Sql = Sql & " Sum([Tips Paid Out]) as TPaidOut, Sum([Gratuity]) as TGrat"
    Sql = Sql & " from [Tip Income] where [Server Name] = '" & ServerName & "'"
    'If ShiftSales > 0 Then
    ' Sql = Sql & " and [Tip Sales] = " & ShiftSales
    'End If
    Sql = Sql & " and TRY_CAST([Date] AS date)  Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
    'Sql = Sql & " and Not Isnull([date])"
    Sql = Sql & " Group By [Server Name]"
ConnectADORS Sql, SL, SRd
     If Not SRd.EOF Then
       GetServerGratuity = Val(SRd!TGrat & "")
     End If
DbClose SL, SRd
'**********
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Sales Grat")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ServerSales(ServerName As String, BeginDate As String, EndDate As String) As Currency
On Error GoTo ErrorHandler
ServerSales = 0
'*************
If BeginDate = "" Or EndDate = "" Then Exit Function
RemoveNullDates
Dim Ar As Integer
Dim sDb As ADODB.Connection, SRS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & "MSaleData", sDb

    Sql = "Select Distinct [Server Name],Sum([Tip Sales]) as TSales from [Tip Income] where [Server Name] = '" & ServerName & "'"
    Sql = Sql & " and TRY_CAST([Date] AS date)  Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
    'Sql = Sql & " and Not Isnull([date])"
    Sql = Sql & " Group By [Server Name]"
ConnectADORS Sql, sDb, SRS
     If Not SRS.EOF Then
       ServerSales = Val(SRS!TSales & "")
     Else
       ServerSales = 0
     End If
DbClose sDb, SRS
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Sales " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetTare(ItemTare As String) As Single
On Error Resume Next
If UCase(ItemTare) = "OPEN" Then
 GetTare = -1
 Exit Function
End If
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
Sql = "Select * from [Tare] Where [Tare Name] = '" & ItemTare & "'"
ConnectADORS Sql, Db, RS
  If Not RS.EOF Then
     GetTare = Val(RS.fields("Tare Wt").Value & "")
  Else
     GetTare = 0
  End If
DbClose Db, RS
End Function
Public Function UnlockedCheckNumber(CheckNo As Long) As Boolean
On Error GoTo ErrorHandler
If Check_Option(50) = 1 Then Exit Function
Dim Db As ADODB.Connection
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
 Sql = "Delete from [LoggedOnWhere] where [Server Name] = '" & CStr(CheckNo) & "'"
Db.Execute Sql
DbClose Db, Nothing
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Sub ReleaseRegisterLock()
On Error GoTo ERH
If Check_Option(50) = 1 Then Exit Sub
Dim Db As ADODB.Connection
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Delete from [LoggedOnWhere] Where [Register Name] = '" & MyRegisterName & "'"
Db.Execute Sql
DbClose Db, Nothing
Exit Sub
ERH:
Err.Clear
End Sub
Public Function IsServerOn(ByVal SvrName As String) As Boolean
    On Error GoTo ERH

    IsServerOn = False
    If Check_Option(50) = 1 Or SvrPath.GroupAudit Then Exit Function

    Dim Db As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim regName As Variant

    Set Db = New ADODB.Connection
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db

    ' 1) Fast scalar lookup: who is logged on for this server?
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = Db
        .CommandType = adCmdText
        .CommandText = "SELECT TOP 1 [Register Name] " & _
                       "FROM [LoggedOnWhere] " & _
                       "WHERE [Server Name] = ?"
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 64, SvrName)
    End With

    regName = cmd.Execute(, , adExecuteNoRecords) ' <-- NOTE: this returns a Recordset, not scalar in VB6

    ' ADO VB6 doesn't return a scalar directly from Execute; handle with a tiny recordset:
    Dim RS As ADODB.Recordset
    Set RS = cmd.Execute
    If Not RS.EOF Then
        regName = RS.fields(0).Value
    Else
        regName = Null
    End If
    RS.Close
    Set RS = Nothing

    If Not IsNull(regName) Then
        If CStr(regName) <> MyRegisterName Then
            See_Msg SvrName & " Is Logged On @ " & CStr(regName), 5, ""
            Trans_Log SvrName & " Is logged On " & CStr(regName)
            DbClose Db, Nothing
            IsServerOn = True
            Exit Function
        End If
        ' If it's already logged on by THIS register, treat as OK and exit.
        DbClose Db, Nothing
        Exit Function
    End If

    ' 2) Insert only if missing (fast, avoids updatable recordset)
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = Db
        .CommandType = adCmdText
        .CommandText = _
            "INSERT INTO [LoggedOnWhere] ([Server Name], [Register Name]) " & _
            "SELECT ?, ? " & _
            "WHERE NOT EXISTS (SELECT 1 FROM [LoggedOnWhere] WHERE [Server Name] = ?)"
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 64, SvrName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 64, MyRegisterName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 64, SvrName)
        .Execute , , adExecuteNoRecords
    End With

    DbClose Db, Nothing
    Exit Function

ERH:
    Trans_Log "Error " & Err.Description & " IsServerOn"
    On Error Resume Next
    DbClose Db, Nothing
    Err.Clear
End Function

Public Sub ReleaseAllLocks()
    On Error GoTo ERH
    If Check_Option(50) = 1 Then Exit Sub

    Dim Db As ADODB.Connection
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db

    Db.Execute "TRUNCATE TABLE dbo.[LoggedOnWhere]", , adExecuteNoRecords

    DbClose Db, Nothing
    Exit Sub

ERH:
    Err.Clear
    On Error Resume Next
    DbClose Db, Nothing
End Sub
Public Function NumberInUse(CheckNo As Long) As Boolean
On Error GoTo ERH
NumberInUse = False
If Check_Option(50) = 1 Then Exit Function
'// option 50 disable check for multi users -- used for drive thru systems.
'// option 50 used for drive thru system
Dim Db As ADODB.Connection
Dim RS  As ADODB.Recordset
 ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Select * from [LoggedOnWhere] where [Server Name] = '" & CStr(CheckNo) & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
   If RS![Register Name] <> MyRegisterName Then
      See_Msg "Check Is Open On " & RS![Register Name] & "", 5, "Check Confict"
      DbClose Db, RS
      NumberInUse = True
      Exit Function
    End If
Else
    RS.AddNew
    RS![Server Name] = CStr(CheckNo)
    RS![Register Name] = MyRegisterName
    RS.Update
End If
DbClose Db, RS
Exit Function
ERH:
Err.Clear
End Function
Public Sub RegisterGratuity(ByCashierName As String, GetCashOwed As FinancialItems)
On Error GoTo ErrorHandler
'*************
Dim sDb As ADODB.Connection
Dim SRS As ADODB.Recordset
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, sDb
    Sql = "Select Distinct [Servers Name],Count([Price]) as TotalGratuity, sum(price) as GratutityCharged From [Check Detail]"
    Sql = Sql & " where [Voided] = 0 "
    If Check_Option(213) = 1 Then  'Cash by Terminal
       Sql = Sql & " and [Cashier Name] <> 'Closed'"
       Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
    Else
       Sql = Sql & " and [Cashier Name] = '" & ByCashierName & "'"
    End If
    Sql = Sql & " and [Department Report] = 'Gratuity'"
    Sql = Sql & " and [Close Out Day] = 0"
    Sql = Sql & " Group By [Servers Name]"
 ConnectADORS Sql, sDb, SRS
Do While Not SRS.EOF
  If GetCashOwed.Item(SRS![Servers Name] & "") Is Nothing Then
     GetCashOwed.add SRS![Servers Name] & ""
  End If
  GetCashOwed.Item(SRS![Servers Name] & "").ItemName = SRS![Servers Name] & ""
  GetCashOwed.Item(SRS![Servers Name] & "").ItemQty = Val(SRS![TotalGratuity] & "")
  GetCashOwed.Item(SRS![Servers Name] & "").Amount = Val(SRS![GratutityCharged] & "")
  SRS.MoveNext
Loop
DbClose sDb, SRS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Gratutity " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub RegisterErrors(StartCheckNo As Long, GetErrors() As TermErrors, Cashier As String)
On Error GoTo ErrorHandler
'*************
Dim sDb As ADODB.Connection
Dim SRS As ADODB.Recordset
Dim R   As Integer
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, sDb
    Sql = "Select Distinct [Manager],[Cashier],[Event], Count([Event]) as TotalEvents From [Error Correct]"
    If Check_Option(60) = 0 Or Check_Option(34) = 1 Then
     Sql = Sql & " Where [Cashier] = '" & Trim$(Cashier) & "'"
    Else
     Sql = Sql & " Where [check Number] > " & StartCheckNo - 1
    End If
    Sql = Sql & " and [Register] = '" & MyRegisterName & "'"
    Sql = Sql & " and [Close Out Day] = 0"
    Sql = Sql & " Group By [Manager],[Cashier],[Event]"
 ConnectADORS Sql, sDb, SRS
ReDim GetErrors(RecordsCount(SRS))
R = -1
Do While Not SRS.EOF
  R = R + 1
  GetErrors(R).Cashier = SRS![Cashier] & ""
  GetErrors(R).Managers = SRS![Manager] & ""
  GetErrors(R).qty = SRS![TotalEvents] & ""
  GetErrors(R).Events = SRS![Event] & ""
  SRS.MoveNext
Loop
DbClose sDb, SRS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Gratutity " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub RegisterServiceCharge(ByCashierName As String, GetCashOwed As FinancialItems)
On Error GoTo ErrorHandler
'*************
Dim sDb As ADODB.Connection
Dim SRS As ADODB.Recordset
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, sDb
    Sql = "Select Distinct [Servers Name],Count([Price]) as TotalGratuity, sum(price) as GratutityCharged From [Check Detail]"
    Sql = Sql & " where [Voided] = 0 "
    If Check_Option(213) = 1 Then  'Cash by Terminal
      Sql = Sql & " and [Cashier Name] <> 'Closed'"
      Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
    Else
      Sql = Sql & " and [Cashier Name] = '" & ByCashierName & "'"
    End If
    Sql = Sql & " and [Department Report] = 'Service Charge'"
    Sql = Sql & " and [Cateory Name] <> 'Processing Fee'"
    Sql = Sql & " and [Close Out Day] = 0"
    Sql = Sql & " Group By [Servers Name]"
 ConnectADORS Sql, sDb, SRS
 Do While Not SRS.EOF
  If GetCashOwed.Item(SRS![Servers Name] & "") Is Nothing Then
     GetCashOwed.add SRS![Servers Name] & ""
  End If
  GetCashOwed.Item(SRS![Servers Name] & "").ItemQty = Val(SRS![TotalGratuity] & "")
  GetCashOwed.Item(SRS![Servers Name] & "").ItemName = SRS![Servers Name] & ""
  GetCashOwed.Item(SRS![Servers Name] & "").Amount = Val(SRS![GratutityCharged] & "")
  SRS.MoveNext
 Loop
DbClose sDb, SRS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Gratutity " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function CheckForUsedDiscount(Sel_Number As Integer) As Boolean
 CheckForUsedDiscount = False
 '// Sel_Number is the menu item number assignment on the check
 '// search check and make sure discount has not be used
 Dim CA As Integer
 For CA = 0 To Order.Disc_List(1).ListCount - 1
   If Sel_Number = Order.Disc_List(1).List(CA) Then
    CheckForUsedDiscount = True
    Exit Function
   End If
 Next CA
End Function
Public Function GetCheckGratutity() As Currency
If GratuityOn = False Then Exit Function

If Val(Input_Option(40)) = 0 Then
      Exit Function
End If
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
Store_Grid Order.Ticket(0)
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Gratuity" Then
       PerCentBack = GetCheckGrat(ProcessItemName(CStr(Grid_array(SrcFor, Rw, 1)), True), MaxGrat)
       If PerCentBack > 0 Then
        If Input_Option(56) = "Net Sales" Then
          Grid_array(SrcFor, Rw, 2) = Format((((Abs(Seat_Sales_Data(SrcFor, 0)) + Abs(Seat_Sales_Data(SrcFor, 4))) - Seat_Sales_Data(SrcFor, 14)) - Val(Grid_array(SrcFor, Rw, 2))) * (PerCentBack / 100), "#.00")
        ElseIf Input_Option(56) = "Discounted" Then
          Grid_array(SrcFor, Rw, 2) = Format((((Seat_Sales_Data(SrcFor, 0) - Seat_Sales_Data(SrcFor, 14)) - Val(Grid_array(SrcFor, Rw, 2)))) * (PerCentBack / 100), "#.00")
        Else
          Grid_array(SrcFor, Rw, 2) = Format(Abs(Seat_Sales_Data(SrcFor, 13)) * (PerCentBack / 100), "#.00")
        End If
        If MaxGrat > 0 And Val(Grid_array(SrcFor, Rw, 2)) > MaxGrat Then Grid_array(SrcFor, Rw, 2) = Format(MaxGrat, "#.00")
       End If
     End If
 Next Rw
Next SrcFor
Order.Load_Grid
End Function
Public Sub GetProcessingFeeEvent(EventName As String)
On Error GoTo ErrorHandler
Dim DiscDb            As ADODB.Connection
Dim Tb                As ADODB.Recordset
Dim DCName            As String
Dim OpenCharge        As Single
Dim DiscAmount        As Single

ConnectADODB SvrPath.LocalTerminalPath & "MSaleData", DiscDb
  Sql = "Select * from [Discount Events] where [Event Name] = '" & EventName & "'"
  ConnectADORS Sql, DiscDb, Tb
  If Not Tb.EOF Then
    DCName = Tb![Discount Name] & ""
  Else
   DbClose DiscDb, Tb
   See_Msg "No Event", 5, "Cannot Add Item"
   Exit Sub
  End If
  Tb.Close
  Sql = "Select * from [Discounts] where [Discount Name] = '" & DCName & "'"
  ConnectADORS Sql, DiscDb, Tb
   If Tb.EOF Then
     See_Msg "Item Not Found", 5, "Cannot process Processing Fee Missing"
     DbClose DiscDb, Tb
     Exit Sub
   End If

   DiscAmount = Val(Tb.fields("Discount Amount").Value & "")

   If EventName = "Processing Fee" Then
    If DiscAmount > 3# Then DiscAmount = 3#
   End If
   
   
   Sales_Totals   'seat_sales_data(Seatnumber,0) = subtotal
    Select Case Tb![Discount Type] & ""
     Case Is = "Charge Percentage"
       OpenCharge = (Seat_Sales_Data(SeatNumber, 0) * DiscAmount / 100) * 100 / 100
     Case Is = "Charge Amount"
       OpenCharge = DiscAmount
     Case Is = "Gratuity"
       OpenCharge = (Seat_Sales_Data(SeatNumber, 0) * DiscAmount / 100) * 100 / 100
     
     Case Else
       See_Msg "Item Not Found", 5, "Charge Amount Not Found"
       DbClose DiscDb, Tb
       Exit Sub
     End Select
    Dim R As Integer
    R = NextRow(SeatNumber)
    Grid_array(SeatNumber, R, 0) = 1
    Grid_array(SeatNumber, R, 1) = Tb![Discount Name] & ""
    Grid_array(SeatNumber, R, 20) = ""
    Grid_array(SeatNumber, R, 14) = AdvanceSelection
    Grid_array(SeatNumber, R, 24) = SeatNumber
    Grid_array(SeatNumber, R, 15) = 1000    'sort order
    Grid_array(SeatNumber, R, 16) = 0      'Printed
    Grid_array(SeatNumber, R, 32) = 0      'InclusiveDisc
    Grid_array(SeatNumber, R, 30) = Tb![Discount Name] & ""
    Grid_array(SeatNumber, R, 17) = Tb![Discount Type] & ""
    Grid_array(SeatNumber, R, 2) = Format(OpenCharge, "0.00")
    Grid_array(SeatNumber, R, 18) = Format(OpenCharge, "0.00")
    Grid_array(SeatNumber, R, 25) = Tb.fields("Discount Id").Value & ""
    If EventName = "Automatic Gratuity" Then
     Grid_array(SeatNumber, R, 26) = "Gratuity"  'Department Label
     Grid_array(SeatNumber, R, 23) = "Gratuity"  'Category Label
    Else
     Grid_array(SeatNumber, R, 26) = "Service Charge"  'Department Label
     Grid_array(SeatNumber, R, 23) = "Processing Fee"  'Category Label
    End If
    Grid_array(SeatNumber, R, 19) = -1
    Grid_array(SeatNumber, R, 10) = Val(Tb![Tax 1] & "")
    Grid_array(SeatNumber, R, 11) = Val(Tb![Tax 2] & "")
    Grid_array(SeatNumber, R, 12) = Abs(Val(Tb![Taxable] & ""))
    Grid_array(SeatNumber, R, 5) = "None" '// 5 to 9 are printers
    Grid_array(SeatNumber, R, 6) = "None"
    Grid_array(SeatNumber, R, 7) = "None"
    Grid_array(SeatNumber, R, 8) = "None"
    Grid_array(SeatNumber, R, 9) = "None"
    Grid_array(SeatNumber, R, 3) = 1   'expanded print
    Grid_array(SeatNumber, R, 13) = 0   'expanded print
    NextRow(SeatNumber) = R + 1
    DbClose DiscDb, Tb
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " Get Coupons")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function GetGrossSalesFromDetail(Db As ADODB.Connection, CheckNo As Long, FeeName As String) As Currency
On Error GoTo ErrorHandler
Dim RS             As ADODB.Recordset
If Db Is Nothing Then
  ConnectADODB CheckStatDb, Db
End If
Sql = "Select * from [Check Detail] Where [Close Out Day] = 0 and [Check Number] = " & CheckNo
Sql = Sql & " and [Menu Item] <> '" & FeeName & "' and [Price] > 0"
ConnectADORS Sql, Db, RS
GetGrossSalesFromDetail = 0
OpenCheckVds.VdSales = 0
Do While Not RS.EOF
    GetGrossSalesFromDetail = GetGrossSalesFromDetail + (Abs(Val(RS.fields("Price").Value & "")) - Val(RS.fields("Inclusive amount").Value))
    If RS.fields("Voided").Value = 1 Then OpenCheckVds.VdSales = OpenCheckVds.VdSales + (Abs(Val(RS.fields("Price").Value & "")))
 RS.MoveNext
Loop
DbClose Nothing, RS
Exit Function
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Gross Sales ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub CreditProcessingFee(Db As ADODB.Connection, TenderAmt As Currency)
On Error GoTo ErrorHandler
If ProcessingFeesOn = False Then Exit Sub
Dim RS             As ADODB.Recordset
Dim FeeName        As String
Dim CreditValue    As Currency
Dim CurrentValue   As Currency
Dim Fees           As Integer
  Fees = CountProcessingFeeInUse
  If Fees > 1 And TenderAmt > 0 Then TenderAmt = TenderAmt / Fees
  If Db Is Nothing Then
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  End If
    CurrentValue = GetFeeReduction(TenderAmt)  '  GetCurrentProcessingFee
    If CurrentValue < 0 Then CurrentValue = 0
    If SurChargesUsed > 0 Then
     CreditValue = SurChargesUsed
    Else
     CreditValue = GetProcessingFee(TenderAmt)
    End If
    FeeName = GetProcessingFeeInUse()
    If CurrentValue = 0 Then CreditValue = 0
    If CurrentValue = 0 Then SetProcessingFeeToZero
    Sales_Totals
    Sql = "Select * from [Open Checks] "
    Sql = Sql & " Where [Close Out Day] = 0 and [check number] = " & Check_no  '& " and [server name] = '" & LoggedOnAs(0).EMPName & "'"
   ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
         Gross_Sales = GetGrossSalesFromDetail(Db, Check_no, FeeName) + CurrentValue
         Chk_Sales_data(6) = OpenCheckVds.VdSales
         Chk_Sales_data(7) = Abs(Val(Chk_Sales_data(7))) + OpenCheckVds.VdTx
         Chk_Sales_data(8) = Abs(Val(Chk_Sales_data(8))) + OpenCheckVds.VdSTx
         Chk_Sales_data(9) = Abs(Val(Chk_Sales_data(9))) + OpenCheckVds.VdLTx
         OpenCheckVds.VdSales = 0: OpenCheckVds.VdDisc = 0: OpenCheckVds.VdLTx = 0: OpenCheckVds.VdSTx = 0: OpenCheckVds.VdTx = 0
      Sql = "Update [Open Checks] Set [Gross Special Tax] = " & (Chk_Sales_data(2) + Chk_Sales_data(20)) & ", [Gross Sales Tax] = " & Abs(Chk_Sales_data(1) + Chk_Sales_data(12)) & ", [Gross Sales] = " & Gross_Sales & " ,[Gross Exempt] =  " & (Chk_Sales_data(13)) & ",[Check Total] = " & Chk_Sales_data(5) 'Chk_Sales_data(16) +
      Sql = Sql & " Where [check number] = " & Check_no & " and [server name] = '" & LoggedOnAs(0).EMPName & "'"
    Else
      Sql = "Update [Closed Checks] Set [Gross Special Tax] = " & (Chk_Sales_data(2) + Chk_Sales_data(20)) & ", [Gross Sales Tax] = " & Abs(Chk_Sales_data(1) + Chk_Sales_data(12)) & ", [Check Total] = " & Chk_Sales_data(5) & ", [Gross Sales] = " & Gross_Sales & " ,[Gross Exempt] =  " & (Chk_Sales_data(13))
      If OrderIs = "Delivery" Then
      
      End If
      Sql = Sql & " Where [close out day] = 0 and [check number] = " & Check_no '& " and [server name] = '" & LoggedOnAs(0).EMPName & "'"
    End If
    If Gross_Sales > 0 Then Db.Execute Sql
    RS.Close
    If Fees > 1 And CreditValue > 0 Then CreditValue = CreditValue / Fees
    If TenderAmt < Chk_Sales_data(5) And TenderAmt > 0 Then
      Sql = "Update [Check Detail] Set [Price] = " & CreditValue
    Else
      Sql = "Update [Check Detail] Set [Price] = " & SurChargesUsed   '   & ",[Menu Item] = [Menu Item] + ' Discount'"
    End If
    Sql = Sql & " Where [close out day] = 0 and [check number] = " & Check_no  '& " and [servers name] = '" & LoggedOnAs(0).EMPName & "'"
    Sql = Sql & " and [Menu Item] = '" & FeeName & "'"                         'deliver order will not set amount to 0
    Db.Execute Sql
    Trans_Log "Credit Processing Fee  " & FeeName & " Check Number " & Check_no & " Credit Value " & CreditValue
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Remove Convenience Fee ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function CountProcessingFeeInUse() As Integer
CountProcessingFeeInUse = 0
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
'If Check_Option(173) = 1 Then
' Store_Grid KioskOrder.Ticket(0)
'Else
 Store_Grid Order.Ticket(0)
'End If
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
        CountProcessingFeeInUse = CountProcessingFeeInUse + 1
     End If
 Next Rw
Next SrcFor
'Order.Load_Grid
End Function
Public Function SetCurrentProcessingFee() As Currency
SetCurrentProcessingFee = 0
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
'If Check_Option(173) = 1 Then
' Store_Grid KioskOrder.Ticket(0)
'Else
 Store_Grid Order.Ticket(0)
'End If
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
       Grid_array(SrcFor, Rw, 2) = SurChargesUsed
     End If
 Next Rw
Next SrcFor
Order.Load_Grid
Sales_Totals
End Function
Public Function GetCurrentProcessingFee() As Currency
GetCurrentProcessingFee = 0
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
'If Check_Option(173) = 1 Then
' Store_Grid KioskOrder.Ticket(0)
'Else
 Store_Grid Order.Ticket(0)
'End If
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
       GetCurrentProcessingFee = GetCurrentProcessingFee + Val(Grid_array(SrcFor, Rw, 2))
     End If
 Next Rw
Next SrcFor

Order.Load_Grid
End Function
Public Function GetCashCheck(Net As Boolean) As Currency
Dim NIP As Integer
Dim NOS As Integer
If NextRow(NIP) = 0 Then Exit Function
Dim OLDGrid() As Variant
OLDGrid() = Grid_array()
For NIP = 0 To Number_In_Party - 1
   For NOS = 0 To NextRow(NIP) - 1
       If Grid_array(NIP, NOS, 23) = "Processing Fee" Then
        Grid_array(NIP, NOS, 2) = 0
       Else
        Grid_array(NIP, NOS, 2) = GetCashPrice(Val(Grid_array(NIP, NOS, 0)), CStr(Grid_array(NIP, NOS, 1)), Val(Grid_array(NIP, NOS, 2)), CStr(Grid_array(NIP, NOS, 26)))
       End If
   Next NOS
Next NIP
CashCalculation = 1
Sales_Totals
If Net Then
 GetCashCheck = Chk_Sales_data(0)
Else
 GetCashCheck = Chk_Sales_data(5)
End If
Grid_array() = OLDGrid()
CashCalculation = 0
'// need to separtate cash calculation from non cash for sales tax on seat totals for split checks.;
Sales_Totals '// must stay to reset total to higher value
End Function
Public Function ReturnTaxFlags() As CouponTaxes
Dim NIP As Integer
Dim NOS As Integer
For NIP = 0 To Number_In_Party - 1
   For NOS = 0 To NextRow(NIP) - 1
        If Val(Grid_array(NIP, NOS, 10)) > 0 Then ReturnTaxFlags.Tax1 = Grid_array(NIP, NOS, 10)
        If Val(Grid_array(NIP, NOS, 11)) > 0 Then ReturnTaxFlags.Tax2 = Grid_array(NIP, NOS, 11)
        If Val(Grid_array(NIP, NOS, 12)) > 0 Then ReturnTaxFlags.Tax3 = Grid_array(NIP, NOS, 12)
   Next NOS
Next NIP
End Function
Public Function GetCashAmount(Net As Boolean) As SalesTaxAmt

Sales_Totals
If Chk_Sales_data(5) = 0 Then Exit Function
Dim CreditFee       As Currency
Dim CashFee         As Currency
If DualPricing = 0 Then
 CashFee = GetCashCheck(Net)
 GetCashAmount.CashValue = CashFee
 GetCashAmount.RegularSalesTax = Chk_Sales_data(21)
Else
 CashFee = GetCashCheck(Net)
 GetCashAmount.CashValue = Chk_Sales_data(5) - CashFee
 GetCashAmount.ChkValue = CashFee
 GetCashAmount.RegularSalesTax = Chk_Sales_data(21)
 Sales_Totals
End If
End Function
Public Function GetProcessingCashPrice() As Currency
GetProcessingCashPrice = GetCashCheck(False)
End Function
Public Function GetFeeReduction(TenderAmount As Currency) As Currency
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
GetFeeReduction = 0
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
   If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
      PerCentBack = GetCheckGrat(ProcessItemName(CStr(Grid_array(SrcFor, Rw, 1)), True), MaxGrat)
       Grid_array(SrcFor, Rw, 2) = Format((((Seat_Sales_Data(SrcFor, 5) - TenderAmount) - (Seat_Sales_Data(SrcFor, 18) - Seat_Sales_Data(SrcFor, 16))) * (PerCentBack / 100) * 100) / 100, "#.00")
       If Val(Grid_array(SrcFor, Rw, 2)) < 0 Then Grid_array(SrcFor, Rw, 2) = 0
      GetFeeReduction = GetFeeReduction + Val(Grid_array(SrcFor, Rw, 2))
   End If    'Subtotal = 0 service Charge =14
 Next Rw
Next SrcFor
End Function
Public Function GetProcessingFee(SplitTender As Currency) As Currency
On Error GoTo ErrorHandler
GetProcessingFee = 0
If CashDiscountData.CreditAmount > 0 Then
   Dim Txamt As SalesTaxAmt
   Txamt = GetCashAmount(False)
   If DualPricing = 0 Then
     GetProcessingFee = Txamt.CashValue
   Else
     GetProcessingFee = Chk_Sales_data(5) - Txamt.CashValue
   End If
   Exit Function
End If
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
'Dim CurrentFee  As Currency
'If Check_Option(173) = 1 Then
' If KioskOrder.Ticket(0).TextMatrix(0, 1) > "" Then
'  Store_Grid KioskOrder.Ticket(0)
' End If
'Else
 If Order.Ticket(0).TextMatrix(0, 1) > "" Then
  Store_Grid Order.Ticket(0)
 End If
'End If
Sales_Totals
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
       PerCentBack = GetCheckGrat(ProcessItemName(CStr(Grid_array(SrcFor, Rw, 1)), True), MaxGrat)
       If PerCentBack > 0 Then
          If Seat_Sales_Data(SrcFor, 5) = SplitTender Then
             Grid_array(SrcFor, Rw, 2) = 0
          ElseIf MaxGrat > 0 Then
             Grid_array(SrcFor, Rw, 2) = MaxGrat  'added for David Rappa - fixed fee
          Else
              Grid_array(SrcFor, Rw, 2) = Format((((Seat_Sales_Data(SrcFor, 5) - SplitTender) - (Seat_Sales_Data(SrcFor, 18) - Seat_Sales_Data(SrcFor, 16))) * (PerCentBack / 100) * 100) / 100, "#.00")
              If Val(Grid_array(SrcFor, Rw, 2)) < 0 Then Grid_array(SrcFor, Rw, 2) = 0
          End If
          GetProcessingFee = GetProcessingFee + Val(Grid_array(SrcFor, Rw, 2))
       End If    'Subtotal = 0 service Charge =14
     End If
 Next Rw
Next SrcFor
'If Check_Option(173) = 1 Then
' KioskOrder.Load_Grid
'Else
 Order.Load_Grid
'End If
Sales_Totals
Exit Function  '' error 9 load from seat_Sales _data
ErrorHandler:
Error_Log Err.Number, Err.Description, "Get processing Fee", "None"
End Function
Public Function GetFeeForSplitPayment(SplitAmount As Currency) As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
Dim MaxGrat     As Currency

For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
       PerCentBack = GetCheckGrat(ProcessItemName(CStr(Grid_array(SrcFor, Rw, 1)), True), MaxGrat)
       If PerCentBack > 0 Then
           GetFeeForSplitPayment = Format((SplitAmount * (PerCentBack / 100) * 100) / 100, "0.00")
           Grid_array(SrcFor, Rw, 2) = Format(Val(Grid_array(SrcFor, Rw, 2)), "0.00") + GetFeeForSplitPayment
       End If    'Subtotal = 0 service Charge =14
     End If
 Next Rw
Next SrcFor
Sales_Totals
End Function
Public Function CalculateProcessingFee(SplitTender As Currency) As Currency
On Error GoTo ErrorHandler
CalculateProcessingFee = 0
If CashDiscountData.CreditAmount > 0 Then
   Dim Txamt As SalesTaxAmt
   Txamt = GetCashAmount(False)
   CalculateProcessingFee = CashTxCrd.CashTotal
   Exit Function
End If
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
Dim SrcFor      As Integer
Dim Rw          As Integer
'If Check_Option(173) = 1 Then
' If KioskOrder.Ticket(0).TextMatrix(0, 1) > "" Then
'  Store_Grid KioskOrder.Ticket(0)
' End If
'Else
 If Order.Ticket(0).TextMatrix(0, 1) > "" Then
  Store_Grid Order.Ticket(0)
 End If
'End If
Sales_Totals
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
       PerCentBack = GetCheckGrat(ProcessItemName(CStr(Grid_array(SrcFor, Rw, 1)), True), MaxGrat)
       If PerCentBack > 0 Then
          If Seat_Sales_Data(SrcFor, 5) = SplitTender Then
             CalculateProcessingFee = 0
          ElseIf MaxGrat > 0 Then
             CalculateProcessingFee = MaxGrat  'added for David Rappa - fixed fee
          Else
             CalculateProcessingFee = Format(((((Seat_Sales_Data(SrcFor, 5) - SplitTender) - (Seat_Sales_Data(SrcFor, 18) - Seat_Sales_Data(SrcFor, 16))) * (PerCentBack / 100) * 100) / 100) + 0.005, "#.00") '- Seat_Sales_Data(SrcFor, 14)
          End If
       End If    'Subtotal = 0 service Charge =14
     End If
 Next Rw
Next SrcFor
Sales_Totals
'If Check_Option(173) = 1 Then
' KioskOrder.Load_Grid
'Else
 Order.Load_Grid
'End If
Exit Function  '' error 9 load from seat_Sales _data
ErrorHandler:
Error_Log Err.Number, Err.Description, "Get processing Fee", "None"
End Function
Public Function GetPartialFee(PaymentAmt As Currency) As Currency
Dim SrcFor As Integer
Dim Rw     As Integer
GetPartialFee = 0
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
      GetPartialFee = Seat_Sales_Data(SrcFor, 5) - PaymentAmt
     End If
 Next Rw
Next SrcFor
End Function
Public Function CalculateSeatGratuity(SeatNo As Integer, GratutityName As String) As Currency
CalculateSeatGratuity = 0
Dim PerCentBack As Single
Sales_Totals
PerCentBack = GetCheckGrat(ProcessItemName(CStr(GratutityName), True), 0)
If PerCentBack > 0 Then
 Dim Calculation As Currency
  Calculation = Seat_Sales_Data(SeatNo, 14) + Seat_Sales_Data(SeatNo, 20)
   'Subtotal = 0 service Charge =14 gratuity = 20
  If Input_Option(56) = "Net Sales" Then
   CalculateSeatGratuity = ((Seat_Sales_Data(SeatNo, 0) - Calculation) + Abs(Seat_Sales_Data(SeatNo, 4))) * (PerCentBack / 100)
  Else
   CalculateSeatGratuity = (Abs(Seat_Sales_Data(SeatNo, 13))) * (PerCentBack / 100)
  End If
End If
End Function
Public Function GetSeatProcessingFee(CSeat As Integer, FeeName As String) As Currency
GetSeatProcessingFee = 0
If CashDiscountData.CreditAmount > 0 Then
   Dim Txamt As SalesTaxAmt
   Txamt = GetCashAmount(False)
   GetSeatProcessingFee = CashTxCrd.CashTotal
   Exit Function
End If
If ProcessingFeesOn = False Then Exit Function
Dim MaxGrat     As Currency
Dim PerCentBack As Single
If FeeName = "" Then Exit Function
PerCentBack = GetCheckGrat(ProcessItemName(CStr(FeeName), True), MaxGrat)
If PerCentBack > 0 Then
   GetSeatProcessingFee = Format((((Seat_Sales_Data(CSeat, 5)) - (Seat_Sales_Data(CSeat, 18) - Seat_Sales_Data(CSeat, 16))) * (PerCentBack / 100) * 100) / 100, "#.00") '- Seat_Sales_Data(CSeat, 14)
End If    'Subtotal = 0 service Charge =14
End Function
Public Sub SetProcessingFeeToZero()
If ProcessingFeesOn = False Then Exit Sub
'If Order.Ticket(0).TextMatrix(0, 1) > "" Then Exit Sub
If Order.Ticket(0).TextMatrix(0, 1) = "" Then Exit Sub
 
 'If Check_Option(173) = 1 Then
 ' Store_Grid KioskOrder.Ticket(0)
 'Else
  Store_Grid Order.Ticket(0)
 'End If
'If IsProcessFeeZero = 0 Then Exit Sub
Dim SrcFor      As Integer
Dim Rw          As Integer
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim FeeName     As String

  For SrcFor = 0 To Number_In_Party - 1
   For Rw = 0 To NextRow(SrcFor) - 1
     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
      If Not AllExemptSales.Item("Service Charge") Is Nothing Then
       AllExemptSales.Item("Service Charge").ExemptSales = AllExemptSales.Item("Service Charge").ExemptSales - Val(Grid_array(SrcFor, Rw, 2))
       If AllExemptSales.Item("Service Charge").ExemptSales < 0 Then AllExemptSales.Item("Service Charge").ExemptSales = 0
      End If
      Grid_array(SrcFor, Rw, 2) = Format("#.00")
     End If
   Next Rw
  Next SrcFor
  Order.Load_Grid
  Sales_Totals
  
  FeeName = GetProcessingFeeInUse()
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select * from [Open Checks] "
    Sql = Sql & " Where [Close Out Day] = 0 and [check number] = " & Check_no '& " and [server name] = '" & LoggedOnAs(0).EMPName & "'"
   ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
       Gross_Sales = GetGrossSalesFromDetail(Db, Check_no, FeeName) + 0
       Chk_Sales_data(6) = OpenCheckVds.VdSales
       Chk_Sales_data(7) = Abs(Val(Chk_Sales_data(7))) + OpenCheckVds.VdTx
       Chk_Sales_data(8) = Abs(Val(Chk_Sales_data(8))) + OpenCheckVds.VdSTx
       Chk_Sales_data(9) = Abs(Val(Chk_Sales_data(9))) + OpenCheckVds.VdLTx
       OpenCheckVds.VdSales = 0: OpenCheckVds.VdDisc = 0: OpenCheckVds.VdLTx = 0: OpenCheckVds.VdSTx = 0: OpenCheckVds.VdTx = 0                                '// Chk_Sales_data(12) this is the tax on inclusive items
      
      Sql = "Update [Open Checks] Set [Gross Special Tax] = " & (Chk_Sales_data(2) + Chk_Sales_data(20)) & ", [Gross Sales Tax] = " & Abs(Chk_Sales_data(1) + Chk_Sales_data(12)) & ", [Gross Sales] = " & Gross_Sales & " ,[Gross Exempt] =  " & (Chk_Sales_data(13)) & ",[Check Total] = " & Chk_Sales_data(5) 'Chk_Sales_data(16) +
      Sql = Sql & " Where [close out day] = 0 and [check number] = " & Check_no
    Else
      Gross_Sales = GetGrossSalesFromDetail(Db, Check_no, FeeName) + 0
      Chk_Sales_data(6) = OpenCheckVds.VdSales
      Chk_Sales_data(7) = Abs(Val(Chk_Sales_data(7))) + OpenCheckVds.VdTx
      Chk_Sales_data(8) = Abs(Val(Chk_Sales_data(8))) + OpenCheckVds.VdSTx
      Chk_Sales_data(9) = Abs(Val(Chk_Sales_data(9))) + OpenCheckVds.VdLTx
      OpenCheckVds.VdSales = 0: OpenCheckVds.VdDisc = 0: OpenCheckVds.VdLTx = 0: OpenCheckVds.VdSTx = 0: OpenCheckVds.VdTx = 0
      Sql = "Update [Closed Checks] Set [Gross Special Tax] = " & (Chk_Sales_data(2) + Chk_Sales_data(20)) & ", [Gross Sales Tax] = " & Abs(Chk_Sales_data(1) + Chk_Sales_data(12)) & ", [Check Total] = " & Chk_Sales_data(5) & ", [Gross Sales] = " & Gross_Sales & " ,[Gross Exempt] =  " & (Chk_Sales_data(13))
      Sql = Sql & " Where [close out day] = 0 and [check number] = " & Check_no '& " and [server name] = '" & LoggedOnAs(0).EMPName & "'"
    End If
    Db.Execute Sql
    RS.Close
    Sql = "Update [Check Detail] Set [Price] = " & 0   '   & ",[Menu Item] = [Menu Item] + ' Discount'"
    Sql = Sql & " Where [close out day] = 0 and [check number] = " & Check_no  '& " and [servers name] = '" & LoggedOnAs(0).EMPName & "'"
    Sql = Sql & " and [Menu Item] = '" & FeeName & "'"                         'deliver order will not set amount to 0
    Db.Execute Sql
    DbClose Db, Nothing
    
    Trans_Log "Set Processing Fee To 0 " & FeeName & " Check Number " & Check_no & " SurCharge value " & SurChargesUsed
    SurChargesUsed = 0
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Remove Convenience Fee ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
'Public Function PutFeeBackOnCheck(AddedFee As Currency) As Boolean
'If Order.Ticket(0).TextMatrix(0, 1) = "" Then Exit Sub
'If Check_Option(173) = 1 Then
'  Store_Grid KioskOrder.Ticket(0)
'Else
'  Store_Grid Order.Ticket(0)
'End If
'Dim SrcFor As Integer
'Dim Rw     As Integer
'For SrcFor = 0 To Number_In_Party - 1
' For Rw = 0 To NextRow(SrcFor) - 1
'     If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
'        Grid_array(SrcFor, Rw, 2) = Format(AddedFee, "#.00")
'        Exit For
'     End If
' Next Rw
'Next SrcFor
'Order.Load_Grid
'End Function

Public Function CheckProceesingFee(SeatNo As Integer) As Boolean
If ProcessingFeesOn = False Then Exit Function
 CheckProceesingFee = False
 Dim SC   As Integer
 For SC = 0 To NextRow(SeatNo) - 1
 If Grid_array(SeatNo, SC, 23) = "Processing Fee" Then
    CheckProceesingFee = True
    Exit Function
 End If
 If Grid_array(SeatNo, SC, 1) = "" Then Exit For
 Next SC
End Function
Public Function GetProcessingFeeInUse() As String
GetProcessingFeeInUse = ""
If ProcessingFeesOn = False Then Exit Function
Dim SrcFor      As Integer
Dim Rw          As Integer
'If Check_Option(173) = 1 Then
' Store_Grid KioskOrder.Ticket(0)
'Else
 Store_Grid Order.Ticket(0)
'End If
For SrcFor = 0 To Number_In_Party - 1
 For Rw = 0 To NextRow(SrcFor) - 1
   If Grid_array(SrcFor, Rw, 23) = "Processing Fee" Then
     GetProcessingFeeInUse = ProcessItemName(CStr(Grid_array(SrcFor, Rw, 1)), True)
   End If
 Next Rw
Next SrcFor
End Function
Public Function Consolidate_Qty(PrintType As String)
On Error Resume Next
Dim CNip        As Integer
Dim CSn         As Integer
Dim CBal        As Integer
Dim CutAst      As String
Dim CPName1     As String
Dim CPName2     As String
Dim CPrice1     As Currency
Dim CPrice2     As Currency
Dim SS          As Integer   'Search Seat
  For CNip = 0 To Number_In_Party - 1
    For CSn = 0 To NextRow(CNip) - 1
      If Left(Trim(Grid_array(CNip, CSn, 1)), 1) = "v" Or Val(Grid_array(CNip, CSn, 13)) = "1" Then GoTo NextItem
      If Val(Grid_array(CNip, CSn, 2)) < 0 Then GoTo NextItem
      If Val(Grid_array(CNip, CSn, 0)) = 0 Then GoTo NextItem
      If Val(Grid_array(CNip, CSn, 0)) > 0 And Val(Grid_array(CNip, CSn + 1, 0)) = 0 And Trim(Grid_array(CNip, CSn + 1, 1)) > "" Then GoTo NextItem
      If Trim(Grid_array(CNip, CSn, 1)) = "<>" Then GoTo NextItem
      If PrintType = "Remote" Then
        If Left(Trim(Grid_array(CNip, CSn, 1)), 1) = "*" Then GoTo NextItem
      End If
      'If PrintType <> "Remote" And NoRoll = 1 Then GoTo NextItem
      For SS = CSn + 1 To NextRow(CNip) - 1
      If Trim$(Grid_array(CNip, SS, 1)) = "<>" Or Val(Grid_array(CNip, SS, 13)) = "1" Then GoTo NextLine
      If Val(Grid_array(CNip, SS, 0)) = 0 And Grid_array(CNip, SS, 1) > "" Then GoTo NextLine
      If Val(Grid_array(CNip, CSn, 0)) > 0 And Val(Grid_array(CNip, SS, 0)) > 0 Then
        If Val(Grid_array(CNip, SS + 1, 0)) = 0 And Trim(Grid_array(CNip, SS + 1, 1)) > "" Then GoTo NextLine
          If PrintType = "Validate" Then
            CPName1 = Trim(Grid_array(CNip, CSn, 1))
            CPName2 = Trim(Grid_array(CNip, SS, 1))
            CPrice1 = Val(Grid_array(CNip, CSn, 2))
            CPrice2 = Val(Grid_array(CNip, SS, 2))
          Else
            CPName1 = ProcessItemName(CStr(Grid_array(CNip, CSn, 1)), True)
            CPName2 = ProcessItemName(CStr(Grid_array(CNip, SS, 1)), True)
            CPrice1 = Val(Grid_array(CNip, CSn, 2))
            CPrice2 = Val(Grid_array(CNip, SS, 2))
          End If
          CutAst = Left(Trim(Grid_array(CNip, SS, 1)), 1)
             If CutAst <> "v" Then
                If Trim$(CPName1) = Trim$(CPName2) Then  ' And CPrice1 = CPrice2 Then    '// price was used  to keep happy hour priced item separate.
                  Grid_array(CNip, CSn, 0) = Val(Grid_array(CNip, CSn, 0)) + Val(Grid_array(CNip, SS, 0))
                  Grid_array(CNip, CSn, 2) = Val(Grid_array(CNip, CSn, 2)) + Val(Grid_array(CNip, SS, 2))
                  Grid_array(CNip, SS, 1) = "<>": Grid_array(CNip, SS, 2) = 0: Grid_array(CNip, SS, 12) = 0
                End If  '//
             End If
        End If
NextLine:
       Next SS
NextItem:
      Next CSn
   Next CNip
 End Function
Public Function ConsolidateVirtualCheck()
On Error Resume Next
Dim RemoteGrid()        As Variant
Dim RemoteNextSeat()    As Integer
Dim CNip                As Integer
Dim CSn                 As Integer
Dim CBal                As Integer
Dim CutAst              As String
Dim CPName1             As String
Dim CPName2             As String
Dim SS                  As Integer   'Search Seat
Dim C                   As Integer
Dim ShiftLookUp         As String
ShiftLookUp = GetConfigInfo("General", "Quanity Consolidate Use Price Shift Look Up", "No", "Functions.INI")
RemoteGrid() = Grid_array()
RemoteNextSeat() = NextRow()
ReDim Grid_array(Number_In_Party, MaxRows, TTL_Col) As Variant
   For CNip = 0 To Number_In_Party - 1
    For CSn = 0 To RemoteNextSeat(CNip) - 1
      If Left(Trim(RemoteGrid(CNip, CSn, 1)), 1) = "v" Or Val(RemoteGrid(CNip, CSn, 13)) = "1" Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 2)) < 0 Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 0)) = 0 Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 0)) > 0 And Val(RemoteGrid(CNip, CSn + 1, 0)) = 0 And Trim(RemoteGrid(CNip, CSn + 1, 1)) > "" Then GoTo NextItem
      If Trim(RemoteGrid(CNip, CSn, 1)) = "<>" Then GoTo NextItem
      For SS = CSn + 1 To RemoteNextSeat(CNip) - 1
      If Trim(RemoteGrid(CNip, SS, 1)) = "<>" Or Val(RemoteGrid(CNip, SS, 13)) = "1" Then GoTo NextLine
      If Val(RemoteGrid(CNip, SS, 0)) = 0 And RemoteGrid(CNip, SS, 1) > "" Then GoTo NextLine
      If Val(RemoteGrid(CNip, CSn, 0)) > 0 And Val(RemoteGrid(CNip, SS, 0)) > 0 Then
        If Val(RemoteGrid(CNip, SS + 1, 0)) = 0 And Trim(RemoteGrid(CNip, SS + 1, 1)) > "" Then GoTo NextLine
          CPName1 = Trim(RemoteGrid(CNip, CSn, 1))
          CPName2 = Trim(RemoteGrid(CNip, SS, 1))
             CutAst = Left(Trim(RemoteGrid(CNip, SS, 1)), 1)
             If CutAst <> "v" Then
             If Trim(CPName1) = Trim(CPName2) Then
               If CutAst = "*" Then
                If RemoteGrid(CNip, CSn, 31) = "None" Then RemoteGrid(CNip, CSn, 31) = ""
                RemoteGrid(CNip, CSn, 31) = CStr(RemoteGrid(CNip, CSn, 31)) & CStr(RemoteGrid(CNip, SS, 14))
               End If
               RemoteGrid(CNip, CSn, 0) = Val(RemoteGrid(CNip, CSn, 0)) + Val(RemoteGrid(CNip, SS, 0))
               RemoteGrid(CNip, CSn, 2) = Format(Val(RemoteGrid(CNip, CSn, 2)) + Val(RemoteGrid(CNip, SS, 2)), "0.00")
               RemoteGrid(CNip, SS, 1) = "<>": RemoteGrid(CNip, SS, 2) = 0: RemoteGrid(CNip, SS, 12) = 0
              End If
             End If
        End If
NextLine:
       Next SS
NextItem:
      Next CSn
   Next CNip
   C = -1
   For CNip = 0 To Number_In_Party - 1
    NextRow(CNip) = 0
    C = -1
    For CSn = 0 To RemoteNextSeat(CNip) - 1
      If RemoteGrid(CNip, CSn, 1) <> "<>" Then
         C = C + 1
         NextRow(CNip) = NextRow(CNip) + 1
         For SS = 0 To TTL_Col
          Grid_array(CNip, C, SS) = RemoteGrid(CNip, CSn, SS)
         Next SS
         If Val(Grid_array(CNip, C, 0)) > 1 And ShiftLookUp = "ENABLED" Then
           Grid_array(CNip, C, 2) = Format(GetShiftConslidated(CStr(Grid_array(CNip, C, 1)), Val(Grid_array(CNip, C, 2)), Val(Grid_array(CNip, C, 0))), "0.00")
         End If
      End If
     Next CSn
  Next CNip
End Function
Public Function GetShiftConslidated(ItemName As String, PriceIn As Currency, QuantityIn As Integer) As Currency
GetShiftConslidated = PriceIn
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim ShiftQty As Integer
ShiftQty = QuantityIn
Dim factor   As Single
'If ShiftQty > 5 Then ShiftQty = 5
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
 Sql = "Select * from [Menu Items] where [Menu Item Name] = '" & ItemName & "'"
ConnectADORS Sql, Db, RS
  If RS.EOF Then
     DbClose Db, RS
     Exit Function
  Else
     If QuantityIn > 1 Then
       If Val(RS.fields("Price 2").Value & "") > 0 Then
          factor = QuantityIn / 2
          GetShiftConslidated = SetDualPrice(RS.fields("Price 2").Value) * (Int(factor))
           If Int((factor - Int(factor)) * 100) > 0 Then
            GetShiftConslidated = GetShiftConslidated + SetDualPrice(RS.fields("Regular Price").Value)
           End If
       ElseIf Val(RS.fields("Price 3").Value & "") > 0 And QuantityIn > 2 Then
          factor = QuantityIn / 3
          GetShiftConslidated = SetDualPrice(RS.fields("Price 3").Value) * Int(factor)
           If Int((factor - Int(factor)) * 100) = 33 Then
            GetShiftConslidated = GetShiftConslidated + SetDualPrice(RS.fields("Regular Price").Value)
           ElseIf Int((factor - Int(factor)) * 100) = 66 Then
            GetShiftConslidated = GetShiftConslidated + SetDualPrice(RS.fields("Regular Price").Value) * 2
           End If
       ElseIf Val(RS.fields("Price 4").Value & "") > 0 And QuantityIn > 3 Then
          factor = QuantityIn / 4
          GetShiftConslidated = SetDualPrice(RS.fields("Price 4").Value) * Int(factor)
           If Int((factor - Int(factor)) * 100) = 25 Then
            GetShiftConslidated = GetShiftConslidated + SetDualPrice(RS.fields("Regular Price").Value)
           ElseIf Int((factor - Int(factor)) * 100) = 50 Then
            GetShiftConslidated = GetShiftConslidated + SetDualPrice(RS.fields("Regular Price").Value) * 2
           ElseIf Int((factor - Int(factor)) * 100) = 75 Then
            GetShiftConslidated = GetShiftConslidated + SetDualPrice(RS.fields("Regular Price").Value) * 3
           End If
       End If
     End If
  End If
  DbClose Db, RS
Exit Function
ERH:
DbClose Db, RS
GetShiftConslidated = PriceIn
End Function
Public Sub ConsolidateODDCheck(RemoteVideoGrid() As Variant, RemoteRowNo() As Integer)
On Error GoTo ERH
Dim CNip                As Integer
Dim CSn                 As Integer
Dim CBal                As Integer
Dim CutAst              As String
Dim CPName1             As String
Dim CPName2             As String
Dim SS                  As Integer   'Search Seat
Dim C                   As Integer
Dim RemoteGrid()        As Variant
Dim RemoteNextSeat()    As Integer
RemoteGrid() = Grid_array()
RemoteNextSeat() = NextRow()
ReDim RemoteVideoGrid(Number_In_Party, MaxRows, TTL_Col) As Variant
ReDim RemoteRowNo(Number_In_Party)
   For CNip = 0 To Number_In_Party - 1
    For CSn = 0 To RemoteNextSeat(CNip) - 1
      If Left(Trim(RemoteGrid(CNip, CSn, 1)), 1) = "v" Or Val(RemoteGrid(CNip, CSn, 13)) = "1" Then GoTo NextItem
      'If Check_Option(239) = 1 Then
      '  If RemoteGrid(CNip, CSn, 5) <> "Monitor 1" Then RemoteGrid(CNip, CSn, 1) = "<>" 'GoTo NextItem
      'End If
      If Val(RemoteGrid(CNip, CSn, 2)) < 0 Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 0)) = 0 Then GoTo NextItem
      
      If Val(RemoteGrid(CNip, CSn, 0)) > 0 And Val(RemoteGrid(CNip, CSn + 1, 0)) = 0 And Trim(RemoteGrid(CNip, CSn + 1, 1)) > "" Then GoTo NextItem
      If Trim(RemoteGrid(CNip, CSn, 1)) = "<>" Then GoTo NextItem
      For SS = CSn + 1 To RemoteNextSeat(CNip) - 1
      If Trim(RemoteGrid(CNip, SS, 1)) = "<>" Or Val(RemoteGrid(CNip, SS, 13)) = "1" Then GoTo NextLine
      If Val(RemoteGrid(CNip, SS, 0)) = 0 And RemoteGrid(CNip, SS, 1) > "" Then GoTo NextLine
      If Val(RemoteGrid(CNip, CSn, 0)) > 0 And Val(RemoteGrid(CNip, SS, 0)) > 0 Then
        If Val(RemoteGrid(CNip, SS + 1, 0)) = 0 And Trim(RemoteGrid(CNip, SS + 1, 1)) > "" Then GoTo NextLine
          CPName1 = Trim(RemoteGrid(CNip, CSn, 1))
          CPName2 = Trim(RemoteGrid(CNip, SS, 1))
             CutAst = Left(Trim(RemoteGrid(CNip, SS, 1)), 1)
             If CutAst <> "v" Then
             If Trim(CPName1) = Trim(CPName2) Then
               If CutAst = "*" Then
                If RemoteGrid(CNip, CSn, 31) = "None" Then RemoteGrid(CNip, CSn, 31) = ""
                RemoteGrid(CNip, CSn, 31) = CStr(RemoteGrid(CNip, CSn, 31)) & CStr(RemoteGrid(CNip, SS, 14))
               End If
               RemoteGrid(CNip, CSn, 0) = Val(RemoteGrid(CNip, CSn, 0)) + Val(RemoteGrid(CNip, SS, 0))
               RemoteGrid(CNip, CSn, 2) = Format(Val(RemoteGrid(CNip, CSn, 2)) + Val(RemoteGrid(CNip, SS, 2)), "0.00")
               RemoteGrid(CNip, SS, 1) = "<>": RemoteGrid(CNip, SS, 2) = 0: RemoteGrid(CNip, SS, 12) = 0
              End If
             End If
        End If
NextLine:
       Next SS
NextItem:
      Next CSn
   Next CNip
   C = -1
   RemoteRowNo(0) = 0
   'C = -1
   For CNip = 0 To Number_In_Party - 1
    'RemoteRowNo(CNip) = 0
    'C = -1
    For CSn = 0 To RemoteNextSeat(CNip) - 1
      If RemoteGrid(CNip, CSn, 1) <> "<>" And RemoteGrid(CNip, CSn, 1) > "" Then
         C = C + 1
         'RemoteRowNo(CNip) = RemoteRowNo(CNip) + 1
         RemoteRowNo(0) = RemoteRowNo(0) + 1
         
         For SS = 0 To TTL_Col
          RemoteVideoGrid(0, C, SS) = RemoteGrid(CNip, CSn, SS)
         Next SS
      End If
      'If C > -1 Then NextRow(CNip) = C
     Next CSn
  Next CNip
Erase RemoteGrid
Erase RemoteNextSeat
Exit Sub
ERH:
Err.Clear
End Sub
Public Function OKToAdd(InItem As String) As Boolean
On Error GoTo ERH
    Select Case Left(InItem, 1)
      Case Is = "v", "*", "<", ""
        OKToAdd = False
      Case Else
        OKToAdd = True
    End Select
  Exit Function
ERH:
  OKToAdd = True
End Function
Public Function ItemInBag(PrintMsg As String) As String
ItemInBag = ""
Dim NIP           As Integer
Dim SN            As Integer
Dim MainQty       As Long
MainQty = 0
For NIP = 0 To Number_In_Party - 1
 For SN = 0 To NextRow(NIP) - 1
    If Val(Grid_array(NIP, SN, 0) & "") > 0 Then
      MainQty = MainQty + Val(Grid_array(NIP, SN, 0) & "")
    End If
 Next SN
Next NIP
ItemInBag = PrintMsg & " " & CStr(MainQty)
End Function
Public Function MultiSeatConsolidate(TempArray() As Variant, TempNextRow() As Integer, SortNip As Integer)
On Error GoTo ErrorHandler
'**************************
Dim Rw            As Integer
Dim T             As Integer
Dim LS            As Integer
Dim SM            As Integer
Dim Cr            As Integer
Dim LTR           As Integer
Dim MainItem      As String
Dim ItemCourse    As String
Dim ActiveCourse  As String
Dim MainQty       As Long
Dim MasterArray() As Variant
ReDim MainArray(0, MaxRows, TTL_Col)
Dim AllMods As SortAll
Dim AllSort As SortAll
'Dim SortItem As SortEach
Dim AHolder As SortAll
Dim MHolder As SortAll
'Set SortItem = New SortEach
Set AllSort = New SortAll
Set AllMods = New SortAll
Set AHolder = New SortAll
Set MHolder = New SortAll
Dim LTS  As Integer
Dim SN   As Integer
'Dim LTR  As Integer
Dim NIP As Integer
For NIP = 0 To Number_In_Party - 1
 For SN = 0 To TempNextRow(NIP) - 1
    TempArray(NIP, SN, 28) = TempArray(NIP, SN, 1)
 Next SN
Next NIP
Dim SR As Integer
Dim SS  As Integer
'Use Alt Print Field for price Shift Items will not print to remote if.
'06.14.2007 Frank New York
For NIP = 0 To Number_In_Party - 1
 For SN = 0 To TempNextRow(NIP) - 1
    If TempArray(NIP, SN, 1) = "" Then Exit For
    If Left(TempArray(NIP, SN, 1), 2) <> "h-" Then
    If Val(TempArray(NIP, SN, 0)) > 0 Then
       If OKToAdd(CStr(TempArray(NIP, SN, 1))) Then
          ItemCourse = TempArray(NIP, SN, 34)
            If AllSort.Item(CStr(TempArray(NIP, SN, 1) & ItemCourse)) Is Nothing Then
               AllSort.add CStr(TempArray(NIP, SN, 1) & ItemCourse)
               AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).MenuItem = TempArray(NIP, SN, 1)
               AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).qty = 0
               AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).SaleAmt = 0
               AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).CourseNames = ItemCourse
            End If
            AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).qty = AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).qty + Val(TempArray(NIP, SN, 0))
            AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).SaleAmt = AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).SaleAmt + Val(TempArray(NIP, SN, 2))
            AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).SeatNo = AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).SeatNo & "(" & CStr(NIP + 1) & ")"
            MainQty = Val(TempArray(NIP, SN, 0))
            MainItem = AllSort.Item(CStr(TempArray(NIP, SN, 1)) & ItemCourse).MenuItem
      End If
    End If
    If Val(TempArray(NIP, SN, 0)) = 0 Then
    If OKToAdd(CStr(TempArray(NIP, SN, 1))) Then
     If AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))) Is Nothing Then
       AllMods.add ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))
       AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).ModName = TempArray(NIP, SN, 1)
       AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).MenuItem = MainItem
       AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).qty = 0
       AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).SaleAmt = 0
       AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).CourseNames = ItemCourse
     End If
     AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).qty = AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).qty + MainQty
     AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).SaleAmt = AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).SaleAmt + Val(TempArray(NIP, SN, 2))
     AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).SeatNo = AllMods.Item(ItemCourse & MainItem & CStr(TempArray(NIP, SN, 1))).SeatNo & "(" & CStr(NIP + 1) & ")"
    End If
   End If
  End If
 Next SN
Next NIP
Cr = -1
Dim MSP As Integer
Dim MsS As Integer
NSI:
   For NIP = 0 To Number_In_Party - 1
    For SN = 0 To TempNextRow(NIP) '- 1
       MainItem = TempArray(NIP, SN, 28)
       ItemCourse = TempArray(NIP, SN, 34)
       If TempArray(NIP, SN, 1) = "" Then Exit For
       If TempArray(NIP, SN, 1) <> "<>" Or Left(TempArray(NIP, SN, 1), 2) <> "h-" Then
        If Not AllSort.Item(MainItem & ItemCourse) Is Nothing Then
        If AllSort.Item(MainItem & ItemCourse).MenuItem = TempArray(NIP, SN, 1) And Val(TempArray(NIP, SN, 0)) > 0 And OKToAdd(CStr(TempArray(NIP, SN, 1))) Then
         If AHolder.Item(MainItem & ItemCourse) Is Nothing Then
          AHolder.add MainItem & ItemCourse
          Cr = Cr + 1
          For LS = 0 To TTL_Col
            MainArray(0, Cr, LS) = TempArray(NIP, SN, LS)
          Next LS
          MainItem = TempArray(NIP, SN, 28)
          ItemCourse = TempArray(NIP, SN, 34)
          TempArray(NIP, SN, 28) = "<>"
          TempArray(NIP, SN, 1) = "<>"
          MainArray(0, Cr, 24) = 0
          MainArray(0, Cr, 12) = 0  'surcharge = 0
          MainArray(0, Cr, 0) = AllSort.Item(MainItem & ItemCourse).qty
          MainArray(0, Cr, 2) = AllSort.Item(MainItem & ItemCourse).SaleAmt
          If Number_In_Party > 1 Then
            MainArray(0, Cr, 1) = MainArray(0, Cr, 1) & AllSort.Item(MainItem & ItemCourse).SeatNo
          End If
          If MainArray(0, Cr, 28) > "" Then
            MainArray(0, Cr, 28) = MainArray(0, Cr, 28) & AllSort.Item(MainItem & ItemCourse).SeatNo
          Else
            MainArray(0, Cr, 28) = MainArray(0, Cr, 1) '& AllSort.Item(Rw).SeatNo
          End If
         End If
         For LTS = 1 To AllMods.count
           For MSP = 0 To Number_In_Party - 1
             For MsS = 0 To UBound(TempArray(), 2) 'TempNextRow(NIP) '- 1
              If IsEmpty(TempArray(MSP, MsS, 1)) Then Exit For
              If AllMods.Item(LTS).MenuItem = MainItem And AllMods.Item(LTS).ModName = TempArray(MSP, MsS, 1) And OKToAdd(CStr(TempArray(MSP, MsS, 1))) Then
                 If MHolder.Item(ItemCourse & MainItem & TempArray(MSP, MsS, 1)) Is Nothing Then
                    MHolder.add ItemCourse & MainItem & TempArray(MSP, MsS, 1)
                    Cr = Cr + 1
                    For LS = 0 To TTL_Col
                      MainArray(0, Cr, LS) = TempArray(MSP, MsS, LS)
                    Next LS
                    MainArray(0, Cr, 24) = 0
                    MainArray(0, Cr, 0) = ""
                    MainArray(0, Cr, 2) = AllMods.Item(ItemCourse & MainItem & TempArray(MSP, MsS, 1)).SaleAmt  'sm
                    MainArray(0, Cr, 1) = CStr(AllMods.Item(ItemCourse & MainItem & TempArray(MSP, MsS, 1)).qty) & " - " & MainArray(0, Cr, 1) & AllMods.Item(ItemCourse & MainItem & TempArray(MSP, MsS, 1)).SeatNo
                    MainArray(0, Cr, 28) = MainArray(0, Cr, 1)
                    TempArray(MSP, MsS, 1) = "<>"
                    TempArray(MSP, MsS, 28) = "<>"
                    TempArray(MSP, MsS, 12) = 0
                 End If
              End If
             Next MsS
           Next MSP
         Next LTS
       End If
      End If
     End If
    Next SN
  Next NIP
NRW:
SortNip = Number_In_Party * -1
TempNextRow(0) = Cr + 1
TempArray() = MainArray()
'Set SortItem = Nothing
Set AllSort = Nothing
Set AllMods = Nothing
Set AHolder = Nothing
Set MHolder = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "MultiSeatConsolidate")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function SetGridCourseColor(CourseName As String) As String
Dim CRL As Integer
    For CRL = 0 To UBound(CourseLineName, 1)
      If CourseName = CourseLineName(CRL).PosCourseName Then
        SetGridCourseColor = CourseLineName(CRL).CourseNumber
        Exit For
      End If
    Next CRL
End Function
Public Function ConsolidateSeat(TempArray() As Variant, TempNetRow() As Integer, SortNip As Integer, SeatNo As Integer) As String
 '14 item number
 On Error GoTo ERH
 Dim SortArray()        As Variant
 Dim MainItemPrice      As Currency
 Dim MainItemQty        As Single
 Dim MainItemName       As String
 Dim KeepItemsToGether  As Integer
 Dim SN                 As Integer
 Dim SubI               As Integer
 Dim SubItemsName()     As String
 Dim SubItemPrice()     As Currency
 Dim Counter            As Integer
 Dim MainItems          As SortAll
 Dim ModItem            As SortAll
 Dim SubItems           As String
 Set MainItems = Nothing
 Set ModItem = Nothing
 Set MainItems = New SortAll
 Set ModItem = New SortAll
 Dim CRow               As Integer
 Dim MRow               As Integer
 SortArray() = TempArray()
 For SN = 0 To TempNetRow(SeatNo) - 1
 If IsEmpty(SortArray(SeatNo, SN, 1)) Then
 
 Else
   If Val(SortArray(SeatNo, SN, 0)) > 0 And SortArray(SeatNo, SN, 1) <> "Counted" Then
      MainItemQty = Val(SortArray(SeatNo, SN, 0))
      MainItemName = Trim$(UCase(SortArray(SeatNo, SN, 1)))
      MainItemPrice = SortArray(SeatNo, SN, 2)
     ' KeepItemsToGether = SortArray(SeatNo, SN, 14)
      Erase SubItemsName: Erase SubItemPrice
      SubItems = ""
      Counter = 0
    For SubI = SN + 1 To TempNetRow(SeatNo) - 1
      If Val(SortArray(SeatNo, SubI, 0)) > 0 Then Exit For
      If IsEmpty(SortArray(SeatNo, SubI, 1)) Then GoTo Done
      'If SortArray(SeatNo, SubI, 14) = KeepItemsToGether Then
         Counter = Counter + 1
         ReDim Preserve SubItemsName(Counter)
         ReDim Preserve SubItemPrice(Counter)
         SubItemsName(Counter) = Trim$(UCase(SortArray(SeatNo, SubI, 1)))
         SubItemPrice(Counter) = Val(SortArray(SeatNo, SubI, 2))
         SubItems = SubItems & Trim$(UCase(SortArray(SeatNo, SubI, 1)))
      'End If
    Next SubI
    If MainItems.Item(MainItemName & SubItems) Is Nothing Then
       MainItems.add MainItemName & SubItems
       MainItems.Item(MainItemName & SubItems).MenuItem = MainItemName
       CRow = SN
       SortArray(SeatNo, CRow, 2) = Val(SortArray(SeatNo, CRow, 2))
    Else
       MainItems.Item(MainItemName & SubItems).MenuItem = MainItemName
       SortArray(SeatNo, SN, 1) = "Counted"
       SortArray(SeatNo, CRow, 0) = Val(SortArray(SeatNo, CRow, 0)) + Val(SortArray(SeatNo, SN, 0))
       SortArray(SeatNo, CRow, 2) = Val(SortArray(SeatNo, CRow, 2)) + Val(SortArray(SeatNo, SN, 2))
    End If
    For SubI = 1 To Counter
      If ModItem.Item(MainItemName & SubItems & SubItemsName(SubI)) Is Nothing Then
        ModItem.add MainItemName & SubItems & SubItemsName(SubI)
        ModItem.Item(MainItemName & SubItems & SubItemsName(SubI)).MenuItem = SubItemsName(SubI)
        MRow = CRow + SubI
        SortArray(SeatNo, MRow, 2) = Val(SortArray(SeatNo, MRow, 2))
      Else
        ModItem.Item(MainItemName & SubItems & SubItemsName(SubI)).MenuItem = SubItemsName(SubI)
        SortArray(SeatNo, SubI, 1) = "Counted"
        SortArray(SeatNo, MRow, 2) = Val(SortArray(SeatNo, MRow, 2)) + Val(SortArray(SeatNo, SubI, 2))
      End If
    Next SubI
   End If
  End If
  'SN = SubI - 1
Next SN
Done:
 
 ReDim TempArray(0, MaxRows, TTL_Col)
 Counter = -1
 For SN = 0 To TempNetRow(SeatNo) - 1
   If SortArray(SeatNo, SN, 1) <> "Counted" Then
    Counter = Counter + 1
    For SubI = 0 To TTL_Col
     TempArray(SeatNo, Counter, SubI) = SortArray(SeatNo, SN, SubI)
    Next SubI
   End If
 Next SN
 TempNetRow(SeatNo) = Counter
 Erase SortArray
 Exit Function
ERH:
MsgBox Err.Description '
'Resume
End Function
Public Function SeatConsolidate(TempArray() As Variant, TempNextRow() As Integer, SortNip As Integer, ReceiptTicket As Boolean)
On Error GoTo ErrorHandler
'**************************
Dim Rw            As Integer
Dim T             As Integer
Dim LS            As Integer
Dim SM            As Integer
Dim Cr            As Integer
Dim MainItem      As String
Dim ModSale       As Currency
Dim MainQty       As Long
Dim ModItems      As String
Dim MasterArray() As Variant
ReDim MainArray(Number_In_Party, MaxRows, TTL_Col)
Dim AllMods As SortAll
Dim AllSort As SortAll
Dim AHolder As SortAll
Dim MHolder As SortAll
Dim SubItem As SortAll
Dim SubList As SortAll
Set AllSort = New SortAll
Set AllMods = New SortAll
Set AHolder = New SortAll
Set MHolder = New SortAll
Set SubItem = New SortAll
Set SubList = New SortAll
Dim SubItems As String
Dim SubPrice As Currency
Dim L As Integer
Dim f As Integer
Dim Sp As String
Dim R  As Integer
Dim TNIP As Long
Dim NIP  As Integer
Dim SN   As Integer
TNIP = Number_In_Party
'CustOption.GroupRemoteItems = "On"
If GetConfigInfo("General", "Group Remote Item it 1st Seat", "Off", "Functions.INI") = "ENABLED" And ReceiptTicket = False Then
  Dim OneSeat() As Variant
  Cr = -1
  OneSeat() = TempArray()
  ReDim TempArray(0, MaxRows, TTL_Col)
  For NIP = 0 To Number_In_Party - 1
   For SN = 0 To TempNextRow(NIP) - 1
     Cr = Cr + 1
     If OneSeat(NIP, SN, 1) = "" Then Exit For
     For LS = 0 To TTL_Col
        TempArray(0, Cr, LS) = OneSeat(NIP, SN, LS)
     Next LS
   Next SN
  Next NIP
  TempNextRow(0) = Cr + 1
  Number_In_Party = 1
End If
 For SN = 0 To TempNextRow(NIP) - 1
   If IsEmpty(TempArray(NIP, SN, 1)) Then
     TempNextRow(NIP) = SN: Exit For
   End If
 Next SN
For NIP = 0 To Number_In_Party - 1
 Sp = CStr(NIP)
 For SN = 0 To TempNextRow(NIP) - 1
   If IsEmpty(TempArray(NIP, SN, 1)) Then Exit For
   If Left(TempArray(NIP, SN, 1), 2) <> "h-" Then
   If Check_Option(36) = 1 And Val(TempArray(NIP, SN, 0)) > 0 And Val(TempArray(NIP, SN, 2)) = 0 And ReceiptTicket Then
         TempArray(NIP, SN, 1) = "<>"
   End If
   If Val(TempArray(NIP, SN, 0)) > 0 And Val(TempArray(NIP, SN, 13)) = 0 Then
            If AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))) Is Nothing Then
               AllSort.add Sp & CStr(TempArray(NIP, SN, 1))
               AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).MenuItem = TempArray(NIP, SN, 1)
               AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).qty = 0
               AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).SaleAmt = 0
            End If
            AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).qty = AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).qty + Val(TempArray(NIP, SN, 0))
            AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).SaleAmt = AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).SaleAmt + Val(TempArray(NIP, SN, 2))
            MainQty = Val(TempArray(NIP, SN, 0))
            MainItem = Sp & AllSort.Item(Sp & CStr(TempArray(NIP, SN, 1))).MenuItem
            SubItems = "": SubPrice = 0
   End If
   If Val(TempArray(NIP, SN, 0)) = 0 And Val(TempArray(NIP, SN, 13)) = 0 Then
    For L = SN To TempNextRow(NIP) - 1
      If Val(TempArray(NIP, L, 0)) = 0 And Val(TempArray(NIP, L, 13)) = 0 Then
       If TempArray(NIP, L, 1) <> "<>" Then
        SubItems = SubItems & TempArray(NIP, L, 1)
        SubPrice = SubPrice + Val(TempArray(NIP, L, 2))
       End If
      Else
       Exit For
      End If
   Next L
   'If MainItem > "" Or SubItems > "" Then
   If MainItem & SubItems = "" Then GoTo NRW
   If SubItem.Item(MainItem & SubItems) Is Nothing Then
       SubItem.add MainItem & SubItems
       SubItem.Item(MainItem & SubItems).MenuItem = MainItem
       SubItem.Item(MainItem & SubItems).qty = 0
       SubItem.Item(MainItem & SubItems).SaleAmt = 0
       SubItem.Item(MainItem & SubItems).SubItem = MainItem & SubItems
   End If
   
   SubItem.Item(MainItem & SubItems).qty = SubItem.Item(MainItem & SubItems).qty + MainQty
    For f = SN To L - 1
      If AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)) Is Nothing Then
       AllMods.add MainItem & SubItems & TempArray(NIP, f, 1)
       AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)).ModName = TempArray(NIP, f, 1)
       AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)).MenuItem = MainItem
       AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)).qty = 0
       AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)).SaleAmt = 0
      End If
      AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)).SaleAmt = AllMods.Item(MainItem & SubItems & TempArray(NIP, f, 1)).SaleAmt + Val(TempArray(NIP, f, 2))
    Next f
    SN = f - 1
   End If
 End If
Next SN
Next NIP
Dim MSP As Integer
Dim MsS As Integer
Dim FI  As Integer
Dim FM  As Integer
NSI:
For NIP = 0 To Number_In_Party - 1
    Cr = -1
    Sp = CStr(NIP)
    For SN = 0 To TempNextRow(NIP) - 1
        If Val(TempArray(NIP, SN, 13)) = 0 Then
        If TempArray(NIP, SN, 1) = "" Then Exit For
            If TempArray(NIP, SN, 1) <> "<>" And Left(TempArray(NIP, SN, 1), 2) <> "h-" Then
                MainItem = Sp & TempArray(NIP, SN, 1)
                'If Len(MainItem) > 0 Then
                If Not AllSort.Item(MainItem) Is Nothing Then
                    If AllSort.Item(MainItem).MenuItem = TempArray(NIP, SN, 1) And Val(TempArray(NIP, SN, 0)) > 0 Then
                        If AHolder.Item(MainItem) Is Nothing Then
                          'If Len(MainItem) = 0 Then Stop
                          AHolder.add MainItem
                          Cr = Cr + 1
                          For LS = 0 To TTL_Col
                            MainArray(NIP, Cr, LS) = TempArray(NIP, SN, LS)
                          Next LS
                          MainArray(NIP, Cr, 24) = NIP  'seat seperation
                          MainArray(NIP, Cr, 0) = AllSort.Item(MainItem).qty
                          MainArray(NIP, Cr, 2) = AllSort.Item(MainItem).SaleAmt
                          MainArray(NIP, Cr, 1) = MainArray(NIP, Cr, 1)
                          If MainArray(NIP, Cr, 28) > "" Then
                             MainArray(NIP, Cr, 28) = MainArray(NIP, Cr, 28)
                          Else
                             MainArray(NIP, Cr, 28) = MainArray(NIP, Cr, 1) '& AllSort.Item(Rw).SeatNo
                          End If
                          SubItems = ""
                          For MsS = SN To TempNextRow(NIP) - 1            'SN
                              If TempArray(NIP, MsS, 1) = "" Then Exit For
                              If TempArray(NIP, MsS, 1) <> "<>" Then
                                If Val(TempArray(NIP, MsS, 0)) = 0 And Val(TempArray(NIP, MsS, 13)) = 0 Then
                                    For L = MsS To TempNextRow(NIP) - 1
                                      If TempArray(NIP, L, 1) <> "<>" Then
                                        If Val(TempArray(NIP, L, 0)) = 0 And Val(TempArray(NIP, L, 13)) = 0 Then
                                            SubItems = SubItems & TempArray(NIP, L, 1)
                                        Else
                                            Exit For
                                        End If
                                      End If
                                    Next L
                                    MsS = L
                                    If Not SubItem.Item(MainItem & SubItems) Is Nothing Then
                                        If SubList.Item(MainItem & SubItems) Is Nothing Then
                                            SubList.add MainItem & SubItems
                                            If ReceiptTicket = False Then
                                                Cr = Cr + 1
                                                For LS = 0 To TTL_Col
                                                    MainArray(NIP, Cr, LS) = TempArray(NIP, SN, LS)
                                                Next LS
                                                If ReceiptTicket = True Then
                                                    MainArray(NIP, Cr, 2) = ""
                                                End If
                                                MainArray(NIP, Cr, 1) = Trim$(CStr(SubItem.Item(MainItem & SubItems).qty) & "---> " & ReturnLabel("With"))
                                                MainArray(NIP, Cr, 0) = ""
                                                MainArray(NIP, Cr, 24) = NIP
                                                MainArray(NIP, Cr, 28) = MainArray(NIP, Cr, 28) 'was MainArray(NIP, Cr, 1)
                                            End If
                                            TempArray(NIP, SN, 30) = "<>"
                                            TempArray(NIP, SN, 1) = "<>"
                                            TempArray(NIP, SN, 12) = 0
                                        End If
                                        For L = SN To TempNextRow(NIP) - 1
                                           If Len(TempArray(NIP, L, 1)) > 0 Then
                                            If Not AllMods.Item(MainItem & SubItems & TempArray(NIP, L, 1)) Is Nothing Then
                                                If MHolder.Item(MainItem & SubItems & TempArray(NIP, L, 1)) Is Nothing Then
                                                    MHolder.add MainItem & SubItems & TempArray(NIP, L, 1)
                                                    Cr = Cr + 1
                                                    For LS = 0 To TTL_Col
                                                        MainArray(NIP, Cr, LS) = TempArray(NIP, L, LS) 'was msp
                                                    Next LS
                                                    MainArray(NIP, Cr, 24) = NIP
                                                    MainArray(NIP, Cr, 0) = ""
                                                    If ReceiptTicket = False Then
                                                        MainArray(NIP, Cr, 1) = String(4, Chr(32)) & TempArray(NIP, L, 1)
                                                    End If
                                                    MainArray(NIP, Cr, 2) = AllMods.Item(MainItem & SubItems & TempArray(NIP, L, 1)).SaleAmt    'sm
                                                    MainArray(NIP, Cr, 28) = MainArray(NIP, Cr, 28) 'was MainArray(NIP, Cr, 1)
                                                    TempArray(NIP, L, 1) = "<>"
                                                    TempArray(NIP, L, 30) = "<>"
                                                    TempArray(NIP, L, 12) = 0
                                                End If
                                            End If
                                          End If
                                        Next L
                                      SubItems = ""
                                   End If
                                End If
                              End If
                            SubItems = ""
                          Next MsS
                        End If
                    End If
                End If
            End If
           MainItem = ""
        End If
     Next SN
     TempNextRow(NIP) = Cr + 1
   Next NIP
NRW:
Number_In_Party = TNIP
SortNip = TNIP
TempArray() = MainArray()
Set AllSort = Nothing
Set AllMods = Nothing
Set AHolder = Nothing
Set MHolder = Nothing
Set SubItem = Nothing
Set SubList = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "SeatConsolidate")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function ALLClockedOut() As Boolean
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ALLClockedOut = True
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Select Case Input_Option(53)
  Case Is = "No"
    OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
      Sql = "Select * from [Employee Records] where "
      Sql = Sql & " [time in] > '" & "" & "' and [time out] = '" & "" & "'"
    OpenTb Sql, RS, Db
    If Not RS.EOF Then
       ALLClockedOut = False
       See_Msg "All employees must be clocked out", 0, "Can not close system"
     '  Time_Reports.SSCommand7_Click
     '  Call MemoryCheck(Time_Reports)
       If See_Msg("Ok to Continue to Close Day?", 4, "Continue Close") = 6 Then ALLClockedOut = True
    Else
       ALLClockedOut = True
    End If
    DbClose Db, RS
    Exit Function
  Case Is = "Out_In", "Out"
    PunchEveryOneOut
    ALLClockedOut = True
  Case Else
    ALLClockedOut = True
End Select
Exit Function
ErrorHandler:
 ALLClockedOut = True
End Function
Public Function ExemptSales(EServer As String, StartNo As Long) As Currency
On Error GoTo ErrorHandler
ExemptSales = 0
Dim EDb As ADODB.Connection, Etb As ADODB.Recordset
Dim TotalSale As Currency
ConnectADODB SvrPath.TerminalPath & CheckStatDb, EDb
    Sql = "Select * from [Check Detail] where [servers Name] = '" & EServer & "'"
    Sql = Sql & " and [close out day] = 0 and [Voided] = 0 "
If StartNo > 0 Then
    Sql = Sql & " and [Check number] > " & StartNo - 1
End If
ConnectADORS Sql, EDb, Etb
Do While Not Etb.EOF
    If TipExempt(Etb![Cateory Name]) = True Then ExemptSales = ExemptSales + Val(Etb![Price] & "")
    Etb.MoveNext
Loop
SetNothing Etb
  Sql = "Select * from [Closed Checks] where [server Name] = '" & EServer & "'"
  Sql = Sql & " and [close out day] = 0"
  If StartNo > 0 Then
    Sql = Sql & " and [Check number] > " & StartNo - 1
  End If
ConnectADORS Sql, EDb, Etb
TotalSale = 0
Do While Not Etb.EOF
   If Etb![Order Destination] > "" Then
    If OrderFor(SaleDestination(Etb![Order Destination])).TipExempt = 1 Then
       TotalSale = Etb![Gross Sales] - Etb![Voided Sales]  '    (Val(Etb![Gross Sales Tax] & "") - Val(Etb![Voided Sales Tax])) + (Val(Etb![Gross Special Tax] & "") - Val(Etb![Voided Special tax]))
       'TotalSale = TotalSale - Etb![Voided Sales]
       TotalSale = TotalSale - (Abs(Etb![Gross Discounts]) - Abs(Etb![Voided Discounts]))
       ExemptSales = ExemptSales + TotalSale
    End If
   End If
  Etb.MoveNext
Loop
DbClose EDb, Etb
Exit Function
ErrorHandler:
Check_no = 0
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Exempt Sales ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function TipExempt(Cat_Name) As Boolean
On Error Resume Next
TipExempt = False
Dim TF As Integer
For TF = 1 To Cats.count
    If Cats.Item(TF).TipExempt = "Yes" Then
        If Cats.Item(TF).CategoryName = Cat_Name Then
            TipExempt = True
            Exit Function
        End If
    End If
Next TF
TipExempt = False
End Function
Public Function GetReturns(ReturnList As FinancialItems) As Currency
''On Error Resume Next
GetReturns = 0
Dim Cdb As ADODB.Connection
Dim Tb As ADODB.Recordset
   ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
     Sql = "Select * from [Debit] where [Tender] = 'EMV Refund' and [Close Out day] = 0 "
   ConnectADORS Sql, Cdb, Tb
   Do While Not Tb.EOF
       If ReturnList.Item(Tb![Account Number]) Is Nothing Then
         ReturnList.add (Tb![Account Number])
       End If
       ReturnList.Item(Tb![Account Number]).ItemName = Tb![Account Number]
       ReturnList.Item(Tb![Account Number]).ItemName1 = Tb.fields("Account Name").Value & ""
       ReturnList.Item(Tb![Account Number]).Amount = Tb![Amount]
       GetReturns = GetReturns + Tb![Amount]
      Tb.MoveNext
   Loop
DbClose Cdb, Tb
End Function
Public Function CentsTax(TableNo, Price, TaxRate) As Currency
On Error GoTo ErrorHandler
CentsTax = 0
If Price = 0 Or TableNo = 0 Then
Exit Function
End If
Dim SRC As Integer
For SRC = 0 To UBound(TaxBreak, 3)
If Price >= TaxBreak(TableNo, 0, SRC) And Price <= TaxBreak(TableNo, 1, SRC) Then
   CentsTax = TaxBreak(TableNo, 2, SRC)
   Exit Function
End If
Next SRC
For SRC = 0 To UBound(TaxBreak, 3)
If Price - Int(Price) >= TaxBreak(TableNo, 0, SRC) And Price - Int(Price) <= TaxBreak(TableNo, 1, SRC) Then
   CentsTax = TaxBreak(TableNo, 2, SRC)
   CentsTax = CentsTax + Int(Price) * TaxRate
   Exit Function
End If
Next SRC
Exit Function
ErrorHandler:
CentsTax = 0
Exit Function
End Function
Public Function GetMinimumDelivery(ZoneIn As String) As Currency
On Error GoTo ErrorHandler
  Dim DiscDb As ADODB.Connection
  Dim Tb As ADODB.Recordset
  OpenDb SvrPath.LocalTerminalPath & "MSaleData", DiscDb
   Sql = "Select * from [Discounts] where [Discount Name] = '" & ZoneIn & "'"
  ConnectADORS Sql, DiscDb, Tb
  
  If Not Tb.EOF Then
     GetMinimumDelivery = Val(Tb![Max Discount Amount] & "")
  End If
  DbClose DiscDb, Tb
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function ShowKeyBoard(KB_Msg As String) As String
On Error GoTo ErrorHandler
                Dim KeyB As CallKeyBoard
                Set KeyB = New CallKeyBoard
                Sleep 0
                ShowKeyBoard = KeyB.ShowKeyBoard(KB_Msg, SvrPath.TerminalPath, 0, TaxIsOff, False)
                Set KeyB = Nothing
                If InStr(ShowKeyBoard, "'") > 0 Then ShowKeyBoard = Replace(ShowKeyBoard, "'", "`")
                If InStr(ShowKeyBoard, "\") > 0 Then ShowKeyBoard = Replace(ShowKeyBoard, "\", "-")
                If InStr(ShowKeyBoard, ",") > 0 Then ShowKeyBoard = Replace(ShowKeyBoard, ",", ".")
                Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Show Keyboard ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetNewTableName() As String
  GetNewTableName = ShowKeyBoard("Enter New Table/Tab Name")
End Function
Public Function PhoneOrdersAreOpen(Svr_Name As String) As Integer
PhoneOrdersAreOpen = -1
If Check_Option(32) = 0 Then Exit Function
If Check_Option(116) = 0 Then Exit Function
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
 Sql = "Select * from [Open Checks] where [Presented] = 1"
 If Svr_Name > "" Then
  Sql = Sql & " and [Server Name] = '" & Svr_Name & "'"
 End If
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  PhoneOrdersAreOpen = RecordsCount(RS)
End If
DbClose Db, RS
If PhoneOrdersAreOpen > -1 Then
   See_Msg "You Have " & PhoneOrdersAreOpen & " Open Phone Orders", 5, "Open Phone Orders Assigned To You."
End If
End Function
Public Function ShowNumberPad(NMsg As String, QtNeeded As Boolean, BoxLeft As Long, BoxTop As Long) As Variant
On Error GoTo ErrorHandler
                Dim NPad As CallNumberPad
                Set NPad = New CallNumberPad
                Sleep 0
                ShowNumberPad = NPad.ShowNumberPad(NMsg, QtNeeded, BoxLeft, BoxTop, False)
                
                Set NPad = Nothing
                'If ShowNumberPad = "" Then ShowNumberPad = 0
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Show Number Pad ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Function
Public Function GetCovers() As Integer
On Error GoTo ErrorHandler
                Dim NPad As CallCovers
                Set NPad = New CallCovers
                Sleep 0
                GetCovers = NPad.ShowCovers
                'FixedCoverCount = GetCovers
                Set NPad = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Covers ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetCashier(DrwNo As Integer) As String
On Error GoTo ErrorHandler
Dim fdb As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim MasterDTName As String
Dim sDb As ADODB.Connection
        If System_config = ExpressService Or Check_Option(34) = 1 Then
            'Must test for c closing var -- will not set bact to\
            'cashier after closing check
              If Check_Option(180) = 1 Then
                ConnectADODB SvrPath.TerminalPath & RegisterOptionsDB, sDb
                  Sql = "Select * from [Data Paths]"
                  ConnectADORS Sql, sDb, Tb
                   If Not Tb.EOF Then
                     MasterDTName = Tb.fields("Register Name").Value & ""
                   Else
                     MasterDTName = MyRegisterName & ""
                   End If
                   DbClose sDb, Tb
                   Sql = "Select * from [Audits] Where [Register Name] = '" & MasterDTName & "' order by [pos] asc"
              Else
                If Check_Option(60) = 1 Then
                  Sql = "Select * from [Audits] Where [Register Name] = '" & MyRegisterName & "' order by [pos] asc"
                Else
                  Sql = "Select * from [Audits] Where [Pos] = " & DrwNo & " and [Register Name] = '" & MyRegisterName & "'"
                End If
              End If
              ConnectADODB SvrPath.TerminalPath & FinancialDb, fdb
              ConnectADORS Sql, fdb, Tb
               If Not Tb.EOF Then
                  GetCashier = Tb![Name]
               Else
                  GetCashier = LoggedOnAs(0).EMPName
               End If
               DbClose fdb, Tb
        ElseIf System_config = FullService Then
                ConnectADODB SvrPath.TerminalPath & "MSaleData", fdb
                Sql = "Select * from [check numbers] where [Drawer No] = " & DrwNo & " and [Audit Status] = 0 "
                Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
                ConnectADORS Sql, fdb, Tb
                    If Not Tb.EOF Then
                       GetCashier = Tb![Servers Name]
                    Else
                       GetCashier = "Not Assigned"
                    End If
                DbClose fdb, Tb
        Else
               GetCashier = "Closed"
        End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Cashier " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function AdvanceSelection() As Long
Selection_Number = Selection_Number + 1
AdvanceSelection = Selection_Number
End Function
Public Function GetLastSelectionNumber(cn As Long, Db As ADODB.Connection) As Long
On Error GoTo ErrorHandler
   Dim RS As ADODB.Recordset
   Sql = "Select top 1 [Selection Number] From [Check Detail] where [Check Number] = " & cn & " and [Close Out Day] = 0 Order by [selection number] asc "
   ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
      GetLastSelectionNumber = RS![Selection Number]
   Else
      GetLastSelectionNumber = 1
   End If
   SetNothing RS
Exit Function
ErrorHandler:
GetLastSelectionNumber = 500
End Function
Public Function GetLastRowOnCheck(cn As Long, Db As ADODB.Connection) As Long
On Error GoTo ErrorHandler
   Dim RS As ADODB.Recordset
   Sql = "Select count([Pos]) as Rows From [Check Detail] where [Check Number] = " & cn & " and [Close Out Day] = 0 "
   ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
      GetLastRowOnCheck = RS![rows] + 1
   Else
      GetLastRowOnCheck = 1
   End If
   SetNothing RS
   '// count the rows on the check and insert the item on the last row +1
Exit Function
ErrorHandler:
GetLastRowOnCheck = 500
End Function
Public Function IsProcessFeeZero() As Currency
IsProcessFeeZero = 0
Dim SC As Integer
   For SC = 0 To NextRow(SeatNumber)
    If Left$(Grid_array(SeatNumber, SC, 1), 1) = "*" Then
      IsProcessFeeZero = Val(Grid_array(SeatNumber, SC, 2))
      Exit For
    End If
   Next SC
End Function
Public Sub MemoryCheck(ByRef f As Form)
    If Not f Is Nothing Then
        Unload f
        'Set F = Nothing
        'this setting will blow up the software
    End If
End Sub
Public Function ApplyPAXTip(CheckNo As Long, Svr_Name As String, TipAmt As Currency, AuthCode As String, DoTip As ApplyTipsToEMV, AcctNo As String, CardBrand As String) As String
On Error GoTo ErrorHandler
ApplyPAXTip = ""
Dim DeviceIp            As String
Dim Answer              As String
Dim OriginalAuthAmount  As Currency
Dim FeeAmount           As Currency
Dim regName             As String
Dim OriginalAuthCode    As String
Dim TipCheck            As Long
Dim TPN                 As String
Dim DJVToken            As String
'Dim OrgTimeOfSale As String

Dim DCAC As ver11DebitCards.Accounting
If Not DCAC Is Nothing Then Exit Function
Set DCAC = New ver11DebitCards.Accounting
Dim AccountFound        As Boolean
Dim GetAccount          As Boolean
Dim AcctData()          As ver11DebitCards.AccountInfo
ReDim AcctData(0)
Dim ScoInfo             As ver11DebitCards.ReturnedResponse
OriginalAuthCode = AuthCode
    AcctData(0).CheckNo = CheckNo
    AcctData(0).ServerName = Svr_Name
    AcctData(0).AuthCode = AuthCode
    AcctData(0).CardNumber = AcctNo
    GetAccount = DCAC.GetAccountChgData(AcctData(), SvrPath.TerminalPath)
    AcctData(0).CheckNo = CheckNo
    AcctData(0).ServerName = Svr_Name
    If GetAccount = False Then
       See_Msg "Information Not Found", 5, "Look Up for transacation failed.."
       Set DCAC = Nothing
       Exit Function
    End If
    If GetAccount Then
            Dim xml As MPxmlParse
            Dim Sales As AllEMVSales
            Set Sales = New AllEMVSales
            If Sales.Item(CStr(AcctData(0).CheckNo)) Is Nothing Then
              Sales.add CStr(AcctData(0).CheckNo)
            End If
            DeviceIp = ""
            Answer = ""
            'Dim OrderIdForDJV As String
            Set xml = New MPxmlParse
            xml.xml = AcctData(0).XMLDataString
            'OrderIdForDJV = xml.token("OriginalPNREF")
            FeeAmount = Val(xml.token("FeeAmount"))
            If IsDejavooInterface(EMVINTERFACE) Then
             If xml.token("TPN") > "" Then
               TPN = xml.token("TPN")
               DJVToken = ReturnDejavooToken(TPN)
             End If
            End If
            If IsPAXInterface(EMVINTERFACE) Then
              If xml.token("PaxIp") > "" Then
                DeviceIp = xml.token("PaxIp")
              ElseIf xml.token("IPAddress") > "" Then
                DeviceIp = xml.token("IPAddress")
              Else
                 DeviceIp = GetPaxIpAddress(AcctData(0).RegisterName)
                 If DeviceIp = "" Then
                   See_Msg "Cannot apply tip at this work station", 0, "Must Use " & AcctData(0).RegisterName
                   Exit Function
                 End If
              End If
            Else
              If Val(xml.token("Authorize")) > 0 Then
                OriginalAuthAmount = Val(xml.token("Authorize"))
              End If
            End If
            Set xml = Nothing
            Trans_Log "sending Tip From " & AcctData(0).RegisterName & " " & AcctData(0).CheckNo
 
            Sales.Item(CStr(AcctData(0).CheckNo)).OrgAuthCode = AcctData(0).AuthCode
            Sales.Item(CStr(AcctData(0).CheckNo)).Cashier = Svr_Name
            Sales.Item(CStr(AcctData(0).CheckNo)).Term = AcctData(0).RegisterName
            Sales.Item(CStr(AcctData(0).CheckNo)).IpAddress = DeviceIp
            If INCREMENTAL > "" Then '// added to keep amex and discover from declining with incremental auth
               'If UCase(CardBrand) <> "AMEX" And UCase(CardBrand) <> "DISCOVER" Then
                If OriginalAuthAmount = AcctData(0).ckTotal Then
                 Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(AcctData(0).ckTotal, "0.00")
                 Sales.Item(CStr(AcctData(0).CheckNo)).TipAmt = TipAmt
                Else
                 Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(AcctData(0).ckTotal + TipAmt, "0.00")
                 Sales.Item(CStr(AcctData(0).CheckNo)).TipAmt = 0  'TipAmt
                End If
            Else
              Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(AcctData(0).ckTotal, "0.00")
              Sales.Item(CStr(AcctData(0).CheckNo)).TipAmt = TipAmt
            End If
            Sales.Item(CStr(AcctData(0).CheckNo)).XMLString = AcctData(0).XMLDataString
            Select Case AcctData(0).ChargeType
              Case Is = 1
                Sales.Item(CStr(AcctData(0).CheckNo)).TranType = "Debit"
              Case Is = 2
                Sales.Item(CStr(AcctData(0).CheckNo)).TranType = "EBT"
              Case Else
                Sales.Item(CStr(AcctData(0).CheckNo)).TranType = "Credit"
            End Select
            Sales.Item(CStr(AcctData(0).CheckNo)).TransNo = CStr(CheckNo)
            Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
            If LoadEMVDevice(AcctData(0).RegisterName, EmvEquip) = "Failed" Then
             See_Msg "No Device Found", 5, "Check Device Setup"
             Exit Function
            End If
            
            Dim PMG As PAXMSG
            Set PMG = New PAXMSG
            Dim SIFO As EMVProcessLib.Process
            If IsDejavooInterface(EMVINTERFACE) And EmvEquip.Item(1).MaxQtyOffline = 1 Then 'used for regisiter with processor for processing fees
              Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt = Format(AcctData(0).ckTotal - FeeAmount, "0.00")
            End If
            
            If IsDataCapInterface(EMVINTERFACE) Then
              If EMVSUPPORTPREAUTH = 1 And IsPadPresent = False Then Exit Function
              Set SIFO = New Process
              If EMVSUPPORTPREAUTH = 1 Then
                If DoTip = DoEMVCapture Then
                  If INCREMENTAL > "" Then
                   If OriginalAuthAmount > 0 And OriginalAuthAmount <> Sales.Item(CStr(AcctData(0).CheckNo)).SalesAmt Then
                    Answer = SIFO.ProcessSale(Sales, DevSaleType_IncrementalAuthByRecordNo, PMG, EmvEquip)
                    If UCase(PMG.Item(1).ResultTxt) = "APPROVED" Then
                     Sales.Item(CStr(AcctData(0).CheckNo)).XMLString = Answer
                    End If
                    Trans_Log "Incremental Auth " & Answer
                   End If
                   'must send the new return xml string
                  End If
                  Answer = SIFO.ProcessSale(Sales, DevSaleType_PreAuthCaptureByRecordNo, PMG, EmvEquip)
                Else
                 GoTo DoApproved
                End If
              Else
                 Set SIFO = New Process
                 Answer = SIFO.ProcessSale(Sales, DevSaleType_AdjustByRecordNo, PMG, EmvEquip)
              End If
            ElseIf IsDejavooInterface(EMVINTERFACE) Then
             If TPN > "" Then
               EmvEquip.Item(1).deviceId = TPN
               EmvEquip.Item(1).MerchantId = TPN
               If DJVToken > "" Then
                 EmvEquip.Item(1).MerchantToken = DJVToken
               End If
             End If
             Set SIFO = New Process
             If DoTip = ApplyEMVTip Then
              Answer = SIFO.ProcessSale(Sales, DevSaleType_djvTipAdjust, PMG, EmvEquip)
             Else
              Answer = SIFO.ProcessSale(Sales, DevSaleType_djvCapture, PMG, EmvEquip)
              MarkRecordCaptured AcctData(0).CheckNo
             End If
             Trans_Log Answer
             'Trans_Log "Process Tip " & Answer & " From " & AcctData(0).RegisterName & " " & AcctData(0).Check_no
            Else  '// PAX
              Set SIFO = New Process
                If DoTip = SkipEMVCapture Then
                  GoTo DoApproved
                End If
                If DoTip = ApplyEMVTip Then
                 Answer = SIFO.ProcessSale(Sales, DevSaleType_XAdjByRecordNo, PMG, EmvEquip)
                ElseIf PAXBARTABPREAUTH = 1 Then
                  If DoTip = DoEMVCapture Then
                   Answer = SIFO.ProcessSale(Sales, DevSaleType_PPostAuth, PMG, EmvEquip)
                  Else
                   GoTo DoApproved
                  End If
                End If
            End If
            Set Sales = Nothing
            Trans_Log EMVINTERFACE
            Trans_Log "Process Tip " & Answer & " From " & AcctData(0).RegisterName & " " & AcctData(0).CheckNo
            If Answer = "POSLINK TRANSACTION ERROR: SEND DATA RECV ACK ERROR" Then
              If See_Msg("Device response " & Answer, 4, "Stopping the batch") = 6 Then
               ApplyPAXTip = "Error " & Answer
               Exit Function
              End If
            End If
            If Answer = "POSLINK TRANSACTION ERROR: CONNECT ERROR" Then
              If See_Msg("Device response " & Answer, 4, "Stopping the batch") = 6 Then
               ApplyPAXTip = "Error " & Answer
               Exit Function
              End If
            End If
            If PMG.Item(1) Is Nothing Then
               See_Msg "Device - No response", 5, Answer
               ApplyPAXTip = "Device No response Error " & Answer
               Exit Function
            End If
            Trans_Log "EMV Respone " & Answer & " Response for " & AcctData(0).CardNumber
            Trans_Log "Message Response " & PMG.Item(1).Message
            If InStr(UCase(PMG.Item(1).ResultTxt), "APPROVED") > 0 Or InStr(UCase(PMG.Item(1).Message), "APPROVED") > 0 Then
            'If UCase(PMG.Item(1).ResultTxt) = "APPROVED" Or UCase(PMG.Item(1).Message) = "APPROVED OFFLINE" Then
               If DoTip = DoEMVCapture Then
                  ApplyPAXTip = "Approved"
                  Set SIFO = Nothing
                  Set PMG = Nothing
                  Exit Function
               End If
DoApproved:
               ApplyPAXTip = "Approved"
               With AcctData(0)
                  .RegisterName = MyRegisterName
                  .CashierName = ReturnMyCashier
                  .AccountName = ""
                  If AcctData(0).AuthCode = "" Then AcctData(0).AuthCode = OriginalAuthCode
                  '// auth number may return null if a reversal is done
                  '// this makes sure that we done over write a good authorization.
                  .AuthCode = AcctData(0).AuthCode
                  .CheckNo = CheckNo
                  .ServerName = Svr_Name
                  .TenderName = "EMV"   'xmls.token("CardType")
                  .TipAmt = TipAmt
                  If IsPAXInterface(EMVINTERFACE) Then
                    .XMLDataString = ""
                  ElseIf IsDejavooInterface(EMVINTERFACE) Then
                    .XMLDataString = ""
                  Else
                    .XMLDataString = Answer
                  End If
              End With
              If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" And Input_Option(29) = "By Cashier" Then
                AssignCashier Svr_Name, ReturnMyCashier, CheckNo
              End If
              DCAC.ChargeAccount AcctData(0), SvrPath.TerminalPath
              SaveEMVPayments AcctData(0).TimeOfSale, AcctData(0).CheckNo, AcctData(0).ServerName, AcctData(0).ckTotal, AcctData(0).TipAmt, AcctData(0).AccountName
              Set DCAC = Nothing
              See_Msg "Tip Applied", 5, ""
              PrintTipSlip Svr_Name, "EMV Sale " & AcctData(0).AuthCode, "", CheckNo, Format(AcctData(0).ckTotal, "0.00"), TipAmt
            Else
              ApplyPAXTip = "Failed " & PMG.Item(1).Message
              See_Msg "Transaction Tip Failed ", 5, PMG.Item(1).Message
              Trans_Log "Tip apply failed " & PMG.Item(1).Message
            End If
            Set SIFO = Nothing
            Set PMG = Nothing
      End If
      Set DCAC = Nothing
Exit Function
ErrorHandler:
See_Msg Err.Description, 5, Err.Description
Trans_Log "EMV Process Failed " & Err.Description & " "
End Function
Public Function IsThisCheckOpen(CheckNo As Long) As Boolean
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
IsThisCheckOpen = False
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select * from [Open checks] where [Check Number] = " & CheckNo & " and [Close out Day] = 0"
  ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
     IsThisCheckOpen = True
   End If
  DbClose Db, RS
End Function
Public Function ReturnMyCashier() As String
 If Check_Option(34) = 1 Or System_config = ExpressService Then
    ReturnMyCashier = GetActiveCashier
    If ReturnMyCashier = "" Then ReturnMyCashier = LoggedOnAs(0).EMPName
 Else
    ReturnMyCashier = GetWhoClosedCheck
 End If
End Function
Public Function IsPadPresent() As Boolean
On Error GoTo ERH
Dim R As String
Dim xmls As MPxmlParse
Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
If LoadEMVDevice(MyRegisterName, EmvEquip) = "Failed" Then
   See_Msg "No Device Found", 5, "Check Device Setup"
   Exit Function
End If
Dim DC As EMVProcessLib.Process
Set DC = New EMVProcessLib.Process
R = DC.PadReset(EmvEquip)
Set DC = Nothing
Set xmls = New MPxmlParse
xmls.xml = R
If xmls.token("CmdStatus") <> "Success" Then
   See_Msg "No EMV Device Found", 5, "Pin Pad Required"
   IsPadPresent = False
Else
   IsPadPresent = True
End If
Set xmls = Nothing
Exit Function
ERH:
See_Msg "error", 5, Err.Description
End Function
Public Function DidYouForGet() As Boolean
On Error GoTo ERH
Dim Rw         As Integer
Dim St         As Integer
DidYouForGet = False
    If Required_Selection.count = 0 Then Exit Function
    For St = 0 To Number_In_Party - 1
     For Rw = 0 To NextRow(St) - 1
       If Not Required_Selection.Item(Grid_array(St, Rw, 23)) Is Nothing Then
        DidYouForGet = False
        Exit Function
       End If
     Next Rw
    Next St
    If See_Msg("", 4, ReturnLabel("Would You Like To Order Something From ") & Required_Selection.Item(1).Category & " ?") = 6 Then
       DidYouForGet = True
    Else
       DidYouForGet = False
    End If
Exit Function
ERH:
End Function
Public Function Required_Menu() As Integer
 On Error GoTo ErrorHandler
 Dim qty        As Integer
 Dim St         As Integer
 Dim Rw         As Integer
 Required_Menu = 0
    For St = 0 To Number_In_Party - 1
     For Rw = 0 To NextRow(St) - 1
       If Not Required_Selection.Item(Grid_array(St, Rw, 23)) Is Nothing Then
        qty = Val(Grid_array(St, Rw, 0))
         Required_Menu = Required_Menu + qty
         If Required_Menu >= Total_required * Number_In_Party Then
           Exit Function
         End If
       End If
     Next Rw
    Next St
    If Required_Menu < Number_In_Party Then
     Call See_Msg(Required_Selection.Item(1).Category & " Required (Item On Check)", 0, CStr(Abs(Number_In_Party - Required_Menu)) & " More Needed ")
    End If
Exit Function
ErrorHandler:
  See_Msg "Error with required Items", 5, Err.Description
  Required_Menu = 1
End Function
Public Function GETMonth(Mth As Integer)
Select Case Mth
 Case Is = 12
    GETMonth = "December"
 Case Is = 1
    GETMonth = "January"
 Case Is = 2
    GETMonth = "February"
 Case Is = 3
    GETMonth = "March"
 Case Is = 4
    GETMonth = "April"
 Case Is = 5
    GETMonth = "May"
 Case Is = 6
    GETMonth = "June"
 Case Is = 7
    GETMonth = "July"
 Case Is = 8
    GETMonth = "August"
 Case Is = 9
    GETMonth = "September"
 Case Is = 10
    GETMonth = "October"
 Case Is = 11
    GETMonth = "November"
End Select
End Function
Public Sub SendAKey(Vk_Code As Integer, scan%)
keybd_event Vk_Code, scan%, 0, 0
keybd_event Vk_Code, scan%, &H2, 0
End Sub
Public Function IsRing() As Boolean
IsRing = False
Dim TL                  As Integer
Dim CallerDeleteTime    As Integer

If Gross_Receipt > 0 Then Exit Function
'\\ added this line to make sure ringing will not show up if an active order is on the screen.

CallerDeleteTime = Val(GetConfigInfo("Phone Orders", "Caller Id Ring Expires in Seconds", "30", "Functions.INI"))
For TL = 1 To 4
 If IsDate(PhoneRing(TL).TimeIn) Then
  If DateDiff("s", PhoneRing(TL).TimeIn, Time) > CallerDeleteTime And PhoneRing(TL).AnsweredBy = "Ring" Then
   PhoneRing(TL).AnsweredBy = ""
   PhoneRing(TL).TimeIn = ""
  End If
 End If
 If PhoneRing(TL).AnsweredBy = "Ring" Then
  IsRing = True
  Exit Function
 End If
Next TL
End Function
Public Function KeyStroke_Ok(KeyTest As Integer) As Integer
Select Case KeyTest
Case Is = 34, 39, 40, 41, 38, 35, 42
  KeyStroke_Ok = 0
  Call See_Msg("Can't Use  " & Chr(KeyTest) & "  Here (Control character reserved by program)", 5, "Try a different character")
  Exit Function
Case Is = 39
   KeyStroke_Ok = 96
   Exit Function
End Select
KeyStroke_Ok = KeyTest
End Function
Public Function Redim_Grid()
On Error GoTo ermsg

Dim SN As Integer
Dim C  As Integer
Dim R  As Integer
Dim array_temp() As Variant
  ReDim array_temp(Number_In_Party, MaxRows, TTL_Col) As Variant
  array_temp() = Grid_array()
NXtS:
LoadArray:
 ReDim Grid_array(Number_In_Party, MaxRows, TTL_Col) As Variant
    For SN = 0 To Number_In_Party - 1
      For R = 0 To NextRow(SN) '- 1
       For C = 0 To TTL_Col
          Grid_array(SN, R, C) = array_temp(SN, R, C)
       Next C
    Next R
nxtR:
   Next SN
ReDim array_temp(0, 0, 0)
Erase array_temp
Order.Ticket(0).rows = MaxRows + 1
Exit Function
ermsg:
Exit Function
Resume Next
End Function
Public Function TerminalFailed(PathToTerminal As String) As Boolean
On Error GoTo ErrorHandler
If DisablePing = "ENABLED" Then
    TerminalFailed = False
Else
    TerminalFailed = False
    If PathToTerminal = "" Then Exit Function
    Dim ttpath As SetPaths
    Set ttpath = New SetPaths
    If ttpath.getRemoteStation(PathToTerminal) Then
     TerminalFailed = False
    Else
     TerminalFailed = True
    End If
    Set ttpath = Nothing
End If
Exit Function
ErrorHandler:
TerminalFailed = False
Set ttpath = Nothing
Trans_Log "Terminal Failed " & Err.Description & " " & PathToTerminal
End Function
Public Function GetPendingOrders() As Boolean
On Error GoTo ErrorHandler
GetPendingOrders = False
If Gross_Receipt > 0 Then Exit Function
'// makes sure no order is in progress when pending order is comming in.
If Check_Option(32) = 0 Then Exit Function
If Check_Option(241) = 1 Then Exit Function
If Val(Input_Option(21)) = 0 Then Exit Function 'Call ahead min's always on if set to 0
Dim Db  As ADODB.Connection
Dim Ctb As ADODB.Recordset
Dim CT  As SendCallin
Set CT = New SendCallin
    Sql = "Select * from [Open Checks] where "
    Sql = Sql & " [Presented] = 0 "
    If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
    ConnectADORS Sql, Db, Ctb
    Do While Not Ctb.EOF
      If CT.CheckTime(Ctb![Pick Up Time], Ctb![Pick Up Date], Val(Ctb![Production] & "")) Then
       GetPendingOrders = True
       DbClose Db, Ctb
       Exit Function
      End If
      Ctb.MoveNext
    Loop
    Set CT = Nothing
    DbClose Db, Ctb
Set CT = Nothing
Exit Function
ErrorHandler:
DbClose Db, Ctb
GetPendingOrders = False
Err.Clear
End Function
Public Function FileCheck(FileName As String) As Boolean
On Error GoTo erHand
Dim FC As Object
Set FC = CreateObject("Scripting.FileSystemObject")
FileCheck = FC.FileExists(FileName)
Set FC = Nothing
Exit Function
erHand:
FileCheck = False
End Function
Public Function SkimFlag() As Boolean
  On Error Resume Next
  If Val(Input_Option(39)) = 0 Then 'if the skim value is set to 0 then disable option.
    SkimFlag = False
    Exit Function
  End If
  Dim ff            As Integer
  Dim SkimCash      As Currency
  Dim SD            As Integer
  Dim SR            As Integer
  Dim CashierName   As String
  Dim TotalCash     As Currency
  Dim SkimThreshold As Currency
  SkimThreshold = Val(Input_Option(39))
  CashierName = Trim$(LoggedOnAs(0).EMPName)
  ff = FreeFile
  SkimCash = 0
   Open SvrPath.TerminalPath & "DAT Files\" & "CashDeposits.Dat" For Random As ff Len = Len(Sdeposit)
          For SD = 1 To LOF(ff) / Len(Sdeposit)
                Get #ff, SD, Sdeposit
                If Trim$(Sdeposit.SvrName) = CashierName Then
                 SkimCash = SkimCash + Sdeposit.DepositAmt
                End If
          Next SD
  Close ff
  Dim CID As ADODB.Connection
  Dim Crd As ADODB.Recordset
  If ConnectADODB(SvrPath.TerminalPath & FinancialDb, CID) = False Then Exit Function
   Sql = "SELECT SUM([Payment Amount]) AS TotalCash FROM [dayPayments] " & _
          "WHERE [Payment Type] = 'Cash' AND [Cashier Name] = ?"
          
   Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = CID
        .CommandText = Sql
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("CashierName", adVarChar, adParamInput, 50, CashierName)
    End With

    ' Execute the query
    Set Crd = cmd.Execute

    ' Calculate the skim flag
    If Not Crd.EOF Then
        TotalCash = Val(Crd!TotalCash & "") ' Handle Null values with Nz function
        If (TotalCash - SkimCash) >= SkimThreshold Then
            SkimFlag = True
        Else
            SkimFlag = False
        End If
    Else
        SkimFlag = False
    End If
  DbClose CID, Crd
End Function
Public Function GetDayOfWeek(DN As Integer) As String
  Select Case DN
     Case Is = 0
        GetDayOfWeek = "Today"
     Case Is = 1
        GetDayOfWeek = "Monday"
     Case Is = 2
        GetDayOfWeek = "Tuesday"
     Case Is = 3
        GetDayOfWeek = "Wednesday"
     Case Is = 4
        GetDayOfWeek = "Thursday"
     Case Is = 5
        GetDayOfWeek = "Friday"
     Case Is = 6
        GetDayOfWeek = "Saturday"
     Case Is = 7
        GetDayOfWeek = "Sunday"
  End Select
End Function
Public Function GetDateToNo(InDate As String) As Integer
 If IsDate(InDate) Then
   GetDateToNo = DatePart("w", InDate, WeekStartsOn, vbFirstJan1)
 Else
   GetDateToNo = 0
 End If
End Function
Public Function GetServerGroup(ReceivingServerIs As String) As String
On Error GoTo ErrorHandler
retry = 0
tryagain:
If Check_Option(39) = 1 Then
  GetServerGroup = SvrPath.TerminalGroupName
 Exit Function
End If

Dim getnewCknum As ADODB.Connection
Dim Records     As ADODB.Recordset
 If ConnectADODB("MSaleData", getnewCknum) = False Then Exit Function
   Sql = "Select * from [Check Numbers] where [Servers Name] = '" & ReceivingServerIs & "'"
   Sql = Sql & " and [Audit Status] = 0 "
 If ConnectADORS(Sql, getnewCknum, Records) = False Then Exit Function
 If Records.EOF Then
    GetServerGroup = SvrPath.TerminalGroupName
 Else
    If Records![Register Group] & "" = "" Then
       GetServerGroup = SvrPath.TerminalGroupName
    Else
       GetServerGroup = Records![Register Group] & ""
    End If
 End If
DbClose getnewCknum, Records
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Register Group " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetDriveThruTime() As String
On Error GoTo ErrorHandler
Dim getnewCknum As ADODB.Connection
Dim Records     As ADODB.Recordset
Dim Cars        As Long
Dim MinsOpen    As Long
MinsOpen = 0: Cars = 0
  If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, getnewCknum) = False Then Exit Function
    Sql = "Select * from [Closed Checks] where [Close Out Day] = 0 and [Order Destination] = 'Drive Thru'"
  If ConnectADORS(Sql, getnewCknum, Records) = False Then Exit Function
  Cars = Cars + RecordsCount(Records)
  Do While Not Records.EOF
     MinsOpen = MinsOpen + DateDiff("n", Records![Open Time], Records![Close Time])
    Records.MoveNext
  Loop
  DbClose getnewCknum, Records
  If Cars > 0 And MinsOpen > 0 Then
     GetDriveThruTime = Format(MinsOpen / Cars, "0.00")
  Else
     GetDriveThruTime = ""
  End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Drive Thru Time " & Sql)
   Select Case ErrorRespose
    Case "Abort"
        ' Exit the function immediately
        Exit Function
    Case "Resume"
        ' Retry the operation that caused the error
        Resume
    Case "Resume Next"
        ' Skip the error and move to the next line
        Resume Next
    Case Else
        ' Handle unexpected responses gracefully
        Debug.Print "Unexpected Error Response: " & ErrorRespose
        Exit Function
   End Select
End Function
Public Function HasCashierAudit(SvrName As String) As Boolean
On Error GoTo ErrorHandler
Dim fdb As ADODB.Connection
Dim Tb  As ADODB.Recordset
   HasCashierAudit = False
   If ConnectADODB(SvrPath.TerminalPath & FinancialDb, fdb) = False Then Exit Function
      'Sql = "Select * from [Audits] where [Name] = '" & SvrName & "'"
      Sql = "SELECT TOP 1 * FROM [Audits] WHERE [Name] = '" & Replace(SvrName, "'", "''") & "'"
      ConnectADORS Sql, fdb, Tb
      If Not Tb.EOF Then
         HasCashierAudit = True
      End If
      DbClose fdb, Tb
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function GetDeliveryDriver() As String
If SvrPath.GroupAudit Then
  GetDeliveryDriver = LoggedOnAs(0).EmpAssignedTo
Else
  GetDeliveryDriver = LoggedOnAs(0).EMPName
End If
  If OrderIs = "Delivery" Then
    If Not AllOrders Is Nothing Then
      If AllOrders.Item(1).DriverName > "" Then
        If GetShiftNo(AllOrders.Item(1).DriverName) > 0 Then
          GetDeliveryDriver = AllOrders.Item(1).DriverName
        ElseIf System_config = ExpressService Then
          GetDeliveryDriver = AllOrders.Item(1).DriverName
        Else
           'If LoggedOnAs(0).EMPName = "Web Order" Then
              '// do nothing sales we go to Web Order
           'Else
            ' See_Msg "No Driver Audit Found", 5, "Sale Transfered To " & LoggedOnAs(0).EMPName
           'End If
           
           AllOrders.Item(1).DriverName = LoggedOnAs(0).EMPName
           GetDeliveryDriver = AllOrders.Item(1).DriverName
        End If
        Exit Function
      End If
    End If
  End If
End Function
Public Function CashDue(TipAmount As Currency) As Currency
Dim Db          As ADODB.Connection '
Dim Tb          As ADODB.Recordset
    TipAmount = 0
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select * from [Payments]"
    Sql = Sql & " Where [Payment type] = 'Cash' and [close out day] = 0"
    ConnectADORS Sql, Db, Tb
    Do While Not Tb.EOF
      CashDue = CashDue + Tb![Payment Amount]
      Tb.MoveNext
    Loop
    SetNothing Tb
    Sql = "Select * from [Payments] Where [Close out day] = 0"
    ConnectADORS Sql, Db, Tb
    Do While Not Tb.EOF
      TipAmount = TipAmount + Tb![Tip Amount]
      Tb.MoveNext
    Loop
    DbClose Db, Tb
CashDue = CashDue - TipAmount
End Function
Public Function IsCheckOpenClosed(CurrentCheck As Long) As Boolean
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Function
Sql = "Select * from [Open Checks] where [Close out Day] = 0 and [Check Number] = " & CurrentCheck
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
Tb.Open Sql, Db, adOpenStatic, adLockReadOnly
If Tb.EOF Then
 IsCheckOpenClosed = True
Else
 IsCheckOpenClosed = False
End If
DbClose Db, Tb
End Function
Public Function TableisOpen(TableNo As String) As Boolean
TableisOpen = False
If UCase(GetConfigInfo("Full Service", "Disable Table Check", "ON", "FUNCTIONS.INI")) = "DISABLED" Then Exit Function

Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
On Error GoTo ErrorHandler
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
Sql = "Select * from [Open Checks] where [Close out Day] = 0 and [Table Number] = '" & TableNo & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
  If Tb.EOF Then
     TableisOpen = False
  Else
     TableisOpen = True
  End If
DbClose Db, Tb
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Table Is Open " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function SvrDeposit(SvrName As String) As Currency
Dim GA As Integer, ff As Integer
ff = FreeFile
SvrDeposit = 0: GA = 0
Open SvrPath.TerminalPath & "DAT Files\" & "CashDeposits.Dat" For Random Shared As ff Len = Len(Sdeposit)
     Do
       GA = GA + 1
       Get #ff, GA, Sdeposit
           If Trim(SvrName) = Trim(Sdeposit.SvrName) Then
              SvrDeposit = SvrDeposit + Sdeposit.DepositAmt
           End If
     Loop Until GA >= LOF(ff) / Len(Sdeposit)
Close ff
End Function
Public Sub ResetSVRDep(SvrName As String)
On Error Resume Next
Dim GA As Integer, ff As Integer
ff = FreeFile
GA = 0
Open SvrPath.TerminalPath & "DAT Files\" & "CashDeposits.Dat" For Random Shared As ff Len = Len(Sdeposit)
     Do
       GA = GA + 1
           Get #ff, GA, Sdeposit
           If Trim(SvrName) = Trim(Sdeposit.SvrName) Then
               Sdeposit.DepositAmt = 0
               Sdeposit.SvrName = ""
           Put #ff, GA, Sdeposit
           End If
     Loop Until GA >= LOF(ff) / Len(Sdeposit)
Close ff
End Sub
Public Function ConvertTimeToMinutes(InMins As Single) As Single
      ConvertTimeToMinutes = 0
      Dim TTl_Hrs           As Single
      Dim THours_Worked     As Single
      Dim TTlMins           As Single
      Dim Mins_Worked       As Single
      Dim MinWork           As Single
      TTl_Hrs = InMins
      If TTl_Hrs < 0 Then TTl_Hrs = TTl_Hrs + 1440
      THours_Worked = Int(TTl_Hrs / 60)
      TTlMins = THours_Worked * 60
      Mins_Worked = TTl_Hrs - TTlMins
      MinWork = Mins_Worked / 60
      ConvertTimeToMinutes = Format(THours_Worked + MinWork, "0.00")
End Function
Function PlayWAV(FileName As String) As Boolean
   'If PlaySounds <> "ENABLED" Then Exit Function
   'If AllLanguage Is Nothing Then Exit Function
   'If Not AllLanguage.Item(FileName) Is Nothing Then
   '  PlayWAV = PlaySound(AllLanguage.Item(FileName).Sounds, 0, SND_ASYNC)
   'Else
    'If Dir(SvrPath.LocalTerminalPath & "DAT Files\" & "Lost Waves.Dat") > "" Then
    ' Dim FF As Integer
    ' FF = FreeFile
    ' Open SvrPath.LocalTerminalPath & "DAT Files\" & "Lost Waves.Dat" For Append As FF
    ' Write #FF, CStr(FileName)
    ' Close FF
    'End If
   'End If
End Function
Public Function RemotePrinterFailed(PrinterName As String, PrinterPort As String, FailedPrint As String, FailureMsg As String) As Boolean
On Error GoTo ERH
RemotePrinterFailed = False
If UCase(Input_Option(33)) = "OFF" Or UCase(Input_Option(33)) = "NO" Then Exit Function

If UCase(Input_Option(33)) <> "NONE" And UCase(Input_Option(33)) <> "ON" Then
 If MyRegisterName = Input_Option(33) Then
  If FileCheck(SvrPath.LocalTerminalPath & PrinterName & ".OFF") = True Then
    RemotePrinterFailed = True
  Else
    RemotePrinterFailed = False
  End If
 Else
  If FileCheck("\\" & Input_Option(33) & "\" & ReturnPOSName & "\" & PrinterName & ".OFF") = True Then
    RemotePrinterFailed = True
  Else
    RemotePrinterFailed = False
  End If
 End If
Else
'Turn Off Remote Printer Error Checking
Dim lf As Integer
Dim RemotePrinterName As String
Dim IsReadyResults    As String
 lf = InStr(3, PrinterPort, "\")
 RemotePrinterName = Mid(PrinterPort, lf + 1, Len(PrinterPort))
 If InStr(PrinterPort, "\") > 0 Then
      RemotePrinterFailed = Check_Remote(PrinterName, PrinterPort)
 Else
   Set IsReadytoPrint = New PrinterCheck.PrinterTest
   'Dim MYPrinters As SystemPrinters
   IsReadyResults = IsReadytoPrint.TryPrinter(PrinterName, PrinterPort, MyRegisterName)
   If IsReadyResults = "Ready" Or IsReadyResults = "Not Checked" Then
      RemotePrinterFailed = False
   Else
      RemotePrinterFailed = True
   End If
   Set IsReadytoPrint = Nothing
 End If
End If
 If RemotePrinterFailed = True And FailedPrint = "None" Then
    See_Msg "Printer Failure", 5, " Print Job Is Waiting for Printer"
    RemotePrinterFailed = False
 ElseIf RemotePrinterFailed = True Then
    Call See_Msg("Printer Failure", 5, " Sending Print Job To " & FailedPrint)
    FailureMsg = PrinterName
    PrinterPort = GetSpoolPrinter(FailedPrint)
    RemotePrinterFailed = False
 End If
 Exit Function
ERH:

Err.Clear
End Function
Public Function CloseDateFound(DateSelected As String) As Boolean
On Error GoTo ERH
If IsDate(DateSelected) = False Then
    CloseDateFound = True
    Exit Function
End If
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Db) = False Then Exit Function
' Sql = "Delete from [Sales] Where IsNull([date]) or Not IsDate([Date])"
' DB.Execute Sql
 Sql = "Select * from [Sales] Where TRY_CAST([date] AS datetime) = '" & Format(CDate(DateSelected), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, Tb
If Not Tb.EOF Then
   CloseDateFound = True
   Trans_Log "Duplicate Closing Date Found"
Else
   CloseDateFound = False
End If
DbClose Db, Tb
Exit Function
ERH:
Err.Clear
End Function
Public Function GetFirstDayOrWeek() As String
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim StartDay As String
Dim EndDay   As String
GetFirstDayOrWeek = "Not Found"
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Select *  from [Closed Checks] Where [Close Out Day] = 0"
  Sql = Sql & "  Order by TRY_CAST([Order Date] AS datetime)+ TRY_CAST([Close Time] AS datetime) asc"
ConnectADORS Sql, Db, RS
If RS.EOF Then
   DbClose Db, RS
   Exit Function
Else
  'StartDay = RS.Fields("Order Date").Value & " " & RS.Fields("Close Time").Value
  StartDay = RS.fields("Order Date").Value & " " & GetStartTime
End If
RS.Close
  Sql = "Select *  from [Closed Checks] Where [Close Out Day] = 0"
  Sql = Sql & "  Order by TRY_CAST([Order Date] AS datetime)+ TRY_CAST([Close Time] AS datetime) desc"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
  EndDay = RS.fields("Order Date").Value & " " & RS.fields("Close Time").Value
End If
DbClose Db, RS
If DateDiff("n", StartDay, EndDay) > 1440 Then
  GetFirstDayOrWeek = "Close Needed"
End If
Exit Function
ErrorHandler:
GetFirstDayOrWeek = Err.Description
Err.Clear
End Function
Public Function RecordsCount(RS As Object) As Long
Dim RA As Long
If RS.RecordCount > -1 Then
   RecordsCount = RS.RecordCount
   Exit Function
Else
    If Not RS.BOF Then RS.MoveFirst
    RA = 0
    Do While Not RS.EOF
       RS.MoveNext
       RA = RA + 1
    Loop
    If Not RS.BOF Then RS.MoveFirst
End If
If RA < 0 Then RA = 0
RecordsCount = RA
End Function
Public Function GetServerSalesItem(Item As String) As Boolean
On Error GoTo ERH
GetServerSalesItem = False
Dim R As Integer
For R = 0 To UBound(Item_Search, 1)
   If Item = Item_Search(R) Then
      GetServerSalesItem = True
      Exit Function
   End If
Next R
ERH:
End Function
Public Function ReturnDate(RS As Object) As String
If Not RS.BOF Then RS.MoveFirst
Do While Not RS.EOF
   ReturnDate = RS(1)
  RS.MoveNext
Loop
If Not RS.BOF Then RS.MoveFirst
End Function
Public Function GetFirstMonday() As String
Dim WeekDay As Integer
Dim NewYear As Date
NewYear = DateSerial(Year(Date), 1, 1)
WeekDay = (NewYear - 2) Mod 7
If WeekDay < 4 Then
  GetFirstMonday = CStr(NewYear - WeekDay)
Else
  GetFirstMonday = CStr(NewYear - WeekDay + 7 + (WeekStartsOn - 2))
End If
End Function
Public Function ReturnWTDPeriod(InDate As String) As String
On Error Resume Next
If Not IsDate(InDate) Then Exit Function
Dim DayNo     As Integer
Dim DaysToGet    As Integer
'// Start Here
Dim DaysInPeriod As Integer
Dim PeriodStart  As String
Dim P_Days       As Integer
Dim DayPeriod    As Integer
PeriodStart = GetFirstMonday
DayPeriod = DateDiff("d", PeriodStart, InDate, vbUseSystemDayOfWeek, vbUseSystem)
DaysInPeriod = DayPeriod / 28
Dim dS As String
dS = CStr(CDate(PeriodStart) + ((DaysInPeriod - 1) * 28))
P_Days = (DateDiff("d", CDate(dS), CDate(InDate), vbUseSystemDayOfWeek, vbUseSystem) \ 7)
PeriodStart = DateAdd("ww", P_Days, CDate(dS))
'// end here
'DayNo = DatePart("w", CDate(InDate), WeekStartsOn, 3)
If CDate(PeriodStart) = CDate(InDate) Then
 ReturnWTDPeriod = CStr(InDate)
Else
 ReturnWTDPeriod = PeriodStart
End If
End Function
Public Function ReturnPeriodToDate(InDate As String) As String
On Error Resume Next
If Not IsDate(InDate) Then Exit Function
Dim DayNo        As Integer
Dim DayPeriod    As Integer
Dim DaysToGet    As Integer
Dim DaysInPeriod As Integer
Dim PeriodStart  As String
PeriodStart = GetFirstMonday
DayPeriod = DateDiff("d", PeriodStart, InDate, vbUseSystemDayOfWeek, vbUseSystem)
DaysInPeriod = DayPeriod / 28
Dim dS As String
dS = CStr(CDate(PeriodStart) + ((DaysInPeriod - 1) * 28))
If DateAdd("d", CDate(dS), CDate(InDate)) = 0 Then
 ReturnPeriodToDate = InDate
Else
 ReturnPeriodToDate = dS
End If
End Function
Public Function SortDrink(IncomingDrink As String) As Long
SortDrink = 0
If IncomingDrink = "" Then Exit Function
On Error Resume Next
Dim PIF As Integer

If InStr(IncomingDrink, Chr(2)) > 0 And InStr(IncomingDrink, Chr(3)) > 0 Then
   If InStr(IncomingDrink, "i") > 0 Then  'AzBar string
     SortDrink = Val(Mid(IncomingDrink, InStr(IncomingDrink, "i") + 1, 4)) & Val(Mid(IncomingDrink, InStr(IncomingDrink, "s") + 1, 1))
   Else
     PIF = GetBergLen
     If PIF > 0 Then  '' strip Null Value
      SortDrink = Val(Mid$(IncomingDrink, 2, PIF))
     Else
      SortDrink = Val(Mid$(IncomingDrink, 2, Len(IncomingDrink) - 3))
     End If
   End If
Else
   SortDrink = 0
End If
Trans_Log "Berg Sort Drint Plu data in " & IncomingDrink & " PLU Return for look up " & SortDrink
End Function  'Berg Speed Pour
Public Function ReturnModOption(ModOp As Integer) As Boolean
 Select Case ModOp
  Case Is = 0
    If Val(GetMenu(37) & "") = 0 Or Val(GetMenu(38) & "") = 0 Or Val(GetMenu(39) & "") = 0 Or Val(GetMenu(40) & "") = 0 Or Val(GetMenu(41) & "") = 0 Or Val(GetMenu(46) & "") = 0 Then
       ReturnModOption = True
    Else
       ReturnModOption = False
    End If
  Case Else
    If Val(GetMenu(37) & "") = 1 Or Val(GetMenu(38) & "") = 1 Or Val(GetMenu(39) & "") = 1 Or Val(GetMenu(40) & "") = 1 Or Val(GetMenu(41) & "") = 1 Or Val(GetMenu(46) & "") = 1 Then
       ReturnModOption = True
    Else
       ReturnModOption = False
    End If
  End Select
End Function
Public Function LoadSpeedTender(Idx As Integer, CheckTotal As Currency) As String
    If Val(CheckTotal) > 32000 Then
     LoadSpeedTender = STenders(Idx).TenderName
     Exit Function
    End If
    If FileCheck(SvrPath.LocalTerminalPath & "INI Files\" & "SpeedTender.INI") = False Then
        If CInt(CheckTotal) < 5 Then
         STenders(0).TenderAmt = 5: STenders(1).TenderAmt = 10: STenders(2).TenderAmt = 20
         STenders(0).TenderName = "5.00": STenders(1).TenderName = "10.00": STenders(2).TenderName = "20.00"
        ElseIf CInt(CheckTotal) > 5 And CInt(CheckTotal) <= 20 Then
         STenders(0).TenderAmt = 10: STenders(1).TenderAmt = 15: STenders(2).TenderAmt = 20
         STenders(0).TenderName = "10.00": STenders(1).TenderName = "15.00": STenders(2).TenderName = "20.00"
        ElseIf CInt(CheckTotal) > 20 And CInt(CheckTotal) <= 30 Then
         STenders(0).TenderAmt = 20: STenders(1).TenderAmt = 25: STenders(2).TenderAmt = 30
         STenders(0).TenderName = "20.00": STenders(1).TenderName = "25.00": STenders(2).TenderName = "30.00"
        ElseIf CInt(CheckTotal) > 30 And CInt(CheckTotal) <= 50 Then
         STenders(0).TenderAmt = 10: STenders(1).TenderAmt = 20: STenders(2).TenderAmt = 50
         STenders(0).TenderName = "10.00": STenders(1).TenderName = "20.00": STenders(2).TenderName = "50.00"
        Else
         STenders(0).TenderAmt = 20: STenders(1).TenderAmt = 50: STenders(2).TenderAmt = 100
         STenders(0).TenderName = "20.00": STenders(1).TenderName = "50.00": STenders(2).TenderName = "100.00"
        End If
    End If
    LoadSpeedTender = STenders(Idx).TenderName
    '// tender amount change based on check amount.
End Function
Public Function GetPrefixMultiplyer(PreFixNameIn As String) As Currency
    Dim Db       As ADODB.Connection
    Dim RS       As ADODB.Recordset
    Dim PFLookUp As String
    GetPrefixMultiplyer = 0
    If Right$(PreFixNameIn, 1) <> "-" Then
     PFLookUp = PreFixNameIn
    Else
     PFLookUp = Left$(PreFixNameIn, Len(PreFixNameIn) - 1)
    End If
    If ConnectADODB(SystemOptionsDB, Db) = False Then Exit Function
    Sql = "Select * from [Modifier Prefix]  where [PreFix Name] = '" & PFLookUp & "'"
    If ConnectADORS(Sql, Db, RS) = False Then Exit Function
    If Not RS.EOF Then
       GetPrefixMultiplyer = Val(RS.fields("Upcharge").Value & "")
    End If
    DbClose Db, RS
End Function
Public Function GetDelivery(Address As String, City As String, State As String, Zip As String) As DrvZoneData
Dim DelLoc      As String
Dim RestLoc     As String
Dim Directions  As String
On Error GoTo ERH
 GetDelivery.ZoneName = ""
 GetDelivery.ZoneDistance = 0
 GetDelivery.Instruction = ""
 If Check_Option(59) = 0 Then Exit Function
 If SiteLic.Google = 0 Then
  See_Msg "No License for Directions", 5, "Please Contact Your Dealer For Help"
  Exit Function
 End If
  Dim GetDir As MSaleDirections.GetDirections
  Set GetDir = New MSaleDirections.GetDirections
  Dim StartingFrom   As String
  Dim EndAt          As String
  Dim DriveResults() As String
  Dim EP             As Integer
  Dim DrvResults     As DrvZoneData
  If GetConfigInfo("Phone Orders", "Alt Address", "", "FUNCTIONS.INI") > "" Then
    StartingFrom = GetConfigInfo("Phone Orders", "Alt Address", "", "FUNCTIONS.INI") & " " & GetConfigInfo("Phone Orders", "Alt City", "", "FUNCTIONS.INI") & "," & GetConfigInfo("Phone Orders", "Alt State", "", "FUNCTIONS.INI") & " " & GetConfigInfo("Phone Orders", "Alt Zip", "", "FUNCTIONS.INI")
  Else
    StartingFrom = StoreAddress & " " & StoreCity & "," & LicInfo.Lic_State & " " & LicInfo.Lic_Zip
  End If
  If City = "" Then City = StoreCity
  If State = "" Then State = LicInfo.Lic_State
  If Zip = "" Then Zip = LicInfo.Lic_Zip
  If InStr(Address, "--") > 0 Then
    EndAt = Trim$(Left$(Address, InStr(Address, "--") - 1)) & " " & City & "," & State & " " & Zip
  Else
    EndAt = Address & " " & City & "," & State & " " & Zip
  End If
  Directions = GetDir.GetDirections(StartingFrom, EndAt)
  If InStr(Directions, "|") = 0 Then
    See_Msg Directions, 0, "Problem"
    Set GetDir = Nothing
    GetDelivery.ZoneName = ""
    GetDelivery.ZoneDistance = 0
    GetDelivery.Instruction = ""
    Exit Function
  End If
  DriveResults = Split(Directions, "|")
  If InStr(Directions, "|") > 0 Then
    GetZoneName DriveResults(0), DriveResults(1), DrvResults
  End If
  If DrvResults.ZoneName <> "Out of Area" Then
   If InStr(Directions, "|") > 0 Then
    GetDelivery.Instruction = ""
    For EP = 0 To 2
      GetDelivery.Instruction = GetDelivery.Instruction & DriveResults(EP) & vbCrLf
    Next EP
   End If
  End If
  GetDelivery.ZoneName = DrvResults.ZoneName
  GetDelivery.ZoneDistance = DrvResults.ZoneDistance
  Set GetDir = Nothing
  ChDir App.path
Exit Function
ERH:
Error_Log Err.Number, Err.Description, "Get Directions", "None"
End Function
Public Function DaysSinceClose() As Long
On Error GoTo ErrorHandler
Dim SDate As String
Dim EDate As String
Dim Cdb   As ADODB.Connection
Dim CRS   As ADODB.Recordset
DaysSinceClose = 0
''''''''''''''''''''''
ConnectADODB CheckStatDb, Cdb
''''''''''''''''''''
Sql = "Select * from [Closed Checks] order by TRY_CAST([order date] AS date) asc "
ConnectADORS Sql, Cdb, CRS
If Not CRS.BOF Then CRS.MoveFirst
 SDate = CRS![Order Date] & ""
If Not CRS.EOF Then CRS.MoveLast
 EDate = CRS![Order Date] & ""
DbClose Cdb, CRS
If IsDate(SDate) And IsDate(EDate) Then
  If DateDiff("d", SDate, EDate) > 7 Then
    'See_Msg "A Weekly Close Out Is Needed", 5, "Days Since Close " & DateDiff("d", SDate, EDate)
    DaysSinceClose = DateDiff("d", SDate, EDate)
  End If
End If
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function CheckIsEmpty() As Boolean
Dim N As Integer
CheckIsEmpty = False
For N = 0 To Number_In_Party - 1
 If NextRow(N) > 0 Then
  Exit Function
 End If
Next N
For N = 0 To Number_In_Party - 1
 If Order.Ticket(0).TextMatrix(N, 1) > "" Then
  Exit Function
 End If
Next N
CheckIsEmpty = True
End Function
Public Function See_Msg(First_Line As String, BoxType As Integer, Second_Line As String) As Integer
On Error GoTo ErrorHandler

Dim MyMsg As SeeMessage
Set MyMsg = New SeeMessage
See_Msg = MyMsg.ShowMsg(First_Line, Second_Line, BoxType, False)
Set MyMsg = Nothing
Exit Function
ErrorHandler:
Err.Clear
Set MyMsg = Nothing
End Function
Public Function EPServiceIsUp() As Boolean
On Error GoTo ERH
Dim SVRF As Boolean
Dim GL   As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
 If GL.GetKeyData("Sql Host") > "" Then
   SVRF = GL.IsSvrRunning(MyRegisterName, "ENTERPRISEWORKER")  'MSSQL$MSALE  'ENTERPRISEWORKER
 End If
Set GL = Nothing
Exit Function  '//Makes sure enterprise service is running
ERH:
Err.Clear
End Function
Public Function ResetFrequentDiner()
On Error Resume Next
 Dim FD As ADODB.Connection
 Dim DaysToKeep As Integer
 DaysToKeep = Val(GetConfigInfo("Data", "Keeps Sales Data Days", 1100, "Functions.INI"))
 If DaysToKeep = 0 Then Exit Function
 If ConnectADODB(SvrPath.TerminalPath & FrequentDinerDb, FD) = False Then Exit Function
  Sql = "Delete from [FrequentDinerDetail] where TRY_CAST([Date] AS date) < '" & CDate(Date - DaysToKeep) & "'"
  FD.Execute Sql
  Sql = "Delete from [Non Reset Data] where TRY_CAST([Date] AS date) < '" & CDate(Date - DaysToKeep) & "'"
  FD.Execute Sql
  See_Msg "Frequent Diner data purged", 5, ""
  FD.Close
 End Function
Public Function GetDeliveryChargesPaidOut(Driver As String) As Currency
Dim Db As ADODB.Connection
Dim OP As ADODB.Recordset
GetDeliveryChargesPaidOut = 0
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
  Sql = "Select * from [Deposits Paid In] where [Driver Name] = '" & Driver & "'"
ConnectADORS Sql, Db, OP
 If Not OP.EOF Then
  GetDeliveryChargesPaidOut = Val(OP![Delivery Charge Collected] & "")
 End If
 DbClose Db, OP
End Function
Public Function GetDriverCashDrop(Driver As String) As Currency
Dim Db As ADODB.Connection
Dim OP As ADODB.Recordset
GetDriverCashDrop = 0
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
  Sql = "Select * from [Deposits Paid In] where [Driver Name] = '" & Driver & "'"
ConnectADORS Sql, Db, OP
 If Not OP.EOF Then
  GetDriverCashDrop = Val(OP![Cash Deposit] & "")
 End If
 DbClose Db, OP
End Function
Public Function GetAdditionalCash(BartenderName As String, StartTime As String, StartDate As String) As Currency
On Error GoTo ErrorHandler
Dim CCDB As ADODB.Connection
Dim CCRS As ADODB.Recordset
GetAdditionalCash = 0
Dim EndDateTime As Date
EndDateTime = Date & " " & Time
   ConnectADODB SvrPath.TerminalPath & FinancialDb, CCDB
     Sql = "Select DISTINCT [Payment Type],"
     Sql = Sql & " Sum([Payment Amount]) as PCollected,"
     Sql = Sql & " Count([Payment Type]) as QtyAccepted,"
     Sql = Sql & " Sum([Tip Amount]) as TipAmt"
     Sql = Sql & " From DayPayments"
     Sql = Sql & " Where [Payment Type] = 'Cash'"    'WAS -9999
     Sql = Sql & " and [Server Name] <> '" & BartenderName & "'"
     Sql = Sql & " and [Cashier Name] = '" & BartenderName & "'"
     'If Check_Option(65) = 1 Then
     '  Sql = Sql & " and [Cashier Name] = 'Closed'"
     'End If
     If IsDate(StartTime) And IsDate(StartDate) Then
           Sql = Sql & " and TRY_CAST([Recorded date] AS datetime) + TRY_CAST([Recorded Time] AS datetime) between '" & CDate(StartDate) + CDate(StartTime) & "' and '" & EndDateTime & "'"
     End If
     Sql = Sql & " Group by [Payment Type]"
     Sql = Sql & " Order by [Payment Type] asc"
     ConnectADORS Sql, CCDB, CCRS
     Do While Not CCRS.EOF
        GetAdditionalCash = GetAdditionalCash + Val(CCRS!PCollected & "")
        CCRS.MoveNext
     Loop
    DbClose CCDB, CCRS
   
Exit Function
ErrorHandler:
Call See_Msg("Error " & Err, 5, Error)
'Resume
Exit Function
End Function
Public Function CashClosedByBartender(SvrsName As String, BegCheckNo As Long) As Currency
On Error GoTo ErrorHandler
Dim CCDB As ADODB.Connection
Dim CCRS As ADODB.Recordset
CashClosedByBartender = 0
  ConnectADODB SvrPath.TerminalPath & FinancialDb, CCDB
     Sql = "Select DISTINCT [Payment Type],"
     Sql = Sql & " Sum([Payment Amount]) as PCollected,"
     Sql = Sql & " Count([Payment Type]) as QtyAccepted,"
     Sql = Sql & " Sum([Tip Amount]) as TipAmt"
     Sql = Sql & " From DayPayments"
     Sql = Sql & " Where [Payment Type] = 'Cash'"    'WAS -9999
     Sql = Sql & " and [Server Name] = '" & SvrsName & "'"
     Sql = Sql & " and [Cashier Name] <> '" & SvrsName & "'"
     Sql = Sql & " and [Check Number] >= " & BegCheckNo
     'If Check_Option(65) = 1 Then
     ' Sql = Sql & " and [Cashier Name] = 'Closed'"
     'End If
     Sql = Sql & " Group by [Payment Type]"
     Sql = Sql & " Order by [Payment Type] asc"
     ConnectADORS Sql, CCDB, CCRS
     Do While Not CCRS.EOF
        CashClosedByBartender = CashClosedByBartender + Val(CCRS!PCollected & "")
        CCRS.MoveNext
     Loop
    DbClose CCDB, CCRS
   
Exit Function
ErrorHandler:
Call See_Msg("Error " & Err, 5, Error)
Exit Function
End Function
Public Function Get_AuditRefund(CashierServerName As String, RefundTender As RefundType) As RefundAmts
On Error GoTo ERH
Dim Db  As ADODB.Connection
Dim RS  As ADODB.Recordset
Get_AuditRefund.RefundAmt = 0: Get_AuditRefund.RefundTax = 0: Get_AuditRefund.RefundList = ""
    
    ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
    If RefundTender = AllCash Then
      Sql = "Select * from Refund Where [Close Out Day] = 0 and [Tender] = 'CASH'"
    Else
     If Check_Option(0) = 1 Then  'Quick Service
      If Check_Option(60) = 0 Then 'Drawer open for audit assignment
       Sql = "Select * from [Refund] Where [Cashier Name] = '" & CashierServerName & "'"
      Else
       Sql = "Select * from Refund Where [Cashier Name] <> 'Closed'"
      End If
      Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
     Else
        Sql = "Select * from Refund Where [Cashier Name] = '" & CashierServerName & "'"
     End If
     Sql = Sql & " and [Close Out Day] = 0"
     Select Case RefundTender
       Case Is = RefundType.Cash
           Sql = Sql & " and [Tender] = 'CASH'"
     End Select
   End If
    ConnectADORS Sql, Db, RS
      Do While Not RS.EOF
         Get_AuditRefund.RefundAmt = Get_AuditRefund.RefundAmt + Format(Val(RS![Amount] & ""), "0.00")
         Get_AuditRefund.RefundTax = Get_AuditRefund.RefundTax + Abs(RS![tax credit]) + Abs(Val(RS![SpcTaxCredit] & ""))
         
         Get_AuditRefund.RefundList = Get_AuditRefund.RefundList & PadString_LF(RS.fields("Tender").Value & "", 26) & Align(7, (Val(RS![Amount] & "") + Abs(RS![tax credit]) + Abs(Val(RS![SpcTaxCredit])) & "")) & vbCrLf
         Get_AuditRefund.RefundList = Get_AuditRefund.RefundList + RS.fields("Reason").Value & "" & vbCrLf
        RS.MoveNext
      Loop
    DbClose Db, RS
Exit Function
ERH:
Error_Log Err.Number, Err.Description, "Error In Refunds", ""
DbClose Db, RS
End Function
Public Function Get_RefundAmts(MyRefunds As FinancialItems, ReturnWhat As GetRefundWhat, TenderName As String) As Currency
Dim MyR As Integer
Get_RefundAmts = 0
TenderName = UCase(TenderName)
For MyR = 1 To MyRefunds.count
    Select Case ReturnWhat
      Case Is = GetRefundWhat.NetRefund
      
       If TenderName > "" Then
        If UCase(MyRefunds.Item(MyR).Tender) = TenderName Then
          Get_RefundAmts = Get_RefundAmts + MyRefunds.Item(MyR).Amount
        End If
       Else
         Get_RefundAmts = Get_RefundAmts + MyRefunds.Item(MyR).Amount
       End If
      
      Case Is = GetRefundWhat.TxRefund
       
       If TenderName > "" Then
        If UCase(MyRefunds.Item(MyR).Tender) = TenderName Then
          Get_RefundAmts = Get_RefundAmts + MyRefunds.Item(MyR).SalesTax
        End If
       Else
         Get_RefundAmts = Get_RefundAmts + MyRefunds.Item(MyR).SalesTax
       End If
      
      Case Is = GetRefundWhat.TotalRefund
       
       If TenderName > "" Then
        If UCase(MyRefunds.Item(MyR).Tender) = TenderName Then
          Get_RefundAmts = Get_RefundAmts + (MyRefunds.Item(MyR).Amount + MyRefunds.Item(MyR).SalesTax)
        End If
       Else
         Get_RefundAmts = Get_RefundAmts + (MyRefunds.Item(MyR).Amount + MyRefunds.Item(MyR).SalesTax)
       End If
     End Select
 Next MyR
End Function
Public Function GetShiftName(ShiftName As String, MenuName As String) As String
If Right(Trim$(ShiftName), 1) = "!" Then
  GetShiftName = MenuName & " " & Left(ShiftName, Len(ShiftName) - 2)
Else
  GetShiftName = ShiftName & MenuName
End If
End Function
Public Function AreAuditsOpen(SvrName As String) As Boolean
On Error GoTo ERH:
AreAuditsOpen = False
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Function
   If SvrName > "" Then
       Sql = "Select * from [Check Numbers] Where [Audit Status] = 0"
       Sql = Sql & " and [Servers Name] = '" & SvrName & "'"
   Else
      Sql = "Select * from [Check Numbers] Where [Audit Status] = 0"
   End If
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
  If Not RS.EOF Then
     AreAuditsOpen = True  'Audit is Open
  Else
     AreAuditsOpen = False   'Audit is Closed
  End If
DbClose Db, RS
Exit Function
ERH:
Err.Clear
AreAuditsOpen = False
End Function
Public Function GetPrefixDeptName(PFix As String, DeptName As String) As String
    If DeptName = "" Then DeptName = "Missing"
    If Check_Option(245) = 1 And Trim(PFix) > "" Then
        GetPrefixDeptName = PFix & " " & DeptName
    Else
        GetPrefixDeptName = DeptName
    End If
End Function
Public Function ReturnSentinal(KeyStroke As Integer) As Integer
 'If Sentinals = False Then
   Select Case KeyStroke
      Case Is = 63, 37, 59
          ReturnSentinal = 0
      Case Else
         ReturnSentinal = KeyStroke
   End Select
 'Else
 '  ReturnSentinal = KeyStroke
 'End If
End Function
Public Function ReturnNoSentinal(KeyStroke As String) As String
   Select Case KeyStroke
      Case Is = ";", "?", "%"
          ReturnNoSentinal = ""
      Case Else
         ReturnNoSentinal = KeyStroke
   End Select
End Function

Public Function SaleDestination(ByVal SaleType As String) As Integer
    Const DEST_DINEIN As Integer = 0
    Const DEST_GO     As Integer = 1

    Dim s As String
    s = Trim$(SaleType)

    ' Default for null/empty ? dine-in
    If LenB(s) = 0 Then
        SaleDestination = DEST_DINEIN
        Exit Function
    End If

    ' Keyword aliases (priority, like your original behavior)
    Select Case UCase$(s)
        Case "DINE IN", "DINE-IN", "HERE", "EAT IN", "EAT-IN"
            SaleDestination = DEST_DINEIN
            Exit Function
        Case "GO", "TO GO", "TO-GO", "TAKE OUT", "TAKE-OUT", "CARRY OUT", "CARRY-OUT", "TAKEAWAY"
            SaleDestination = DEST_GO
            Exit Function
    End Select

    ' Look for an exact destination match in OrderFor(), case-insensitive
    Dim i As Long
    For i = LBound(OrderFor, 1) To UBound(OrderFor, 1)
        If StrComp(s, Trim$(OrderFor(i).Destination), vbTextCompare) = 0 Then
            SaleDestination = i
            Exit Function
        End If
    Next i

    ' Fallback
    SaleDestination = DEST_DINEIN
End Function
Public Function TaxRefund(sPath As Variant) As Currency
On Error Resume Next
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB(sPath & FinancialDb, Db) = False Then Exit Function
TaxRefund = 0
Sql = "Select Distinct Sum([Tax Credit]+[SpcTaxCredit]) as TC from Refund "
ConnectADORS Sql, Db, RS
If Not RS.EOF Then TaxRefund = Val(RS!TC)
DbClose Db, RS
End Function
Public Function UpdateCashiers(CurCheck As Long)
On Error Resume Next
Dim UDb As ADODB.Connection
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, UDb) = False Then Exit Function
If Check_Option(213) = 1 Then
 Sql = "Update [Check Detail] set [Cashier Name] = '" & MyRegisterName & "' Where [Check Number] = " & CurCheck
Else
 Sql = "Update [Check Detail] set [Cashier Name] = '" & LoggedOnAs(0).EMPName & "' Where [Check Number] = " & CurCheck
End If
UDb.Execute Sql
SetNothing UDb
End Function
Public Function DayIsClosed() As Boolean
On Error Resume Next
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
'If Check_Option(65) = 1 Then  '24 hour operation
'    DayIsClosed = True
'    Exit Function
'End If
    If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
    Sql = "Select * from [check Detail] Where [Close out Day] = 0"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
      If Not Tb.EOF Then
         DayIsClosed = False
         DbClose Db, Tb
         Exit Function
      End If
 DayIsClosed = True
 DbClose Db, Tb
End Function
Public Function WeekIsClosed() As Boolean
On Error Resume Next
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
    If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
    Sql = "Select * from [check Detail] where [Close Out Day] <> 0"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
        If Not Tb.EOF Then
           WeekIsClosed = False
           DbClose Db, Tb
           Exit Function
        End If
    
 WeekIsClosed = True
DbClose Db, Tb
End Function
Public Function ChangeToRealCheckNumber(PhoneOrderNumber As Long) As Boolean
On Error GoTo ErrorHandler
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
    Dim Rst As String
    ChangeToRealCheckNumber = False
    Sql = "Update [Debit] Set [Check Number] = " & PhoneOrderNumber & ""  'cashier name must stay as phone order
    Sql = Sql & " where [Note] = '" & PhoneOrderNumber & "'"
    Sql = Sql & " and [Close Out Day] = 0 "
    Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & DebitDb))
    If Rst = "Success" Then
     ChangeToRealCheckNumber = True
    End If
Set GRS = Nothing
Exit Function
ErrorHandler:
ChangeToRealCheckNumber = False
Err.Clear
End Function

Public Function GetRealCheckNo(PhoneOrderNumber As Long) As Long
On Error GoTo ErrorHandler
  GetRealCheckNo = PhoneOrderNumber
  Dim Db As ADODB.Connection
  Dim Tb As ADODB.Recordset
  Dim Getcheck() As String
  Dim MyCheckNo  As Long
  MyCheckNo = -1
  OpenDb SvrPath.LocalTerminalPath & DebitDb, Db
    Sql = "Select * from [Debit] where [Note] = '" & CStr(PhoneOrderNumber) & "'"
  ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
    MyCheckNo = Tb.fields("Check Number").Value
  End If
  DbClose Db, Tb
  If MyCheckNo = -1 Then Exit Function
  OpenDb SvrPath.LocalTerminalPath & CheckStatDb, Db
    Sql = "Select * from [Closed Checks] where [Check Number] = " & MyCheckNo
  ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
    If InStr(Tb.fields("Table Number").Value, "-") > 0 Then
       Getcheck = Split(Tb.fields("Table Number").Value, "-")
       If Val(Getcheck(1)) > 0 Then
         If PhoneOrderNumber = Val(Getcheck(1)) Then
           GetRealCheckNo = Tb.fields("Check Number").Value
         End If
       End If
    End If
  End If
  DbClose Db, Tb
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function GetCheckGrat(GratName As String, MaxGrat As Currency) As Single
On Error GoTo ErrorHandler
  Dim DiscDb As ADODB.Connection
  Dim Tb As ADODB.Recordset
  OpenDb SvrPath.LocalTerminalPath & "MSaleData", DiscDb
   Sql = "Select * from [Discounts] where [Discount Name] = '" & GratName & "'"  ' and [Discount type] = 'Gratuity'"
  ConnectADORS Sql, DiscDb, Tb
  If Not Tb.EOF Then
     GetCheckGrat = Val(Tb![Discount Amount])
     MaxGrat = Val(Tb![Max Discount Amount] & "")
  End If
  DbClose DiscDb, Tb
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function CalculateSvrCharge() As Currency
 Dim SC   As Integer
 Dim SN   As Integer
 CalculateSvrCharge = 0
 For SN = 0 To Number_In_Party - 1
  For SC = 0 To NextRow(SN) - 1
    CalculateSvrCharge = CalculateSvrCharge + (CollectSvrCharge(CStr(Grid_array(SN, SC, 26)), Val(Grid_array(SN, SC, 2))) * Val(Grid_array(SN, SC, 0)))
    If Grid_array(SN, SC, 1) = "" Then Exit For
  Next SC
 Next SN
End Function
Public Function CollectSvrCharge(DepartmentName As String, ItemAmount As Currency) As Currency
On Error GoTo ErrorHandler
  CollectSvrCharge = 0
  Dim DiscDb As ADODB.Connection
  Dim Tb     As ADODB.Recordset
  OpenDb SvrPath.LocalTerminalPath & "MSaleData", DiscDb
   Sql = "Select * from [Discounts] where [Disc Department] = '" & DepartmentName & "' and [Type] = 'Service Charge'"
  ConnectADORS Sql, DiscDb, Tb
  If Not Tb.EOF Then
     If ItemAmount > Val(Tb.fields("Max Discount Amount").Value & "") Then
      CollectSvrCharge = Val(Tb![Discount Amount])
     End If '
  End If
  DbClose DiscDb, Tb
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Function PrintedChecks(CheckNumber As Long) As Long
On Error GoTo ErrorHandler
PrintedChecks = 0
If Val(Input_Option(77)) = 0 Then Exit Function
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
 If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
    Sql = "Select * from [Printed Receipts] Where [Check Number] = " & CheckNumber
 ConnectADORS Sql, Db, Tb
  If Tb.EOF Then
    Tb.AddNew
    PrintedChecks = 1
  Else
    PrintedChecks = Val(Tb![Printed]) + 1
  End If
    Tb![Check Number] = CheckNumber
    Tb![Printed] = PrintedChecks
   Tb.Update
 DbClose Db, Tb
Exit Function
ErrorHandler:
PrintedChecks = 0
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Printed Check ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function TransCounter(PathtoFile As String, CounterName As String, CouponNo As Long) As Long
On Error Resume Next
Dim ff As Integer  'make sure delete at end of day
ff = FreeFile
Dim GetChkNo As TransactionCount
Open PathtoFile & CounterName & "TransCounter.CTR" For Random Shared As ff Len = 8
  Get ff, 1, GetChkNo
   GetChkNo.TransNo = GetChkNo.TransNo + 1
   CouponNo = GetChkNo.TransCoupon
  Put ff, 1, GetChkNo
 TransCounter = GetChkNo.TransNo
 Close ff
End Function
Public Sub ZeroCounter(PathtoFile As String, CounterName As String, NxtCoupon As Long)
On Error Resume Next
Dim ff As Integer  'make sure delete at end of day
ff = FreeFile
Dim GetChkNo As TransactionCount
Open PathtoFile & CounterName & "TransCounter.CTR" For Random Shared As ff Len = 8
  Get ff, 1, GetChkNo
   GetChkNo.TransNo = 0
   GetChkNo.TransCoupon = NxtCoupon + 1
  Put ff, 1, GetChkNo
 Close ff
End Sub
Public Function GetDBaseName(FileAndPath As String) As String
Dim PS As Integer
Dim FB As Integer
Dim PN As Integer
GetDBaseName = ""
PN = 0
For PS = Len(FileAndPath) To 0 Step -1
  If Mid(FileAndPath, PS, 1) = "\" Then Exit For
  FB = PS
  PN = PN + 1
Next PS
If FB = 0 And PN = 0 Then Exit Function
GetDBaseName = Mid(FileAndPath, FB, PN)
End Function
Public Function Align(ByVal NOS As Integer, ByVal NumberToAlign As Currency) As String
Dim AlignPos As Integer
AlignPos = NOS - Len(Format(NumberToAlign, ".00"))
 If AlignPos < 0 Then AlignPos = 0
Align = String(AlignPos, 32) & Format(NumberToAlign, ".00")
End Function
Public Function ReturnChange(ChargeAmt As Currency) As Currency
    ReturnChange = CCur(Right(CStr(Format(ChargeAmt, "0.00")), 2)) / 100
End Function

Public Function Get_DayofWeek(ClosingDate) As String
On Error Resume Next
Dim dw As Byte
If ClosingDate = "" Then ClosingDate = Format(Date, "mm/dd/yyyy")
dw = DatePart("w", CDate(ClosingDate), vbMonday, 1)
   Select Case dw
        Case Is = 1
           Get_DayofWeek = "Monday "
        Case Is = 2
           Get_DayofWeek = "Tuesday "
        Case Is = 3
           Get_DayofWeek = "Wednesday "
        Case Is = 4
           Get_DayofWeek = "Thursday "
        Case Is = 5
           Get_DayofWeek = "Friday "
        Case Is = 6
           Get_DayofWeek = "Saturday "
        Case Is = 7
           Get_DayofWeek = "Sunday "
  End Select
End Function
Public Function TryDiscountNumber(CouponNo As String, CouponDesc As String, CouponAmt As Currency) As Boolean
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
TryDiscountNumber = False
If CouponNo = "" Then Exit Function
If ConnectADODB(SvrPath.TerminalPath & ProductHistory, Db) = False Then Exit Function
  'Sql = "Select * from [Coupon Tracking] where [Coupon Number] = '" & DiscountNo & "' and [Coupon Description] = '" & CellNo & "'"
  Sql = "Update [Coupon Tracking] Set [date] = '" & Now & "',[Coupon Number] = '" & CouponNo & "',[Coupon Description] = '" & CouponDesc & "'"
  Sql = Sql & ",[Coupon Amount] = " & CouponAmt
Db.Execute Sql
DbClose Db, RS
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Try Discounts ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function IsWholeNumber(ByVal v As Double) As Boolean
    IsWholeNumber = (Fix(v) = v)
End Function

Public Function PurgeGroupons() As Boolean
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
 If ConnectADODB(SvrPath.TerminalPath & ProductHistory, Db) = False Then Exit Function
  Sql = "Delete from [Coupon Tracking] where TRY_CAST([Date] AS date) < '" & Date - 120 & "'"
 Db.Execute ReturnDateString(Sql)
SetNothing Db
Exit Function
ErrorHandler:
Call Main_Error_Handler(Err.Number, Err.Description, "PurgeGroupons ")
Exit Function
End Function
Public Function Verify_Discount(Match_Name As String) As Boolean
Verify_Discount = False
Dim SF As Integer
If AssignedCoupon Is Nothing Then Exit Function
For SF = 1 To AssignedCoupon.count
   If Trim$(Match_Name) = Trim$(AssignedCoupon.Item(SF).MItem) Then
      Verify_Discount = True
      Exit Function
   End If
 Next SF
End Function
Public Function SetLicFile() As String
Dim FName As String
FName = Dir(DBFilePath & "*." & CStr(App.Major) & "x")  '& CStr(App.Minor)
If FName > "" Then
 SetLicFile = FName
Else
 SetLicFile = "Terminal." & App.Major & App.Minor & "x"
End If
SaveSetting "ProductID", "License File", "Terminal ID", SetLicFile
End Function
Public Function GetLicFile() As String
GetLicFile = GetSetting("ProductID", "License File", "Terminal ID")
If GetLicFile = "" Then
  GetLicFile = SetLicFile
End If
End Function
Public Function DateCode&(SD$)
Dim Mnth%
Dim Yr%
Dim Dy%
Mnth% = Month(CDate(SD$))
Yr% = Year(CDate(SD$))
Dy% = Day(CDate(SD$))
DateCode& = (Yr% * 1000000) + (Mnth% * 1000) + Dy%
End Function
Public Function QuarterRound(ByVal QtrMins As Long, PunchDate As String) As String
'added to stop 00:00 in database cannot convert.
PunchDate = Format(Date, "MM/DD/YYYY")
If Minute(Time) < 8 Then
    QuarterRound = Format(Hour(Time) & ":" & "00", "HH:MM ampm")
    QtrMins = 0
ElseIf Minute(Time) >= 8 And Minute(Time) < 23 Then
    QuarterRound = Format(Hour(Time) & ":" & "15", "HH:MM ampm")
    QtrMins = 15
ElseIf Minute(Time) >= 22 And Minute(Time) < 37 Then
    QuarterRound = Format(Hour(Time) & ":" & "30", "HH:MM ampm")
    QtrMins = 30
ElseIf Minute(Time) >= 36 And Minute(Time) < 52 Then
    QuarterRound = Format(Hour(Time) & ":" & "45", "HH:MM ampm")
    QtrMins = 45
ElseIf Minute(Time) >= 51 And Minute(Time) <= 59 Then
    If Hour(Time) + 1 = 24 Then
      QuarterRound = "12:00 am"
      PunchDate = DateAdd("d", 1, Date)
    Else
      QuarterRound = Format(Hour(Time) + 1 & ":00", "HH:mm ampm")
    End If
    QtrMins = 0
Else
    QuarterRound = Format(Time, "HH:MM ampm")
    QtrMins = 0
End If
End Function
Public Function RoundMinute(InMin As Long) As Long
If Check_Option(2) = 0 Then  'Returns minutes for pool table times
   RoundMinute = InMin
   Exit Function
Else
   Dim HRS As Integer
   Dim HM  As Integer
   If InMin > 60 Then
     HRS = Fix((InMin / 60)) * 60
     HM = InMin - HRS
   Else
     HM = InMin
     HRS = 0
   End If
   If HM < 8 Then
      RoundMinute = 0
   ElseIf HM >= 8 And HM < 23 Then
      RoundMinute = 15
   ElseIf HM >= 22 And HM < 37 Then
      RoundMinute = 30
   ElseIf HM >= 36 And HM < 52 Then
      RoundMinute = 45
   ElseIf HM >= 51 And HM <= 59 Then
      RoundMinute = 60
   Else
      RoundMinute = InMin
   End If
End If
RoundMinute = HRS + RoundMinute
End Function
Public Function RoundToNickel(ItemAmount As Currency) As Currency
Dim Chg As Single
Dim RAmt As Single
If ItemAmount > 0 Then ItemAmount = Int(ItemAmount * 100) / 100
Chg = (ItemAmount * 100) Mod 10
    Select Case Chg
         Case Is = 5
             RoundToNickel = Format(ItemAmount, "0.00")
         Case 0, 1, 2
             RoundToNickel = Format(ItemAmount, "0.0")
         Case 3, 4
             RAmt = (0.05 * 100) / 100
             RoundToNickel = Format(ItemAmount, "0.0") + RAmt
         Case 6, 7
             RAmt = ItemAmount * 10
             RAmt = Int(RAmt)
             RoundToNickel = Format(RAmt / 10, "0.00") + 0.05
         Case Else
             RAmt = ((Val(0.05) * 100 - Chg) / 100) ' - Val(CustOption.SaleRounding)
             RAmt = (Val(0.05)) + RAmt
             RoundToNickel = Format(ItemAmount, "0.00") + RAmt
    End Select
End Function
Public Function RoundToQuarter(ByVal CheckAmount As Currency) As Currency
On Error GoTo ERH
If CheckAmount = 0 Then
  RoundToQuarter = 0
  Exit Function
End If
Dim cents As Integer
Dim Change() As String
Change = Split(CheckAmount, ".")
cents = Val(Change(1))
If Val(cents) < 13 Then
    RoundToQuarter = Int(CheckAmount) & ".00"
ElseIf Val(cents) >= 13 And Val(cents) <= 37 Then
    RoundToQuarter = Int(CheckAmount) & ".25"
ElseIf Val(cents) >= 38 And Val(cents) <= 62 Then
    RoundToQuarter = Int(CheckAmount) & ".50"
ElseIf Val(cents) >= 63 And Val(cents) <= 87 Then
    RoundToQuarter = Int(CheckAmount) & ".75"
ElseIf Val(cents) >= 88 And Val(cents) <= 99 Then
    RoundToQuarter = Int(CheckAmount) + 1 & ".00"
Else
    RoundToQuarter = Format(CheckAmount, "0.00")
End If
Exit Function
ERH:
RoundToQuarter = Format(CheckAmount, "0.00")
End Function
Public Function ReturnTotalMins(ByVal InTime As String, ByVal OutTime As String) As Long
Dim MinRemaining As Long
Select Case DateDiff("n", CDate(InTime), CDate(OutTime))
  Case Is <= 60
    ReturnTotalMins = RoundMinute(DateDiff("n", CDate(InTime), CDate(OutTime)))
  Case Else
    MinRemaining = DateDiff("n", CDate(InTime), CDate(OutTime)) Mod Int((DateDiff("n", CDate(InTime), CDate(OutTime))) / 60) * 60
    ReturnTotalMins = (Int((DateDiff("n", CDate(InTime), CDate(OutTime))) / 60) * 60) + RoundMinute(MinRemaining)
End Select
End Function
Public Function GetAvgServingTime() As Long
On Error GoTo ErrorHandler
Dim BeginTime   As Date
Dim TotalOpen   As Long
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim C           As Integer
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
  Sql = "Select * from [Open Checks] Where [Close Out Day] = 0 order by [Check Number] asc"
  'adoCursorType = 3
ConnectADORS Sql, Db, Tb
C = -1: TotalOpen = 0
Do While Not Tb.EOF
    If Not IsDate(Tb(19)) Then
      BeginTime = Format(Date, "mm/dd/yyyy") & " " & Tb(4)
    Else
      BeginTime = Tb(19) & " " & Tb(4)
    End If
    C = C + 1
    TotalOpen = TotalOpen + ElapsedMinutes(BeginTime, Now)
  Tb.MoveNext
Loop
DbClose Db, Tb
If C > 0 And TotalOpen > 0 Then
  GetAvgServingTime = TotalOpen / 3
Else
  GetAvgServingTime = 0
End If
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Serving time ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function IsQSROn() As Boolean
If Input_Option(88) = "QSR" Then
 IsQSROn = True
Else
 IsQSROn = False
End If
End Function
Public Function Check_Null(ByVal VarName As Variant) As Variant
    On Error Resume Next
    ' Check for Null value and return an empty string if Null, or return the original value
    If IsNull(VarName) Then
        Check_Null = ""
    Else
        Check_Null = VarName
    End If
    ' Handle any unexpected errors
    If Err.Number <> 0 Then
        Check_Null = ""
        Err.Clear
    End If
End Function
Public Function ReturnProcessingFees(ByVal InDate As String, ByVal EDate As String) As Currency
    ReturnProcessingFees = 0
    If ProcessingFeesOn = False Then Exit Function
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
        Sql = "Select Distinct [Category Name], sum([Item Price]) as TExempt from [Detail] "
        Sql = Sql & " Where [Category Name] = 'Processing Fee'"
        Sql = Sql & " and TRY_CAST([Date] AS datetime) Between #" & CDate(InDate & " " & GetStartTime) & "#"
        Sql = Sql & " and #" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "#"
        Sql = Sql & " Group by [Category Name] "
    ConnectADORS ReturnDateString(Sql), Db, RS
    If Not RS.EOF Then
       ReturnProcessingFees = Format(RS.fields("TExempt").Value, "0.00")
    End If
    DbClose Db, RS
End Function
Public Function AverageTime() As String
    On Error GoTo ErrorHandler
    Dim Adb         As ADODB.Connection
    Dim ARs         As ADODB.Recordset
    Dim Sql         As String
    Dim trecords    As Long
    Dim rTime       As Long
    Dim TotalTime   As Long
    Dim avgTime     As Double

    ' Initialize the variables
    TotalTime = 0
    trecords = 0

    ' Connect to the database
    If Not ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Adb) Then
        AverageTime = "Error: Failed to connect."
        Exit Function
    End If

    Sql = "SELECT * FROM [Closed Checks] WHERE [Close out day] = 0"
    ConnectADORS Sql, Adb, ARs

    ' Process each record
    Do While Not ARs.EOF
        If ARs![Open Time] <> ARs![Close Time] Then
            TotalTime = TotalTime + DateDiff("n", CDate(ARs![Open Time]), CDate(ARs![Close Time]))
            trecords = trecords + 1
        End If
        ARs.MoveNext
    Loop

    ' Calculate average if applicable
    If trecords > 0 Then
        avgTime = TotalTime / trecords
        AverageTime = Format$(avgTime, "0.0")
    Else
        AverageTime = "0"
    End If
    ' Clean up
    GoTo CleanUp
ErrorHandler:
    AverageTime = "Error: " & Err.Description

CleanUp:
    DbClose Adb, ARs
End Function
Public Function GetActiveCategory(ByVal CategoryEvent As String) As Boolean
On Error Resume Next
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim FR As ADODB.Recordset
If Input_Option(51) <> "Yes" Then
 If CategoryEvent = "" Or CategoryEvent = "None" Then
  GetActiveCategory = True
  Exit Function
 End If
End If
GetActiveCategory = False
If Not ConnectADODB(SvrPath.LocalTerminalPath & "CRMENU.MDB", Db) Then
    Exit Function
End If
 Sql = "Select * from [Meal Period] where [Period Name] = '" & CategoryEvent & "'"
ConnectADORS Sql, Db, RS
 If Not RS.EOF Then
   If isValid(RS![Period time] & "") Then
    GetActiveCategory = True
    DbClose Db, RS
    Exit Function
   End If
 End If
DbClose Db, RS
End Function
Public Function GetActivePeriod() As String
On Error Resume Next
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim FR As ADODB.Recordset
GetActivePeriod = "All"
'ActivePeriod = "All"
If Check_Option(5) = 0 Then Exit Function
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
 Sql = "Select * from [Meal Period] "
OpenTb Sql, RS, Db
 Do While Not RS.EOF
  If isValid(RS![Period time] & "") Then
   GetActivePeriod = RS![Period Name] & ""
   DbClose Db, RS
   Exit Function
  End If
  RS.MoveNext
 Loop
DbClose Db, RS
GetActivePeriod = "None"
'ActivePeriod = "None"
End Function
Public Function isValid(EventName As String) As Boolean
    Dim Lk As Integer
    Dim trimmedEventName As String

    ' Normalize input for case-insensitivity and remove surrounding whitespace
    trimmedEventName = Trim$(UCase(EventName))

    ' If the event name is "NONE," it's automatically considered valid
    If trimmedEventName = "NONE" Then
        isValid = True
        Exit Function
    End If

    ' Loop through the list of events to find a matching event name
    For Lk = 0 To Total_Events
        If trimmedEventName = Trim$(UCase(Time_Events(Lk).EName)) Then
            ' Check if the event is within a valid date and time range
            If CheckDate(Lk) And IsValidTime(Time, Time_Events(Lk).StartTime, Time_Events(Lk).EndTime, Lk) Then
                isValid = True
                Exit Function
            End If
        End If
    Next Lk

    ' Default to returning False if no match found or not in a valid date/time range
    isValid = False
End Function
Public Function TimeByServer(ServerData() As ServerTimeData, ByServer As String) As Integer
On Error Resume Next
Dim Adb         As ADODB.Connection
Dim ARs         As ADODB.Recordset
Dim SRS         As ADODB.Recordset
Dim Reg         As Integer
Dim trecords    As Long
Dim rTime       As Long
Dim PrvRecords  As Long
Dim CurServer   As String
Dim NSrv        As Integer
ReDim ServerData(0)
rTime = 0: trecords = 0: PrvRecords = 0
NSrv = -1
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Adb) = False Then Exit Function
  If Trim$(ByServer) = "" Then
     Sql = "Select Distinct [Server Name] from [Closed Checks] Where [Close Out Day] = 0 "
  Else
     Sql = "Select * [Server Name] from [Closed Checks] Where [Close Out Day] = 0 and [Server Name] = '" & ByServer & "'"
  End If
  ConnectADORS Sql, Adb, SRS
   If Not SRS.EOF Then
      ReDim Preserve ServerData(RecordsCount(SRS) + PrvRecords)
      PrvRecords = PrvRecords + RecordsCount(SRS)
      If Not SRS.EOF Then
        Do While Not SRS.EOF
          CurServer = SRS![Server Name]
          Sql = "Select * from [Closed Checks] Where [Close out day] = 0 and [Server Name] = '" & CurServer & "'"
          ConnectADORS Sql, Adb, ARs
          Do While Not ARs.EOF
             If ARs![Open Time] <> ARs![Close Time] Then
               rTime = rTime + DateDiff("n", CDate(ARs![Open Time]), CDate(ARs![Close Time]), 2, vbFirstFullWeek)
               trecords = trecords + 1
             End If
             ARs.MoveNext
          Loop
          If Abs(rTime) > 0 And trecords > 0 Then
             NSrv = NSrv + 1
             ServerData(NSrv).Name = CurServer
             ServerData(NSrv).ETime = Format(rTime / trecords, "0.0")
             ServerData(NSrv).ERcd = Format(trecords, "0")
             rTime = 0: trecords = 0
          End If
          SRS.MoveNext
         Loop
        End If
    End If
    TimeByServer = NSrv
    DbClose Adb, SRS
End Function
Public Function CalChargeBack(CardName As String, ChargeAmount As Currency) As Currency
On Error GoTo ERH
CalChargeBack = 0
Dim cb   As ADODB.Connection
Dim CbTb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, cb
   Sql = "Select * From [Server Charge Back] where [Card name] = '" & CardName & "'"
ConnectADORS Sql, cb, CbTb
  If Not CbTb.EOF Then
    CalChargeBack = CbTb![Charge Perc] * ChargeAmount
  End If
DbClose cb, CbTb
Set cb = Nothing
Exit Function
ERH:
Trans_Log Err.Description & " Tip Fee"
End Function
Public Function ProcessItemName(ByVal m_ItemName As String, ByRef HasSpecialPrefix As Boolean) As String
    Dim firstChar As String
    Dim trimmedName As String

    trimmedName = Trim$(m_ItemName)
    firstChar = Left$(trimmedName, 1)

    Select Case firstChar
        Case "*"
            ' Remove the asterisk and update the flag
            ProcessItemName = Trim$(Mid$(trimmedName, 2))
            HasSpecialPrefix = True

        Case "v"
            ' Special handling when the first character is 'v'
            If Check_Option(99) = 0 Then
                ProcessItemName = Trim$(Mid$(trimmedName, 8))
                HasSpecialPrefix = True
            Else
                ProcessItemName = trimmedName
                HasSpecialPrefix = False
            End If

        Case Else
            ' Return the name as is if it does not start with an asterisk or 'v'
            ProcessItemName = trimmedName
            HasSpecialPrefix = False
    End Select
End Function

Public Function GetDiscountTax(DiscountName As String) As CouponTaxes
  Dim DiscDb As ADODB.Connection
  Dim Tb     As ADODB.Recordset
  GetDiscountTax.Tax1 = 0
  GetDiscountTax.Tax2 = 0
  ConnectADODB SvrPath.LocalTerminalPath & "MSaleData", DiscDb
     Sql = "Select * from [Discounts] where [Discount Name] = '" & DiscountName & "'"
  ConnectADORS Sql, DiscDb, Tb
   If Not Tb.EOF Then
      GetDiscountTax.Tax1 = Nz(Tb.fields("Tax 1").Value)
      GetDiscountTax.Tax2 = Nz(Tb.fields("Tax 2").Value)
   End If
   DbClose DiscDb, Tb
End Function
Public Function Check_Tipped(Worked_Where As String, FName As String, LName As String) As Integer
On Error Resume Next
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & "Employee.mdb", Db
Sql = "Select * from [Job Codes] Where [Job Code] ='" & Worked_Where & "'"
Sql = Sql & " and [First Name] = '" & FName & "' and [Last Name] = '" & LName & "'"
ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     Check_Tipped = Abs(Val(Tb![tip] & ""))
  Else
     Check_Tipped = 0
  End If
'if no job title found set to not tipped
DbClose Db, Tb
End Function
Public Function StartAudit(Worked_Where As String) As Integer
On Error Resume Next
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
Sql = "Select top 1 * from [Job Titles] Where [Job Title] ='" & Worked_Where & "'"
ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     StartAudit = Abs(Val(Tb![Tipped Position] & ""))
  Else
     StartAudit = 0
  End If
'if no job title found set to not tipped
DbClose Db, Tb
End Function
Public Function OpensCashDrawer(Worked_Where As String) As Integer
On Error GoTo ERH
OpensCashDrawer = 0
'If Check_Option(251) = 0 Then Exit Function
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
Sql = "Select * from [Job Titles] Where [Job Title] ='" & Worked_Where & "'"
ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     OpensCashDrawer = Abs(Val(Tb![Drawer] & ""))
  Else
     OpensCashDrawer = 0
  End If
'if no job title found set to not tipped
DbClose Db, Tb
Exit Function
ERH:
OpensCashDrawer = 0
DbClose Db, Tb
End Function
Public Function ReturnJobCode(JobName As String) As Integer
On Error Resume Next
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
  Sql = "Select * from [Job Titles] Where [Job Title] ='" & JobName & "'"
ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     ReturnJobCode = Val(Tb![Job Code Number] & "")
  Else
     ReturnJobCode = 1
  End If
'if no job title found set to not tipped
DbClose Db, Tb
End Function
Public Function CheckNoIsValid(CheckNo As Long) As Boolean
   Dim GL As GeneralLib.Utilities
   Set GL = New GeneralLib.Utilities
     CheckNoIsValid = GL.IsCheckNoNotInUse(CheckNo)
   Set GL = Nothing
End Function
Public Function CheckGratuity(CheckNo As Long) As Boolean
On Error GoTo ErrorHandler
   Dim Db As ADODB.Connection
   Dim RS As ADODB.Recordset
   If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
   Sql = "Select * from [Check Detail] Where [Close Out Day] = 0 and [Cateory Name] = 'Gratuity' and [Check Number] = " & CheckNo
   ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
      CheckGratuity = True
   Else
      CheckGratuity = False
   End If
   DbClose Db, RS
   Exit Function
ErrorHandler:
End Function
Public Function SeatHasGratuity(SeatNo As Integer) As Boolean
SeatHasGratuity = False
Dim FG As Integer
For FG = 0 To NextRow(SeatNo) - 1
 If Grid_array(SeatNo, FG, 23) = "Gratuity" Then
   SeatHasGratuity = True
   Exit Function
 End If
Next FG
End Function
Public Function HasDiscount(ItemToVoid As String, SeatNo As Integer, SelectNo As Integer) As Boolean
On Error GoTo ErrorHandler
   Dim Db As ADODB.Connection
   Dim RS As ADODB.Recordset
   If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Function
   Sql = "Select * from [Check Detail] Where [Close Out Day] = 0 and [Selection Number] <> " & SelectNo & " and [Disc Type] = '" & ProcessItemName(ItemToVoid, False) & "' and [Voided] = 0 and [Seat Number] = " & SeatNo & " and [Check Number] = " & Check_no
   ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
      HasDiscount = True
   Else
      HasDiscount = False
   End If
   DbClose Db, RS
   Exit Function
ErrorHandler:
HasDiscount = False
End Function
Public Function GetPlanName(MItem As String) As String
Dim Sp As Integer
On Error GoTo ErrorHandler
GetPlanName = ""
If MPAllItems Is Nothing Then Exit Function
If MPAllItems.Item(CStr(AllMenuItems(MItem).MenuItemName)) Is Nothing Then
   If MPAllItems.Item(CStr(AllMenuItems(MItem).DepartmentName)) Is Nothing Then
     If MPAllItems.Item(CStr(AllMenuItems(MItem).ReportCategory)) Is Nothing Then
        GetPlanName = ""
        Exit Function
     Else
        GetPlanName = AllMenuItems(MItem).ReportCategory
        Exit Function
     End If
   Else
        GetPlanName = AllMenuItems(MItem).DepartmentName
        Exit Function
   End If
Else
        GetPlanName = AllMenuItems(MItem).MenuItemName
   Exit Function
End If
Exit Function
ErrorHandler:
If Err = 91 Then Resume Next
End Function
Public Function PlanDiscountTake(CurrentPlan As String) As Boolean
Dim La As Integer
Dim LS As Integer
Dim ItemName As String
PlanDiscountTake = False
   For La = 0 To Number_In_Party - 1
      For LS = 0 To NextRow(La) - 1
         If Val(Grid_array(La, LS, 2)) < 0 Then
           ItemName = Replace(CStr(Grid_array(La, LS, 1)), "%", "")
         Else
           ItemName = CStr(Grid_array(La, LS, 1))
         End If
         If ProcessItemName(ItemName, True) = MealPlans.Item(CurrentPlan).P_Discount Then
            PlanDiscountTake = True
            Exit For
         End If
      Next LS
   Next La
End Function
Function Can_Continue(Accessfor As String) As Boolean
'// handles access for system
On Error GoTo ErrorHandler
If LoggedOnAs(0).EMPName = "" Then
  Can_Continue = True
  Exit Function
End If
If Accessfor = "None" Then
  Can_Continue = True
  Exit Function
End If

Dim AccessRequired As Byte
'Dim FingerPrintID  As String
If Accessfor = "Terminal Configuration" Then
  AccessRequired = 9
Else
    If Accessfor = "Supervisor" Then
      AccessRequired = 9
    Else
     If AccessLevels.Item(Accessfor) Is Nothing Then
      AccessRequired = 1
     Else
      AccessRequired = AccessLevels.Item(Accessfor).m_Access
     End If
    End If
    If AccessRequired = 10 And LoggedOnAs(0).EmpAccessLevel < 9 Then
      Can_Continue = False: Exit Function
    ElseIf AccessRequired = 10 And LoggedOnAs(0).EmpAccessLevel = 9 Then
      AccessRequired = 9
    End If
End If
'Locks Button only access level of 9 will access system
'Function added 05-10-01
'///
If LoggedOnAs(0).EmpAccessLevel >= AccessRequired Then
    Can_Continue = True
Else
    ReturnAccessToNormal = LoggedOnAs(0).EmpAccessLevel
    ReturnToUser = LoggedOnAs(0).EMPName
    Can_Continue = False
            See_Msg "Access Required", 5, "Higher Access Required..."
            AccessOnly = True
            Punch_Clock = False
            Check_ID.Show 1
            AccessOnly = False
            Call MemoryCheck(Check_ID)
                 If Val(LoggedOnAs(0).EmpAccessLevel) < AccessRequired Then
                     See_Msg Accessfor, 5, "Access Denied.  "
                     Can_Continue = False
                     UserVoid.Manager = ""
                     LoggedOnAs(0).VoidAccess = 0
                     LoggedOnAs(0).EmpAccessLevel = ReturnAccessToNormal
                     Exit Function
                 Else
                     Can_Continue = True
                     LoggedOnAs(0).ManagerOk = UserVoid.Manager
                     LoggedOnAs(0).VoidAccess = LoggedOnAs(0).EmpAccessLevel
                     'LoggedOnAs(0).EmpAccessLevel = ReturnAccessToNormal
                     '// causes problems with order screen to managers menu
                     Exit Function
                 End If
    Can_Continue = False
    ResetUser
End If
Exit Function
ErrorHandler:
AccessRequired = 1
Resume Next
'// if no access if found the default is level 1
End Function
Public Function SetAttributes()
On Error Resume Next
Dim MYName As String
MYName = Dir(DBFilePath & "*.Mdb")   ' Retrieve the first entry.
Do While MYName <> ""   ' Start the loop.
   SetAttr DBFilePath & MYName, vbArchive
   MYName = Dir   ' Get next entry.
Loop
End Function
Public Function SetVoid() As Boolean
On Error Resume Next
Dim Cr          As Integer
Dim JK          As String
Dim Asterisk    As Boolean
For Cr = 0 To NextRow(0) - 1
    JK = ProcessItemName(CStr(Grid_array(0, Cr, 1)), Asterisk)
    If Asterisk = True Then
      SetVoid = True
      Exit For
    End If
Next Cr
End Function
Public Function DiscountNotAllowed(ItemIn As String) As Boolean
On Error GoTo ERH
DiscountNotAllowed = False
If AllMenuItems(ItemIn) Is Nothing Then Exit Function
 If CouponDepartment <> "Discount" And CouponDepartment <> "None" Then
  If AllMenuItems(ItemIn).ReportCategory = CouponDepartment And AllMenuItems(ItemIn).AllowDiscount = 0 Or CouponDepartment = "Department" And AllMenuItems(ItemIn).AllowDiscount = 0 Or CouponDepartment = "Item Discount" And AllMenuItems(ItemIn).AllowDiscount = 0 Then
   DiscountNotAllowed = False
   If Check_Option(87) = 1 Then ItemIn = AllMenuItems(ItemIn).ReportCategory
  Else
   DiscountNotAllowed = True
  End If
Else
 If AllMenuItems(ItemIn).AllowDiscount = 1 Then
  DiscountNotAllowed = True
 Else
  DiscountNotAllowed = False
 End If
End If
Exit Function
ERH:
DiscountNotAllowed = False
End Function
Public Function CheckDate(EventNo As Integer) As Boolean
  If Time_Events(EventNo).StartDate = "None" Or Time_Events(EventNo).EndDate = "None" Then
     CheckDate = True
     Exit Function
  End If
  If Not IsDate(Time_Events(EventNo).StartDate) Or Not IsDate(Time_Events(EventNo).EndDate) Then
     CheckDate = True
     Exit Function
  End If
  Dim DateStart As Date
  Dim DateEnd   As Date
  Dim DaysLeft  As Integer
  Dim TotalDays As Integer
  Dim MyDate    As Date
  MyDate = Date '"1/09"
  'Time_Events(EventNo).StartDate = "10/01"
  'Time_Events(EventNo).EndDate = "04/07"
  CheckDate = False
  If DateDiff("d", CDate(Time_Events(EventNo).StartDate), CDate(Time_Events(EventNo).EndDate)) < 0 Then
     TotalDays = DateDiff("d", CDate(Time_Events(EventNo).StartDate), CDate(Time_Events(EventNo).EndDate)) + 365
     If DateDiff("d", Format(MyDate, "MM/DD"), CDate(Time_Events(EventNo).StartDate)) <= 0 Then
       DaysLeft = DateDiff("d", Format(MyDate, "MM/DD"), CDate(Time_Events(EventNo).EndDate)) + 365
       If DaysLeft >= 0 Then
          CheckDate = True
       End If
     ElseIf DateDiff("d", Format(MyDate, "MM/DD"), CDate(Time_Events(EventNo).EndDate)) >= 0 Then
       DaysLeft = DateDiff("d", Format(MyDate, "MM/DD"), CDate(Time_Events(EventNo).EndDate))
       If DaysLeft >= 0 Then
         CheckDate = True
       End If
     End If
  Else
     TotalDays = DateDiff("d", CDate(Time_Events(EventNo).StartDate), CDate(Time_Events(EventNo).EndDate))
     DaysLeft = DateDiff("d", Format(MyDate, "MM/DD"), CDate(Time_Events(EventNo).EndDate))
     If DaysLeft >= 0 And TotalDays - DaysLeft >= 0 Then
       CheckDate = True
     End If
  End If
End Function
Public Function CheckDayEvent(EventDate As Date, EventNo As Integer) As Boolean
 CheckDayEvent = False
                 Select Case WeekDay(EventDate, vbMonday)
                       Case 1
                          If Time_Events(EventNo).Monday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                       Case 2
                          If Time_Events(EventNo).Tuesday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                       Case 3
                          If Time_Events(EventNo).Wednesday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                       Case 4
                          If Time_Events(EventNo).Thursday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                       Case 5
                          If Time_Events(EventNo).Friday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                       Case 6
                          If Time_Events(EventNo).Saturday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                       Case 7
                          If Time_Events(EventNo).Sunday = 1 Then
                             CheckDayEvent = True
                             Exit Function
                          End If
                  End Select
End Function
Public Function Check_Available(EventName As String, GoodOnly As String, WhichTable As Integer) As Boolean
On Error GoTo erHand
Dim Cat_Event(5) As String
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Lk As Integer
Check_Available = True
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
If WhichTable = 0 Then
  Sql = "Select * from [Sub Categories] where [Sub Name] = '" & EventName & "'"
Else
  Sql = "Select * from [Shift Keys] where [Key Name] = '" & Trim$(EventName) & "' or [Description] = '" & EventName & "'"
End If
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
 DbClose Db, Tb
 Exit Function
Else
 Cat_Event(0) = Tb![Event 1] & ""
 Cat_Event(1) = Tb![Event 2] & ""
 Cat_Event(2) = Tb![Event 3] & ""
 Cat_Event(3) = Tb![Event 4] & ""
 Cat_Event(4) = Tb![Event 5] & ""
End If
DbClose Db, Tb
Dim SRC   As Integer
Dim Wkday As Integer
     For SRC = 0 To 4
          If UCase(Cat_Event(0)) = "NONE" Then
             Check_Available = True
             Exit For
          Else
            For Lk = 0 To Total_Events
             If Trim(Cat_Event(SRC)) = Trim(Time_Events(Lk).EName) Then
               If CheckDate(Lk) And IsValidTime(Time, Time_Events(Lk).StartTime, Time_Events(Lk).EndTime, Lk) Then
                  Check_Available = True
                  Exit Function
               Else
                  GoodOnly = "Good Only From " & Time_Events(Lk).StartTime & " to " & Time_Events(Lk).EndTime
                  If Time_Events(Lk).StartDate <> "None" Then
                    GoodOnly = "Good Only On " & Time_Events(Lk).StartDate & " To " & Time_Events(Lk).EndDate
                  End If
                  Check_Available = False
               End If
              End If
          Next Lk
      End If
     Next SRC
Erase Cat_Event
Exit Function
erHand:
Err.Clear
End Function
Public Function OKTOROUTE(EventName As String) As Boolean
Dim Lk As Integer
     For Lk = 0 To Total_Events
       If Trim(EventName) = Trim(Time_Events(Lk).EName) Then
         If CheckDate(Lk) And IsValidTime(Time, Time_Events(Lk).StartTime, Time_Events(Lk).EndTime, Lk) Then
             OKTOROUTE = True
             Exit Function
         Else
             OKTOROUTE = False
             Exit Function
         End If
       End If
    Next Lk
End Function
Public Function AccountIquire(AccountNo As String) As Boolean
On Error GoTo ErrorHandler
Dim PrinterPort As String
Dim AllData$
Dim FailureMsg  As String
Dim ReturnedBal As Currency
Dim Txt         As String
ReturnedBal = GetAccountBalance(AccountNo)
If ReturnedBal < 0 Then
  Txt = "Balance Owed "
Else
  Txt = "Balance Remaining "
End If
See_Msg "Account Balance", 0, Txt & Format(Abs(ReturnedBal), "0.00")
If See_Msg("Print Copy ?", 4, "") = 6 Then
    PrinterPort = GetSpoolPrinter(AllPrinter.Item("Receipt").PrinterName)
    If RemotePrinterFailed(AllPrinter.Item("Receipt").PrinterName, PrinterPort, AllPrinter.Item("Receipt").Failed, FailureMsg) = True Then Exit Function
    EC(7) = vbCrLf
    AllData$ = EC(0)  'Int printer
    AllData$ = AllData$ & EC(8) & EC(1) & EC(3) & StoreName & EC(2) & EC(7)
    AllData$ = AllData$ & EC(9) & EC(1) & EC(3) & String(40, "*") & EC(2) & EC(7)
    AllData$ = AllData$ & EC(1) & EC(3) & "Account Number   " & AccountNo & EC(2) & EC(7)
    AllData$ = AllData$ & EC(1) & EC(3) & Txt & Format(Abs(ReturnedBal), "0.00") & EC(2) & EC(7)
    AllData$ = AllData$ & EC(1) & EC(3) & String(40, "*") & EC(2) & EC(7)
    AllData$ = AllData$ & EC(6) & EC(5)
    ReceiptChit AllData$, PrinterPort
End If
Exit Function
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Account Inquire ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetAccountData(ByVal InCommingData As String) As Boolean
On Error GoTo ErrorHandler
SetMemberToNothing
GetAccountData = False
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Dim RF              As Long
Dim TenderNameLink  As String
Dim ACCTNAME        As String
Dim Tb              As ADODB.Recordset
Dim GRS             As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
    Sql = "Select * from [HouseAccountData] where  (([Acct Number] = '" & InCommingData & "')  Or ([Scan Name] = '" & InCommingData & "')"
    Sql = Sql & " Or ([Spc Interest 1] = '" & InCommingData & "') Or ([Spc Interest 2] = '" & InCommingData & "')"
    Sql = Sql & " Or ([Spc Interest 3] = '" & InCommingData & "')  Or ([Spc Interest 4] = '" & InCommingData & "'))"
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
    If Not Tb Is Nothing Then
       If Not Tb.BOF Then Tb.MoveFirst
    Else
      See_Msg "Cannot Connect to House Account", 5, ""
      Exit Function
    End If
    RF = RecordsCount(Tb)
    If RF > 1 Then
      See_Msg RF & " Records Found", 5, "Please Select Name Search Button."
    End If
    If Tb.EOF Then
        DbClose Nothing, Tb
        Call See_Msg("Unable to charge on this account", 5, "Not a Valid Account")
        Exit Function
    Else
       TenderNameLink = GetGroupTenderLink(Tb![Account Type] & "")
       If TenderNameLink > "" And TenderNameLink <> "None" Then
         If TenderLink > "" And TenderLink <> TenderNameLink Then
           DbClose Nothing, Tb
           Call See_Msg("Wrong Tender For This Type Of Account", 0, "Use " & TenderNameLink)
           Exit Function
         End If
       End If
       If ReturnExp(Tb![Account Expires On] & "") = True Then
        See_Msg "Unable to charge on this account", 0, "Account Expired On " & Tb![Account Expires On] & ""
        DbClose Nothing, Tb
        Exit Function
       End If
       If Tb![Acct Number] & "" = "" Then
         DbClose Nothing, Tb
         See_Msg "No Account Number ", 0, "Account Number Empty"
         Exit Function
       End If
       ACCTNAME = Tb![Acct Number] & ""
       CurrentAccount = ACCTNAME
       If Check_Option(141) = 1 Then
        MemberName.m_Number = Tb(0) & ""
        MemberName.m_Name = Tb![Acct Number] & ""
        MemberName.m_Address = Tb![Address] & ""
       Else
        MemberName.m_Name = Tb(0) & ""
        MemberName.m_Number = Tb![Acct Number] & ""
        MemberName.m_Address = Tb![Address] & ""
       End If
       If HouseAccounts Is Nothing Then Set HouseAccounts = New AllAccounts
       With HouseAccounts
          If .Item(ACCTNAME) Is Nothing Then
             .add ACCTNAME
          End If
             .Item(ACCTNAME).AccountName = Tb(0)
             .Item(ACCTNAME).DailyMax = Val(Tb(1) & "")
             .Item(ACCTNAME).WeeklyMax = Val(Tb(2) & "")
             .Item(ACCTNAME).AccountNo = Tb(3) & ""
             .Item(ACCTNAME).DayCharge = Abs(Val(Tb(4) & ""))
             .Item(ACCTNAME).AccountOnHold = Abs(Val(Tb(5) & ""))
             .Item(ACCTNAME).AccountNotActive = Abs(Val(Tb(6) & ""))
             .Item(ACCTNAME).CreditBalOnly = Abs(Val(Tb![Credit Balance] & ""))
             .Item(ACCTNAME).AccountMemo = ""
             .Item(ACCTNAME).AccountExpired = ReturnExp(Tb![Account Expires On] & "")
             .Item(ACCTNAME).AccountType = Tb![Account Type] & ""
             .Item(ACCTNAME).MaxMeals = Val(Tb![Meal Limit] & "")
             .Item(ACCTNAME).AccountAddress = Tb![Address] & ""
             If Tb![Meal Limit] & "" > "" Then
              .Item(ACCTNAME).DiscountOnFile = Tb![Discount] & ""
             Else
              .Item(ACCTNAME).DiscountOnFile = "None"
             End If
             .Item(ACCTNAME).CardOnFile = ""
             .Item(ACCTNAME).CardExpOnFile = ""
             GetAccountData = True
        End With
    End If
DbClose Nothing, Tb
Set GRS = Nothing
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Description, Err.Number, "Get Account Data " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub AccountOnHold(ActNo As String)
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & ActNo & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
    Tb![Suspend] = -1
    Tb![Memo] = Tb![Memo] & vbCrLf & " Charge Declined " & Format(Date, "MM/DD/YYYY") & vbCrLf
    Tb.Update
End If
DbClose Db, Tb
Exit Sub
ERH:
Err.Clear
End Sub
Public Sub SuspendAllOpenChecks()
On Error GoTo ErrorHandler
  Dim Db As ADODB.Connection
  Dim Tb As ADODB.Recordset
  If Check_Option(139) = 0 Then Exit Sub   '// option 139 allows for suspended Checks
  See_Msg "Suspending all open checks", 5, ""
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
   Sql = "select * from [Open Checks] where [close out day] = 0 "
  If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
  Do While Not Tb.EOF
    Sql = "Update [Check Detail] Set [Close Out Day] = -1 where [close out day] = 0 and [check number] = " & Tb.fields("Check Number").Value
    Db.Execute Sql
   Tb.MoveNext
  Loop
  Sql = "Update [Open Checks] set [close out day] = -1 where [close out day] = 0 "
  Db.Execute Sql
  Db.Close
  Trans_Log "Auto Suspend all open checks"
  Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Suspend All Checks")
  Select Case ErrorRespose
    Case "Abort"
      Exit Sub
    Case Is = "Resume"
      Resume
    Case Is = "Resume Next"
      Resume Next
  End Select
End Sub
Public Function UnSuspendAllOpenChecks()
  On Error GoTo ErrorHandler
  If Check_Option(139) = 0 Then Exit Function
  '// option 139 allows for suspended Checks
  '// allow daily close with open checks
  Dim Db  As ADODB.Connection
  Dim Adb As ADODB.Connection
  Dim ARs As ADODB.Recordset
  Dim RS  As ADODB.Recordset
  Dim dS  As ADODB.Recordset
  See_Msg "Reopening all suspended checks", 5, ""
  ConnectADODB SvrPath.TerminalPath & "MSaleData", Adb
  If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then
     See_Msg "Failed To Connect to Check File", 5, ""
     Exit Function
  End If
   Sql = "Update [Open Checks] Set [Close Out Day] = 0 where [close out day] = -1 "
  Db.Execute Sql
   Sql = "Update [Check Detail] Set [Close Out Day] = 0 where [close out day] = -1 "
  Db.Execute Sql
   Sql = "Select distinct [Server Name] from [open checks] Where [close out day] = 0 order by [Server Name] asc"  ' group by [server name]"
   If ConnectADORS(Sql, Db, RS) = False Then
     See_Msg "Failed To Connect to open checks", 5, ""
     Exit Function
   End If
   Do While Not RS.EOF
       Sql = "Select * from [Check Detail] where [servers name] = '" & RS.fields("Server Name").Value & "' and [close out day] = 0 order by [check number] asc"
       If ConnectADORS(Sql, Db, dS) Then
         If Not dS.EOF Then
           Sql = "Select * from [Check Numbers]"
           ConnectADORS Sql, Adb, ARs
           ARs.AddNew
           ARs.fields("servers name").Value = RS.fields("Server Name").Value
           ARs.fields("check no").Value = dS.fields("Check number").Value
           ARs.fields("Time").Value = Format(Str(Time), "HH:MM ampm")
           ARs.fields("Start Date").Value = Format(Date, "MM/DD/YY")
           ARs.fields("Audit Status").Value = 0
           ARs.fields("In Use").Value = 0
           ARs![Register Name] = ""
           ARs![Register Group] = SvrPath.TerminalGroupName
           ARs![Drawer No] = 0
           ARs![Work Area] = ""
           ARs![Starting Bank] = 0
           ARs![Deposits Paid In] = 0
           ARs.fields("Starting Balance").Value = 0
           ARs.Update
           See_Msg "Starting Audit For " & RS.fields("Server Name").Value, 5, ""
           Trans_Log "Starting Audit for " & RS.fields("Server Name").Value
           RecordShiftNumber RS.fields("Server Name").Value, 1
         End If
       End If
   RS.MoveNext
  Loop
  Sql = "Update [Check Detail] Set [Check date] = '" & Format(Date, "MM/DD/YYYY") & "' where [close out day] = 0 "
  Db.Execute Sql
  Sql = "Update [Open Checks] Set [Order date] = '" & Format(Date, "MM/DD/YYYY") & "' where [close out day] = 0 "
  dS.Close
  Db.Execute Sql
  DbClose Db, RS
  DbClose Adb, ARs
  Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Unsuspend All Checks")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetMainModifierPick(AnchorModifier As String, ModTitle As String, ModName As String) As String
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim ModLink As String
OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
   Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
   Sql = Sql & " and [Modifier Category] = '" & ModName & "'"
ConnectADORS Sql, Db, Tb
   If Not Tb.EOF Then
      GetMainModifierPick = ModName
   Else
      GetMainModifierPick = ""
   End If
  DbClose Db, Tb
End Function
Public Function IsChargeable(Db As ADODB.Connection, ModTitle As String, ModName As String, StartRow As Integer, EndRow As Integer) As Integer
Dim Tb      As ADODB.Recordset
Dim ModLink As String
Dim FI      As Integer
IsChargeable = 0
For FI = StartRow + 1 To EndRow
    Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
    Sql = Sql & " and [Modifier Category] = '" & Grid_array(SeatNumber, FI, 1) & "'"
   ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
      If Val(Grid_array(SeatNumber, FI, 2)) = 0 Then
       IsChargeable = IsChargeable + 1
       If ModName = Grid_array(SeatNumber, FI, 1) Then
        IsChargeable = IsChargeable - 1
        Exit For
       End If
      End If
    End If
   SetNothing Tb
 Next FI
End Function
Public Function GetModifierPick(AnchorModifier As String, ModTitle As String, ModName As String) As String
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim ModLink As String
OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
   Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
   Sql = Sql & " and [Modifier Category] = '" & ModName & "'"
ConnectADORS Sql, Db, Tb
   If Not Tb.EOF Then
      GetModifierPick = ModName
   Else
     If AnchorModifier > "" Then
       SetNothing Tb
       Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
       Sql = Sql & " and [Modifier Category] = '" & AnchorModifier & "'"
       Sql = Sql & " and [Modified By] <> 'None'"
       ConnectADORS Sql, Db, Tb
       If Not Tb.EOF Then
          ModLink = Tb![Modified By] & ""
          Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModLink & "'"
          Sql = Sql & " and [Modifier Category] = '" & ModName & "'"
          SetNothing Tb
          ConnectADORS Sql, Db, Tb
           If Not Tb.EOF Then
            GetModifierPick = ModName
            DbClose Db, Tb
            Exit Function
           End If
       Else
        GetModifierPick = ""
       End If
     Else
       Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
       Sql = Sql & " and [Modifier Category] = '" & ModName & "'"
       

     
     End If
     GetModifierPick = ""
   End If
  DbClose Db, Tb
End Function
Public Function GetModifierLink(AnchorModifier As String, ModTitle As String, ModName As String) As String
  Dim Db      As ADODB.Connection
  Dim Tb      As ADODB.Recordset
  Dim ModLink As String
  OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
       Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
       Sql = Sql & " and [Modifier Category] = '" & AnchorModifier & "'"
       Sql = Sql & " and [Modified By] <> 'None'"
       ConnectADORS Sql, Db, Tb
       If Not Tb.EOF Then
          ModLink = Tb![Modified By] & ""
          Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModLink & "'"
          Sql = Sql & " and [Modifier Category] = '" & ModName & "'"
          SetNothing Tb
          ConnectADORS Sql, Db, Tb
           If Not Tb.EOF Then
            GetModifierLink = ModName
            DbClose Db, Tb
            Exit Function
           End If
       Else
        GetModifierLink = ""
       End If
  DbClose Db, Tb
End Function
Public Function ModifierComesWith(ModTitle As String) As Boolean
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
   Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModTitle & "'"
   Sql = Sql & " and [ComesWith] > 0"
ConnectADORS Sql, Db, Tb
   If Not Tb.EOF Then
      ModifierComesWith = True
   Else
      ModifierComesWith = False
   End If
  DbClose Db, Tb
End Function
Public Function GetAltModifier(ModName As String) As String
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
 ConnectADODB SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
   Sql = "Select * from [Modifiers] Where [Modifier Category] = '" & ModName & "'"
   Sql = Sql & " and [Alt Modifier] <> 'None'"
 ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     GetAltModifier = Tb.fields("Alt Modifier").Value & ""
  Else
     GetAltModifier = "None"
  End If
  DbClose Db, Tb
End Function
Public Function FindOrigMenuItem(MainItemIn As String) As String
On Error GoTo ERH
Dim Db       As ADODB.Connection
Dim Tb       As ADODB.Recordset
Dim s        As Integer
Dim L        As Integer
Dim AllPcs() As String
Dim BuildStr As String
AllPcs = Split(MainItemIn, " ")
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
BuildStr = ""
For L = 0 To UBound(AllPcs)
 BuildStr = ""
 For s = L To UBound(AllPcs)
  BuildStr = BuildStr & " " & AllPcs(s)
  Sql = "Select * from [Menu Items] where [Menu Item Name] = '" & Trim$(BuildStr) & "'"
  ConnectADORS Sql, Db, Tb
   If Not Tb.EOF Then
      FindOrigMenuItem = Trim$(BuildStr)
      DbClose Db, Tb
      Exit Function
   End If
   SetNothing Tb
 Next s
Next L
ERH:
DbClose Db, Tb
FindOrigMenuItem = MainItemIn
End Function
Public Function GetOnTheFlyShiftPrice(ModifierTitle As String, ModCategory As String, ShiftLevel As Integer) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Hp As Single
   If UCase(Left$(ModifierTitle, 2)) = "NO" Then
      GetOnTheFlyShiftPrice = 0
      Exit Function
   End If
   Hp = 1
   OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
   If Left$(Trim$(ModifierTitle), 3) = Trim$(Input_Option(93)) Or Left$(Trim$(ModifierTitle), 3) = Trim$(Input_Option(94)) Then
      Sql = "Select * from [Modifier Names] Where [Modifiers] = '" & ModCategory & "'"
      ConnectADORS Sql, Db, RS
      If Not RS.EOF Then
        Hp = GetHalfCharge(RS.fields("Department").Value & "")
      End If
      RS.Close
   End If
   Sql = "Select * from [Price Levels] where [Modifier Title] = '" & RemovePreFix(ModifierTitle) & "'"
   Sql = Sql & " and [Modifier Category] = '" & Trim$(ModCategory) & "'"
   Sql = Sql & " and [Level] = " & ShiftLevel - 1
   ConnectADORS Sql, Db, RS
   If Not RS.EOF Then
     GetOnTheFlyShiftPrice = Format(SetDualPrice(Val(RS.fields("Price").Value & "")) * Hp, "0.00")
   Else
     GetOnTheFlyShiftPrice = 0
   End If
 DbClose Db, RS
End Function
Public Function GetHalfCharge(ItemIn As String) As Single
Dim GetChg() As String
If InStr(ItemIn, "-") Then
  GetChg = Split(ItemIn, "-")
  If Val(GetChg(1)) > 0 Then
    GetHalfCharge = Val(GetChg(1)) / 100
  End If
ElseIf InStr(ItemIn, "$") Then
  GetChg = Split(ItemIn, "$")
  'If Val(GetChg(1)) > 0 Then
    GetHalfCharge = Val(GetChg(1))
  'End If
Else
 If Val(Input_Option(31)) > 0 Then
  GetHalfCharge = Val(Input_Option(31)) / 100
 Else
  GetHalfCharge = GetPrefixMultiplyer(ItemIn)
 End If
End If
End Function
Public Function GetQtrCharge(ItemIn As String) As Single
'GetQtrCharge = GetPrefixMultiplyer(ItemIn)
'If GetQtrCharge > 0 Then Exit Function
If InStr(ItemIn, "-") Then
  Dim GetChg() As String
  GetChg = Split(ItemIn, "-")
  If Val(GetChg(1)) > 0 Then
    GetQtrCharge = (Val(GetChg(1)) / 2) / 100
  End If
Else
  If Val(Input_Option(31)) > 0 Then
    GetQtrCharge = (Val(Input_Option(31)) / 2) / 100
  Else
    GetQtrCharge = 1
  End If
End If
End Function

Public Function GetModifierToShow(ItemName As String, ModName As String, AddTo As Integer) As String
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim Mods(5) As String
Dim f       As Integer
'Dim AltMod  As String
GetModifierToShow = "None"
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
Sql = "Select * from [Menu Items] where [Menu Item Name] = '" & ItemName & "'"
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
    'DbClose DB, Tb
    'FindOrigMenuItem (EditItem)
    SetNothing Tb
    Sql = "Select * from [Menu Items] where [Menu Item Name] = '" & FindOrigMenuItem(ItemName) & "'"
    ConnectADORS Sql, Db, Tb
    If Tb.EOF Then
      DbClose Db, Tb
      Exit Function
    Else
      GoTo DM
    End If
Else
DM:
  Mods(0) = Tb![Modifier 1] & ""
  Mods(1) = Tb![Modifier 2] & ""
  Mods(2) = Tb![Modifier 3] & ""
  Mods(3) = Tb![Modifier 4] & ""
  Mods(4) = Tb![Modifier 5] & ""
  Mods(5) = Tb![Modifier 6] & ""
  DbClose Db, Tb
 ' ModName = Mods(0)
  If ModName = "" Then
    GetModifierToShow = Mods(0)
    Exit Function
  End If
  OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
  For f = 0 To 5
   Sql = "Select * from [Modifiers] where [Modifier Name] = '" & Mods(f) & "'"
   Sql = Sql & " and [Modifier Category] = '" & Trim$(ModName) & "'"
  ConnectADORS Sql, Db, Tb
   If Not Tb.EOF Then
      If AddTo = -2 And Tb![Modified By] & "" <> "None" Then   'ShiftLevel = -2 And
        GetModifierToShow = Tb![Modified By] & ""
      Else
        GetModifierToShow = Mods(f)
      End If
      DbClose Db, Tb
      Exit Function
   End If
   SetNothing Tb
  Next f
  DbClose Db, Tb
End If
End Function
Public Function GetModPriceShiftEdit(ItemName As String, ModName As String, ShiftLVL As Integer, ShiftPFx As String) As Currency
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim Mods(4) As String
Dim f       As Integer
GetModPriceShiftEdit = -1
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
Sql = "Select * from [Menu Items] where [Menu Item Name] = '" & ItemName & "'"
ConnectADORS Sql, Db, Tb
  If Tb.EOF Then
    DbClose Db, Tb
    Exit Function
  Else
    Mods(0) = Tb![Shift Mod 1] & ""
    Mods(1) = Tb![Shift Mod 2] & ""
    Mods(2) = Tb![Shift Mod 3] & ""
    Mods(3) = Tb![Shift Mod 4] & ""
    Mods(4) = Tb![Shift Mod 5] & ""
    'Mods(5) = Tb![Shift Mod 6] & ""
  End If
  DbClose Db, Tb
  OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
  For f = 0 To 4
   Sql = "Select * from [Price Levels] where [Modifier Category] = '" & Mods(f) & "'"
   Sql = Sql & " and [Modifier Title] = '" & ModName & "'"
   Sql = Sql & " and [Level] = " & ShiftLVL
   ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
     GetModPriceShiftEdit = Format(SetDualPrice(Val(Tb![Price] & "")), "0.00")
     ShiftPFx = Tb.fields("PreFix").Value & ""
     Exit For
    End If
  Next f
  DbClose Db, Tb
End Function
Public Function GetSubModifier(ItemName As String, ModName As String) As String
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim AltMod  As String
Dim Mods(5) As String
Dim f       As Integer
GetSubModifier = "None"
OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
Sql = "Select * from [Menu Items] where [Menu Item Name] = '" & ItemName & "'"
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
    DbClose Db, Tb
    Exit Function
Else
  Mods(0) = Tb![Modifier 1] & ""
  Mods(1) = Tb![Modifier 2] & ""
  Mods(2) = Tb![Modifier 3] & ""
  Mods(3) = Tb![Modifier 4] & ""
  Mods(4) = Tb![Modifier 5] & ""
  Mods(5) = Tb![Modifier 6] & ""
End If
DbClose Db, Tb
  OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
  For f = 0 To 5
   Sql = "Select * from [Modifiers] where [Modifier Name] = '" & Mods(f) & "'"
   Sql = Sql & " and [Modified by] <> 'None'"
  ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
     AltMod = Tb![Modified By] & ""
     SetNothing Tb
     Sql = "Select * from [Modifiers] where [Modifier Name] = '" & AltMod & "'"
     Sql = Sql & " and [Modifier Category] = '" & ModName & "'"
     ConnectADORS Sql, Db, Tb
      If Not Tb.EOF Then
       GetSubModifier = AltMod & ""
       DbClose Db, Tb
       Exit Function
      End If
    End If
  Next f
  DbClose Db, Tb
End Function
Public Function GetPreFixMod(ModName As String) As Boolean
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
 Sql = "Select * from [Modifiers] where [Modifier Name] = '" & ModName & "'"
 Sql = Sql & " and [Modified by] <> 'None'"
ConnectADORS Sql, Db, Tb
 If Tb.EOF Then
   GetPreFixMod = False
 Else
   GetPreFixMod = True
 End If
DbClose Db, Tb
End Function
Public Function GetMatrixPrice(ModCat As String, ModName As String, PFXItem, SR As Integer, Er As Integer) As Integer
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim NoChargeItems   As Integer
GetMatrixPrice = 0
If PFXItem = "" Then
  Exit Function
End If
OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", Db
Sql = "Select * from [Modifier Names] Where [Modifiers] = '" & ModCat & "'"
OpenTb Sql, RS, Db
If Not RS.EOF Then
 If Left(UCase(RS.fields("Department").Value), 5) = "UPICK" Then
   If InStr(PFXItem, "-") > 0 Then
     Sql = "Select * from [Price Levels] Where [Modifier Category] = '" & ModCat & "' and [Modifier Title] = '" & ModName & "' and [Prefix] = '" & PFXItem & "'"
    OpenTb Sql, RS, Db
     If Not RS.EOF Then
      GetMatrixPrice = Val(RS.fields("Level").Value & "") + 1
     End If
    DbClose Db, RS
    Exit Function
   End If
   NoChargeItems = IsChargeable(Db, ModCat, ModName, SR, Er)
   If NoChargeItems < Right(RS.fields("Department").Value, 1) Then
      DbClose Db, RS
      GetMatrixPrice = 0
      Exit Function
   End If
   Sql = "Select * from [Price Levels] Where [Modifier Category] = '" & ModCat & "' and [Modifier Title] = '" & PFXItem & "' and [Price] > 0"
  OpenTb Sql, RS, Db
   If Not RS.EOF Then
    GetMatrixPrice = Val(RS.fields("Level").Value & "") + 1
   End If
 End If
End If
DbClose Db, RS
End Function

Public Function ReturnExp(InDate As String) As Boolean
On Error Resume Next
If IsDate(InDate) Then
  If CDate(InDate) <= Date Then
     ReturnExp = True
     Exit Function
  End If
End If
ReturnExp = False
End Function
Public Function GetMember() As String
  SetMemberToNothing
  Scan_Charge.Show 1
  If GetSelected = "" Then Exit Function
  MemberName.m_Name = GetAccountName(GetSelected)
  GetSelected = ""
  If MemberName.m_Name = "" Then
     GetMember = ""
  Else
     GetMember = MemberName.m_Name
 End If
 'handy comp interface use account number for member accounts
End Function
Public Function Punched_In(SrvName As String) As Boolean
'function added to suport check option (74) clock out before audit
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Dim TRDB As ADODB.Connection, TRRS As ADODB.Recordset
On Error GoTo ErrorHandler
   OpenDb SvrPath.TerminalPath & TimeRecordsDb, TRDB
    Sql = "Select * from [Employee Records] where [Employee Name] = '" & SrvName & "'"
    Sql = Sql & " and [Time In] > '' and [Time out] = ''"
   ConnectADORS Sql, TRDB, TRRS
   If TRRS.EOF Then
      Punched_In = False
   Else
      Punched_In = True
   End If
DbClose TRDB, TRRS
Exit Function
ErrorHandler:
Punched_In = False
End Function

Public Function AssignedEvents(EventItem As String) As Integer
     AssignedEvents = 0
     If AllMenuItems(EventItem).TimeEvent1 <> "None" Then
        AssignedEvents = 1
        'Exit Function
     End If
     If AllMenuItems(EventItem).TimeEvent2 <> "None" Then
        AssignedEvents = 2
        'Exit Function
     End If
     If AllMenuItems(EventItem).TimeEvent3 <> "None" Then
        AssignedEvents = 3
        'Exit Function
     End If
     If AllMenuItems(EventItem).TimeEvent4 <> "None" Then
        AssignedEvents = 4
        'Exit Function
     End If
     If AllMenuItems(EventItem).TimeEvent5 <> "None" Then
        AssignedEvents = 5
        'Exit Function
     End If
End Function
Public Function TimedPrice(MItem As String, CurrentPrice As Currency, EventName As String) As Currency
On Error GoTo erHand
Dim ETime As Date
Dim Lk As Integer
Dim Cev As Integer
EventName = ""
Order.Information.Caption = ""
If Check_Option(222) = 1 Then
 If OrderIs = GetOrderType(GoSale) Then
   TimedPrice = CurrentPrice
   Exit Function
 End If
End If
If OrderIs = GetOrderType(PKUPSale) Then
   If GetConfigInfo("Phone Orders", "Timed Event Pricing is Excluded Pick Up", "Yes", "Functions.INI") = "ENABLED" Then
     TimedPrice = CurrentPrice
     Order.Information.Caption = "No Event Pricing for Phone"
     Exit Function
   End If
End If
If OrderIs = GetOrderType(DelSale) Then
   If GetConfigInfo("Phone Orders", "Timed Event Pricing is Excluded Delivery", "Yes", "Functions.INI") = "ENABLED" Then
     TimedPrice = CurrentPrice
     Order.Information.Caption = "No Event Pricing for Phone"
     Exit Function
   End If
End If


'If Trim(ItemTime) = "" Then
  ETime = Format(Time, "HH:MM AmPm")
'Else
'  ETime = CDate(ItemTime)
'End If
If AllMenuItems(MItem) Is Nothing Then
  TimedPrice = CurrentPrice
  Exit Function
Else
    If AllMenuItems(MItem).PriceLevels = 0 Then   'Timed Prices are Not used
      TimedPrice = CurrentPrice
      Exit Function
    End If
End If
If AllMenuItems(MItem).TimeEvent1 = "On Sale" Then
   TimedPrice = Val(AllMenuItems(MItem).TimePrice1)
   Exit Function
End If
           For Lk = 0 To AssignedEvents(MItem) - 1
                'check for match
 
              For Cev = 0 To Total_Events 'UBound(Time_Events(), 1) - 1
                Select Case Lk
                   Case Is = 0
                      If AllMenuItems(MItem).TimeEvent1 = Trim$(Time_Events(Cev).EName) Then
                         If CheckDate(Cev) And IsValidTime(ETime, Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev) Then
                             TimedPrice = Val(AllMenuItems(MItem).TimePrice1)
                             ModShiftPLevel = AllMenuItems(MItem).EModShift1
                             ModEventPrice = 0
                             If Time_Events(Cev).PriceAdjust <> 0 Then
                               TimedPrice = CurrentPrice - (Time_Events(Cev).PriceAdjust)
                               If TimedPrice < 0 Then TimedPrice = 0
                               ModEventPrice = Time_Events(Cev).PriceAdjust
                             End If
                              If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
                               TimedPrice = CurrentPrice
                               Order.Information.Caption = "Timed Price Disabled"
                               Exit Function
                              End If
                              If Check_Option(109) = 1 Then
                               EventName = Left(Trim(Time_Events(Cev).EName), 5) & " "
                              End If
                             Exit Function
                           End If
                       End If
                  Case Is = 1
                      If AllMenuItems(MItem).TimeEvent2 = Trim(Time_Events(Cev).EName) Then
                         If CheckDate(Cev) And IsValidTime(ETime, Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev) Then
                             TimedPrice = Val(AllMenuItems(MItem).TimePrice2)
                             ModShiftPLevel = AllMenuItems(MItem).EModShift2
                             ModEventPrice = 0
                             If Time_Events(Cev).PriceAdjust <> 0 Then
                               TimedPrice = CurrentPrice - (Time_Events(Cev).PriceAdjust)
                               If TimedPrice < 0 Then TimedPrice = 0
                               ModEventPrice = Time_Events(Cev).PriceAdjust
                             End If
                             If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
                              TimedPrice = CurrentPrice
                              Order.Information.Caption = "Timed Price Disabled"
                              Exit Function
                             End If
                             If Check_Option(109) = 1 Then
                               EventName = Left(Trim(Time_Events(Cev).EName), 5) & " "
                             End If
                             Exit Function
                           End If
                    End If
                  Case Is = 2
                      If AllMenuItems(MItem).TimeEvent3 = Trim(Time_Events(Cev).EName) Then
                         If CheckDate(Cev) And IsValidTime(ETime, Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev) Then
                             TimedPrice = Val(AllMenuItems(MItem).TimePrice3)
                             ModShiftPLevel = AllMenuItems(MItem).EModShift3
                             ModEventPrice = 0
                             If Time_Events(Cev).PriceAdjust <> 0 Then
                               TimedPrice = CurrentPrice - (Time_Events(Cev).PriceAdjust)
                               If TimedPrice < 0 Then TimedPrice = 0
                               ModEventPrice = Time_Events(Cev).PriceAdjust
                             End If
                             If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
                              TimedPrice = CurrentPrice
                              Order.Information.Caption = "Timed Price Disabled"
                              Exit Function
                             End If
                             If Check_Option(109) = 1 Then
                               EventName = Left(Trim(Time_Events(Cev).EName), 5) & " "
                             End If
                             Exit Function
                           End If
                       End If
                  Case Is = 3
                      If AllMenuItems(MItem).TimeEvent4 = Trim(Time_Events(Cev).EName) Then
                         If CheckDate(Cev) And IsValidTime(ETime, Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev) Then
                             TimedPrice = Val(AllMenuItems(MItem).TimePrice4)
                             ModShiftPLevel = AllMenuItems(MItem).EModShift4
                             ModEventPrice = 0
                             If Time_Events(Cev).PriceAdjust <> 0 Then
                               TimedPrice = CurrentPrice - (Time_Events(Cev).PriceAdjust)
                               If TimedPrice < 0 Then TimedPrice = 0
                               ModEventPrice = Time_Events(Cev).PriceAdjust
                             End If
                             If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
                               TimedPrice = CurrentPrice
                               Order.Information.Caption = "Timed Price Disabled"
                               Exit Function
                             End If
                             If Check_Option(109) = 1 Then
                               EventName = Left(Trim(Time_Events(Cev).EName), 5) & " "
                             End If
                             Exit Function
                           End If
                       End If
                  Case Is = 4
                      If AllMenuItems(MItem).TimeEvent5 = Trim(Time_Events(Cev).EName) Then
                         If CheckDate(Cev) And IsValidTime(ETime, Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev) Then
                             TimedPrice = Val(AllMenuItems(MItem).TimePrice5)
                             ModShiftPLevel = AllMenuItems(MItem).EModShift5
                             ModEventPrice = 0
                             If Time_Events(Cev).PriceAdjust <> 0 Then
                               TimedPrice = CurrentPrice - (Time_Events(Cev).PriceAdjust)
                               If TimedPrice < 0 Then TimedPrice = 0
                               ModEventPrice = Time_Events(Cev).PriceAdjust
                             End If
                             If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
                               TimedPrice = CurrentPrice
                               Order.Information.Caption = "Timed Price Disabled"
                               Exit Function
                             End If
                             If Check_Option(109) = 1 Then
                               EventName = Left(Trim(Time_Events(Cev).EName), 5) & " "
                             End If
                             Exit Function
                           End If
                       End If
              End Select
             Next Cev
            Next Lk
    If AllMenuItems(MItem).Avalibility = 0 Then
     TimedPrice = -1
     See_Msg "This item is not available at this time", 5, "Important Message.."
     EventName = ""
     ModShiftPLevel = 0
     ModEventPrice = 0
    Else
     TimedPrice = CurrentPrice
     ModShiftPLevel = 0 '
     ModEventPrice = 0
    End If
Exit Function
erHand:
Error_Message "Timed Prices", Err.Description
TimedPrice = CurrentPrice
Exit Function
'Resume Next
End Function
Public Function TimedRatesPrice(MItem As String, StartingTimeandDate As String, ReturnStr As String) As Currency
On Error GoTo erHand
Dim ETime                   As Date
Dim Lk                      As Integer
Dim Cev                     As Integer
Dim EventMins               As Long
Dim OpenMinsTotal           As Long
Dim EventRates(0 To 5)      As TimeRates
Dim EventMinsTotal          As Long
Dim AllMins                 As Long
For Cev = 0 To 4
  EventRates(Cev).EventMins = 0
  EventRates(Cev).EventRate = 0
Next Cev
EventMinsTotal = 0
If IsDate(StartingTimeandDate) Then
  OpenMinsTotal = RoundMinute(DateDiff("n", CDate(StartingTimeandDate), Time & Date))
End If
AllMins = OpenMinsTotal
If OpenMinsTotal < 0 Then
  See_Msg "Invalid Time", 5, "Cannot Calculate "
  TimedRatesPrice = 0
  Exit Function
End If
If AllMenuItems(MItem) Is Nothing Then
  TimedRatesPrice = 0
  Exit Function
End If
EventMinsTotal = 0
For Lk = 0 To AssignedEvents(MItem)
   For Cev = 0 To UBound(Time_Events(), 1) - 1
     Select Case Lk
      Case Is = 0
       If AllMenuItems(MItem).TimeEvent1 = Trim(Time_Events(Cev).EName) Then
         If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
         
         Else
          If CheckDate(Cev) Then
             EventMins = EventMinsBetween(CDate(StartingTimeandDate), Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev)
             EventRates(Lk).EventMins = EventMins '- (AllMins + EventMinsTotal)
             EventMinsTotal = EventMinsTotal + EventRates(Lk).EventMins
             If AllMenuItems(MItem).TimePrice1 > 0 Then
               EventRates(Lk).EventRate = (Val(AllMenuItems(MItem).TimePrice1 / 60) * EventRates(Lk).EventMins)
             End If
           End If
         End If
       End If
      Case Is = 1
      If AllMenuItems(MItem).TimeEvent2 = Trim(Time_Events(Cev).EName) Then
         If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
         
         Else
          If CheckDate(Cev) Then
           EventMins = EventMinsBetween(CDate(StartingTimeandDate), Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev)
             EventRates(Lk).EventMins = EventMins
             EventMinsTotal = EventMinsTotal + EventRates(Lk).EventMins
             If AllMenuItems(MItem).TimePrice2 > 0 Then
               EventRates(Lk).EventRate = (Val(AllMenuItems(MItem).TimePrice2 / 60) * EventMins)
             End If
          End If
         End If
      End If
      Case Is = 2
      If AllMenuItems(MItem).TimeEvent3 = Trim(Time_Events(Cev).EName) Then
         If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
         
         Else
            If CheckDate(Cev) Then
              EventMins = EventMinsBetween(CDate(StartingTimeandDate), Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev)
                EventRates(Lk).EventMins = EventMins
                EventMinsTotal = EventMinsTotal + EventRates(Lk).EventMins
                If AllMenuItems(MItem).TimePrice3 > 0 Then
                 EventRates(Lk).EventRate = Val(AllMenuItems(MItem).TimePrice3 / 60) * EventMins
                End If
            End If
         End If
      End If
      Case Is = 3
      If AllMenuItems(MItem).TimeEvent4 = Trim(Time_Events(Cev).EName) Then
         If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
         
         Else
         If CheckDate(Cev) Then
             EventMins = EventMinsBetween(CDate(StartingTimeandDate), Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev)
             EventRates(Lk).EventMins = EventMins
             EventMinsTotal = EventMinsTotal + EventRates(Lk).EventMins
             If AllMenuItems(MItem).TimePrice4 > 0 Then
              EventRates(Lk).EventRate = Val(AllMenuItems(MItem).TimePrice4 / 60) * EventMins
             End If
         End If
         End If
      End If
      Case Is = 4
      If AllMenuItems(MItem).TimeEvent5 = Trim(Time_Events(Cev).EName) Then
         If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
         
         Else
            If CheckDate(Cev) Then
                EventMins = EventMinsBetween(CDate(StartingTimeandDate), Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev)
                EventRates(Lk).EventMins = EventMins
                EventMinsTotal = EventMinsTotal + EventRates(Lk).EventMins
                If AllMenuItems(MItem).TimePrice5 > 0 Then
                 EventRates(Lk).EventRate = Val(AllMenuItems(MItem).TimePrice5 / 60) * EventMins
                End If
            End If
         End If
      End If
    End Select
  Next Cev
Next Lk
Dim EventRateTotal As Currency
EventRateTotal = 0
For Cev = 0 To 4
   EventRateTotal = EventRateTotal + EventRates(Cev).EventRate
Next Cev
ReturnStr = MItem & "(Mins Open " & CStr(OpenMinsTotal) & ")"
If AllMenuItems(MItem).Avalibility = 0 Then
  OpenMinsTotal = EventMinsTotal
Else
  OpenMinsTotal = OpenMinsTotal - EventMinsTotal
End If
'If OpenMinsTotal = 0 Then OpenMinsTotal = 1
  If AllMenuItems(MItem).RegularPrice > 0 And OpenMinsTotal > 0 And AllMenuItems(MItem).Avalibility <> 0 Then
    TimedRatesPrice = (OpenMinsTotal * (AllMenuItems(MItem).RegularPrice / 60)) + EventRateTotal
  Else
    TimedRatesPrice = EventRateTotal
  End If
  If SaleRounding = ".25" Then
     TimedRatesPrice = Format(RoundToQuarter(TimedRatesPrice), "0.00")
  End If
Exit Function
erHand:
Error_Message "Timed Rates", Err.Description
Exit Function
End Function
Public Function IsValidTime(ETime As Date, StartTime As String, EndTime As String, EventNo As Integer) As Boolean
Dim EndEventDate As Date
Dim StartEventDate As Date
IsValidTime = False
If DateDiff("n", StartTime, EndTime) > 0 Then
  EndEventDate = Date:  StartEventDate = Date
  GoTo de   'DO EVENT
End If
If DateDiff("n", StartTime, ETime) < 0 And DateDiff("n", EndTime, ETime) > 0 Then
  Exit Function
End If
If DateDiff("n", StartTime, ETime) > 0 And DateDiff("n", EndTime, ETime) < 0 Then
  Exit Function
End If
If DateDiff("n", StartTime, ETime) < 0 And DateDiff("n", EndTime, ETime) < 0 Then
  EndEventDate = Date: StartEventDate = Date - 1
End If
If DateDiff("n", StartTime, ETime) > 0 And DateDiff("n", EndTime, ETime) > 0 Then
  EndEventDate = Date + 1: StartEventDate = Date
End If
de:
If EventNo = -1 Then
  If ETime + Date <= TimeValue(EndTime) + EndEventDate Then
    If ETime + Date >= TimeValue(StartTime) + StartEventDate Then
         IsValidTime = True
         Exit Function
    End If
  End If
Else
  If ETime + Date <= TimeValue(EndTime) + EndEventDate Then
    If ETime + Date >= TimeValue(StartTime) + StartEventDate Then
       If CheckDayEvent(StartEventDate, EventNo) Then
         IsValidTime = True
         Exit Function
       End If
    End If
  End If
End If
End Function
Public Function EventMinsBetween(TimeStart As Date, StartTime As String, EndTime As String, EventNo As Integer) As Long
Dim EndEventDate As Date
Dim StartEventDate As Date
EventMinsBetween = 0
If DateDiff("n", StartTime, EndTime) > 0 Then
  EndEventDate = Date:  StartEventDate = Date
  GoTo de   'DO EVENT
End If
If DateDiff("n", StartTime, Time) < 0 And DateDiff("n", EndTime, Time) < 0 Then
  EndEventDate = Date: StartEventDate = Date - 1
End If
If DateDiff("n", StartTime, Time) > 0 And DateDiff("n", EndTime, Time) > 0 Then
  EndEventDate = Date + 1: StartEventDate = Date
End If
de:
If CheckDayEvent(StartEventDate, EventNo) Then
  If TimeStart <= TimeValue(StartTime) + StartEventDate Then
      If Time + Date >= EndTime + EndEventDate Then
        EventMinsBetween = DateDiff("n", StartTime + StartEventDate, EndTime + EndEventDate)
      Else
        EventMinsBetween = DateDiff("n", StartTime + StartEventDate, Time + Date)
      End If
      If EventMinsBetween < 0 Then EventMinsBetween = 0
      Exit Function
  ElseIf TimeStart >= TimeValue(StartTime) + StartEventDate Then
      If Time + Date >= EndTime + EndEventDate Then
        EventMinsBetween = DateDiff("n", TimeStart, EndTime + EndEventDate)
      Else
        EventMinsBetween = DateDiff("n", TimeStart, Time + Date)
      End If
      If EventMinsBetween < 0 Then EventMinsBetween = 0
      Exit Function
  End If
End If
End Function
Public Function GetModShiftPrice(ModifierCategory As String, ModName As String, PreFix As String) As Currency
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
If PriceShiftItems Is Nothing Then
    PreFix = ""
    GetModShiftPrice = 0
    Exit Function
End If
If Not PriceShiftItems.Item(ModifierCategory & ModName & (ModShiftPLevel - 1)) Is Nothing Then '-1
   PreFix = PriceShiftItems.Item(ModifierCategory & ModName & (ModShiftPLevel - 1)).AliasName
   GetModShiftPrice = PriceShiftItems.Item(ModifierCategory & ModName & (ModShiftPLevel - 1)).Price
Else
   PreFix = ""
   GetModShiftPrice = 0  '-1
End If
Exit Function
ERH:
Err.Clear
End Function
Public Function CheckGiftCardsForTip(ckNo As Long) As Boolean
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
CheckGiftCardsForTip = False
If CPath.TerminalFailed Then
  Exit Function
End If
Sql = "Select * from [HouseCharges] Where [Day of Charge] = 0 and [voided] = 0 and [Check Number] = " & ckNo

If Check_Option(165) = 1 Then
  Sql = Sql & " and [location Name] = '" & LocationReset & "'"
End If
  'only allow void from location check was entered
'added 07/27/2020 filter for multi store gift cards

If ConnectADODB(HouseAcct, Db) = False Then
 Exit Function
End If
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
If Not RS.EOF Then CheckGiftCardsForTip = True
DbClose Db, RS
End Function
Public Function IsPastEndOfDay(PickUpDate As String, PickUpTime As String) As Boolean
Dim EndTime As String
Dim OutofRange As Boolean
If Not IsDate(PickUpDate) Then PickUpDate = Date
If Not IsDate(PickUpTime) Then PickUpTime = Time
If Not IsDate(Input_Option(84)) Then   '// end of business day  -- default is 4:00 am
  EndTime = "4:00 am"
Else
  EndTime = Input_Option(84)
End If
IsPastEndOfDay = CDate(PickUpDate & " " & PickUpTime) > CDate(SetEndingDate(GetBusinessDay) & " " & EndTime)
End Function
Public Function GetShiftEvent(EventName As String, RegularPrice As Currency, EventPrice As Currency, EventPFx As String) As Currency
On Error GoTo erHand
Dim ETime As Date
Dim Lk As Integer
Dim Cev As Integer
ETime = Time
If EventName = "" Or EventName = "None" Then
  GetShiftEvent = RegularPrice
  Exit Function
End If
If Check_Option(222) = 1 Then
 If OrderIs = GetOrderType(GoSale) Then
   GetShiftEvent = RegularPrice
   Exit Function   'Mod for Richard 08/26/2015
 End If
End If  ' disable timed events for go orders

If EventName = "Follow Main Events" Then
    GetShiftEvent = TimedPrice(CStr(GetMenu(0)), RegularPrice, EventName)
    If GetShiftEvent = 0 Then GetShiftEvent = EventPrice
    Exit Function
Else
              For Cev = 0 To UBound(Time_Events(), 1)
                       If EventName = Trim(Time_Events(Cev).EName) Then
                          If Val(GetSetting("Micro$ale", "Register Setup", Trim(Time_Events(Cev).EName))) = 1 Then
                              Order.Information.Caption = "Shift Timed Price Disabled"
                          Else
                           If CheckDate(Cev) And IsValidTime(ETime, Time_Events(Cev).StartTime, Time_Events(Cev).EndTime, Cev) Then
                             If Time_Events(Cev).PriceAdjust <> 0 Then
                               GetShiftEvent = RegularPrice - Time_Events(Cev).PriceAdjust
                             Else
                               GetShiftEvent = Val(EventPrice)
                               If GetShiftEvent = 0 Then GetShiftEvent = RegularPrice
                             End If
                             If Check_Option(109) = 1 Then
                               EventPFx = EventPFx & Left(Trim(Time_Events(Cev).EName), 5) & " "
                             End If
                             Exit Function
                           End If
                          End If
                       End If
              Next Cev
GetShiftEvent = RegularPrice
End If
Exit Function
erHand:
Error_Message "Get Shift Events", Err.Description
Resume Next
End Function
Public Function GetNonSalesRev(StartDate As String, EndDate As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetNonSalesRev = 0
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Db) = False Then Exit Function
  Sql = "Select Distinct Sum([Exempt Sales]) as TeX, Sum([Charged Gratuity]) as TGt from [Sales] Where "
  Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  ConnectADORS (Sql), Db, RS
   If Not RS.EOF Then
     GetNonSalesRev = Format(Val(RS![TEX] & "") - Val(RS![TGt] & ""), "0.00")
   End If
DbClose Db, RS
End Function
Public Function GetExSales(PathtoFile, SelectDay As String, ChkNO As Long, SvrName As String)
GetExSales = GetExemptSales(CStr(PathtoFile), SelectDay, ChkNO, SvrName, 0)
End Function
Public Function BuildSql_Table(TableName As String, Db As ADODB.Connection, CTR As Integer) As String
        BuildSql_Table = ""
        Dim RS As ADODB.Recordset
        Dim GA As Integer
        If ConnectADORS(TableName, Db, RS) = False Then Exit Function
        BuildSql_Table = "Insert Into " & TableName & " ("
        For GA = 0 To RS.fields.count - CTR
            If GA = RS.fields.count - CTR Then
                BuildSql_Table = BuildSql_Table & "[" & RS.fields.Item(GA).Name & "])"
                Exit For
            Else
                BuildSql_Table = BuildSql_Table & "[" & RS.fields.Item(GA).Name & "],"
            End If
        Next
        SetNothing RS
    End Function
Public Sub ExportOrderHistory(CName As String, CPhone As String, iDb As ADODB.Connection)
On Error GoTo ErrorHandler
If Check_Option(85) = 1 Then Exit Sub
Dim eRs     As ADODB.Recordset
Dim Er      As Integer
Dim DString As String
Dim FLD     As Integer
  Sql = "Delete from [Order History] where [Customer Name] = '" & CName & "' and [Customer Phone] = '" & CPhone & "'"
  iDb.Execute Sql
  DString = BuildSql_Table("[Order History]", iDb, 4)
   Sql = "Select * From [Check Detail] Where [Customer Name] = '" & CName & "'" & " and [Customer phone] = '" & CPhone & "'"
  ConnectADORS Sql, iDb, eRs
   Do While Not eRs.EOF
     Sql = DString & " Values ("
       For FLD = 0 To eRs.fields.count - 1
          Select Case eRs.fields(FLD).Type
          Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTimeStamp, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adBSTR, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adVarWChar, ADODB.DataTypeEnum.adWChar
            If FLD = eRs.fields.count - 4 Then
               Sql = Sql & "'" & eRs.fields(FLD).Value & "')"
               Exit For
            Else
               Sql = Sql & "'" & eRs.fields(FLD).Value & "',"
            End If
         Case Else
            If FLD = eRs.fields.count - 4 Then
              Sql = Sql & Val(eRs.fields(FLD).Value) & ")"
              Exit For
            Else
              Sql = Sql & Val(eRs.fields(FLD).Value) & ","
            End If
         End Select
      Next FLD
     eRs.MoveNext
     iDb.Execute Sql
   Loop
   SetNothing eRs
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Export Order History " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub SaveOrderHistory(CustomerName As String, CustomerPhone As String, PhoneOrderNumber As Long, OrderNumber As Long, PHDB As ADODB.Connection)
If Check_Option(85) = 1 Then Exit Sub '// turns off order history
'Dim OHDB As ADODB.Connection
Dim OHRS As ADODB.Recordset
Dim Db   As ADODB.Connection
Dim RS   As ADODB.Recordset
  Sql = "Delete from [Order History] where [Customer Name] = '" & CustomerName & "' and [Customer Phone] = '" & CustomerPhone & "'"
 PHDB.Execute (Sql)
 Sql = "Select * from [Check Detail] Where [Check Number] = " & PhoneOrderNumber
 Sql = Sql & " Order by [POS] asc"
 ConnectADORS Sql, PHDB, RS
 ConnectADORS "Select * from [Order History]", PHDB, OHRS
  Do While Not RS.EOF
    OHRS.AddNew
    OHRS.fields("Check Number").Value = OrderNumber
    OHRS.fields("Check Date").Value = RS.fields("Check Date").Value
    OHRS.fields("Quantity").Value = RS.fields("Quantity").Value
    OHRS.fields("Menu Item").Value = RS.fields("Menu Item").Value
    OHRS.fields("Price").Value = RS.fields("Price").Value
    OHRS.fields("Selection Number").Value = RS.fields("Selection Number").Value
    OHRS.fields("POS").Value = RS.fields("POS").Value
    OHRS.fields("Record Type").Value = RS.fields("Record Type").Value
    OHRS.fields("Menu Number").Value = RS.fields("Menu Number").Value
    OHRS.Update
    RS.MoveNext
  Loop
  RS.Close
  OHRS.Close
End Sub

Public Sub SavePhoneOrderCheck(CName As String, CPhone As String, CallInOrderNo As Long, RegularCheckNo As Long, CloserName As String)
On Error GoTo ErrorHandler
Dim Cdb As ADODB.Connection
Dim sDb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim SRS As ADODB.Recordset
SavePhoneOrderDetail CName, CPhone, CallInOrderNo, RegularCheckNo, CloserName
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Cdb) = False Then Exit Sub
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, sDb) = False Then Exit Sub
   Sql = "Select * From [Open Checks] Where [Check Number] = " & CallInOrderNo & " and [Customer Name] = '" & CName & "' and [Customer Phone] = '" & CPhone & "'"
   ConnectADORS Sql, Cdb, CRS
      Sql = "Select * from [Closed Checks]"
   ConnectADORS Sql, sDb, SRS
               If Not CRS.EOF Then
                   SRS.AddNew
                   SRS![Check Number] = RegularCheckNo
                   SRS![Server Name] = CloserName  '  CRS![Server Name]
                   '// need to use the drivers name if your not assigning orders.
                   'SRS![Table Number] = "Call In-" & Str(CRS![Check Number])
                   SRS![Table Number] = CRS![Customer Name] & "-" & Str(CRS![Check Number])
                   SRS![Number in Party] = CRS![Number in Party]
                   SRS![Open Time] = CRS![check Opened]
                   SRS![Close Time] = CStr(Time)
                   SRS![Gross Sales] = CRS![Gross Sales]
                   SRS![Gross Sales Tax] = CRS![Gross Sales Tax]
                   SRS![Gross Special Tax] = CRS![Gross Special Tax]
                   SRS![Gross Liquor Tax] = CRS![Gross Liquor Tax]
                   SRS![Gross Discounts] = CRS![Gross Discounts]
                   SRS![Voided Sales] = CRS![Voided Sales]
                   SRS![Voided Sales Tax] = CRS![Voided Sales Tax]
                   SRS![Voided Special tax] = CRS![Voided Special tax]
                   SRS![Voided liquor Tax] = CRS![Voided liquor Tax]
                   SRS![Voided Discounts] = CRS![Voided Discounts]
                   SRS![Transfer From] = CRS![Transfer From]
                   SRS![Actual Cust Count] = CRS![Actual Cust Count]
                   SRS![Order Destination] = CRS![Order Destination]
                   SRS![Order Date] = Format(Date, "MM/DD/YYYY") ' Format(CRS![Order Date], "MM/DD/YYYY")
                   SRS![Register Name] = MyRegisterName
                   SRS![Close out day] = CRS![Close out day]
                   SRS![Gross Exempt] = Val(CRS![Gross Exempt] & "")
                   SRS![Voided Exempt] = Val(CRS![Voided Exempt] & "")
                   SRS![Check Total] = Val(CRS![Check Total] & "")
                   SRS![Inclusive Total] = Val(CRS![Inclusive Total] & "")
                   SRS.fields("Seat Number").Value = 0
                   SRS![Payments] = Val(CRS![Payments] & "")
                   SRS![Cashier Name] = ReturnMyCashier
                   If CPhone > "" Then
                     SRS![Member Name] = CPhone
                   Else
                     SRS![Member Name] = ""
                   End If
                 SRS.Update
              End If
       DbClose Cdb, CRS
       DbClose sDb, SRS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Close Call In " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub SavePhoneOrderDetail(CName As String, CPhone As String, CallInOrderNo As Long, RegularCheckNo As Long, CloserName As String)
'Public Sub SavePhoneOrderDetail(CName As String, CPhone As String, CallInOrderNo As Long, RegularCheckNo As Long, CloserName As String)

On Error GoTo ErrorHandler
Dim Cdb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim ORs As ADODB.Recordset
Dim oDb As ADODB.Connection
Dim Er  As Integer
     If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Cdb) = False Then Exit Sub
     Sql = "Select * From [Check Detail] Where [Customer Name] = '" & CName & "'" & " and [Customer phone] = '" & CPhone & "'"
     Sql = Sql & " and [Check Number] = " & CallInOrderNo
     ConnectADORS Sql, Cdb, CRS
     If Not CRS.EOF Then
       If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, oDb) = False Then Exit Sub
       Sql = "Update [Check Detail] set Deleted = 1 Where [Menu Item] In ('Online Tip','Online Fee','Online Delivery')"
       oDb.Execute Sql
       
       ConnectADORS "Select * from [Check Detail]", oDb, ORs
     End If
     Do While Not CRS.EOF
        ORs.AddNew
           For Er = 0 To ORs.fields.count - 1
              Select Case UCase(ORs.fields(Er).Name)
                Case Is = "SERVERS NAME"
                   ORs.fields(Er).Value = CloserName
                Case Is = "CASHIER NAME"
                   ORs.fields(Er).Value = ReturnMyCashier
                Case Is = "REGISTER NAME"
                   ORs.fields(Er).Value = MyRegisterName
                Case Is = "CHECK NUMBER"
                   ORs.fields(Er).Value = RegularCheckNo
                Case Is = "ORDER TYPE"
                   ORs.fields(Er).Value = SvrPath.TerminalGroupName
                Case Is = "CHECK DATE"
                   ORs.fields(Er).Value = Format(Date, "MM/DD/YYYY")
                Case Else
                   ORs.fields(Er).Value = CRS.fields(ORs.fields(Er).Name).Value
              End Select
            Next Er
           ORs.Update
         CRS.MoveNext
       Loop
DbClose oDb, ORs
DbClose Cdb, CRS
Exit Sub
ErrorHandler:
If Err.Number = 3265 Then Resume Next
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Export Order History " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function TrayNamesAreOn() As Boolean
TrayNamesAreOn = False
If Not AllMenuItems("Tray Names") Is Nothing Then
  TrayNamesAreOn = True
  Exit Function
End If
If Not AllMenuItems("Seat Names") Is Nothing Then
  TrayNamesAreOn = True
End If
End Function
Public Function GetCustNameBySeat(DataIn As String, SeatNo As Integer, MyCheckNo As Long) As String
On Error GoTo ERH
GetCustNameBySeat = DataIn
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
 Sql = "Select * from [Seat Names] Where [Check Number] = " & MyCheckNo  'AllOrders.Item(1).PickUpNo
 Sql = Sql & " and [Seat Number] = " & SeatNo
ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   GetCustNameBySeat = Tb![Customer Name] & ""
 End If
ERH:
DbClose Db, Tb
End Function
Public Function GetSeatName(Db As ADODB.Connection, CallIn_Check_No As Long, SeatNames() As String) As Boolean
On Error GoTo ErrorHandler
       If Db Is Nothing Then Exit Function
       GetSeatName = False
       Dim RS  As ADODB.Recordset
       Dim R   As Integer
       Sql = "Select * from [Seat Names] Where [Check Number] = " & CallIn_Check_No
       Sql = Sql & " Order by [Seat Number] asc"
       If ConnectADORS(Sql, Db, RS) = False Then Exit Function
       If Not RS.EOF Then
        GetSeatName = True
        R = -1
        Do While Not RS.EOF
         R = R + 1
         SeatNames(R) = RS.fields("Customer Name").Value & ""
         RS.MoveNext
        Loop
       End If
       SetNothing RS
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Seat Names " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub RemoveSeatNames(CallIn_Check_No As Long)
On Error GoTo ErrorHandler
       Dim Db  As ADODB.Connection
       If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Sub
       Sql = "Delete from [Seat Names] Where [Check Number] = " & CallIn_Check_No
       Db.Execute Sql
       SetNothing Db
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Seat Names " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub AddSeatName(CallIn_Check_No As Long, CustName As String, SeatNo As Long)
On Error GoTo ErrorHandler
       Dim sDb As ADODB.Connection
       Dim RS  As ADODB.Recordset
       If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, sDb) = False Then Exit Sub
       Sql = "Select * from [Seat Names] Where [Check Number] = " & CallIn_Check_No
       Sql = Sql & " and [Seat Number] = " & SeatNo
       If ConnectADORS(Sql, sDb, RS) = False Then Exit Sub
       If RS.EOF Then RS.AddNew
       RS.fields("Customer Name").Value = CustName
       RS.fields("Check Number").Value = CallIn_Check_No
       RS.fields("Seat Number").Value = SeatNo
       RS.Update
       DbClose sDb, RS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Seat Names " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function UpdateCloseAllTotals(NewCheckNo As Long) As String
If Gross_Sales = 0 Then Exit Function
Dim Db  As ADODB.Connection
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Rst As String
    Sql = "Update [Closed Checks] Set [Gross Special Tax] = " & (Chk_Sales_data(2) + Chk_Sales_data(20)) & ", [Gross Sales Tax] = " & Abs(Chk_Sales_data(1) + Chk_Sales_data(12)) & ", [Check Total] = " & Chk_Sales_data(5) & ", [Gross Sales] = " & Gross_Sales & " ,[Gross Exempt] =  " & (Chk_Sales_data(13))
    Sql = Sql & " Where [close out day] = 0 and [check number] = " & NewCheckNo '& " and [server name] = '" & LoggedOnAs(0).EMPName & "'"
    Rst = GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CheckStatDb))
Set GRS = Nothing
UpdateCloseAllTotals = Rst
End Function

Public Sub RemoveCallinOrder(CallIn_Check_No As Long, CustName As String, CustPhone As String)
On Error GoTo ErrorHandler
       Dim sDb As ADODB.Connection
       If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, sDb) = False Then Exit Sub
       'SaveOrderHistory CustName, CustPhone, OrderNumber, CallIn_Check_No, sDb
       ExportOrderHistory CustName, CustPhone, sDb
       Sql = "Delete from [check Detail] where [Check Number] = " & CallIn_Check_No
       sDb.Execute Sql
       Sql = "Delete from [Open Checks] where [Check Number] = " & CallIn_Check_No
       sDb.Execute Sql
       Sql = "Delete from CustomerNotification where CheckNumber = '" & CStr(CallIn_Check_No) & "'"
       sDb.Execute Sql
       SetNothing sDb
       UnlockedCheckNumber CallIn_Check_No
       Trans_Log "Phone Order Removed " & CustName & " " & CallIn_Check_No
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Close Call In Order " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function OnTheClock(sPI As String) As Boolean
On Error Resume Next
TAgain:
Dim getnewCknum As ADODB.Connection
Dim Records As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", getnewCknum) = False Then Exit Function
OnTheClock = False
If SvrPath.GroupAudit Then
 LoggedOnAs(0).EMPName = LoggedOnAs(0).EmpAssignedTo
End If
 Sql = "Select * from [Check Numbers] Where [Servers Name] = '" & LoggedOnAs(0).EMPName & "'"
 Sql = Sql & " and [Audit Status] = 0"
ConnectADORS Sql, getnewCknum, Records

If Records.EOF Then
    If SvrPath.GroupAudit Then
       If LoggedOnAs(0).EMPName = "None" Then
          See_Msg "You are not part of this group", 0, "Punch In is required to ring orders.. "
          Exit Function
       End If
    End If
    If Check_Option(105) = 0 Then
        See_Msg LoggedOnAs(0).EMPName & " has no audit open", 0, "Audit is required to ring orders.. "
     Else
        If See_Msg(LoggedOnAs(0).EMPName & " has no audit open, would you like to start an Audit ?", 4, "Audit is required to ring orders.. ") = 6 Then
            DbClose getnewCknum, Records
            StartServerAudit LoggedOnAs(0).EMPName, False, LoggedOnAs(0).CurrentJobCode, 0
            GoTo TAgain
        End If
    End If
    OnTheClock = False
Else
 If Not ALLRevCenter.Item(Records.fields("Register Group").Value & "") Is Nothing Then
  LoggedOnAs(0).MyRegisterGroup = Records.fields("Register Group").Value & ""
 Else
  LoggedOnAs(0).MyRegisterGroup = ""
 End If
 OnTheClock = True
End If
DbClose getnewCknum, Records
End Function
Public Function GetBankDeposits() As Currency
If GetConfigInfo("Misc", "End Of Day Deposits Are Off", "Yes", "FUNCTIONS.INI") = "ENABLED" Then GetBankDeposits = 0: Exit Function
On Error GoTo ErrorHandler
Dim GL As GeneralLib.Utilities
Set GL = New GeneralLib.Utilities
  GetBankDeposits = Val(GL.GetEndofDayCashFromAudits)
Set GL = Nothing
Exit Function
ErrorHandler:
GetBankDeposits = 0
End Function
Public Function ClosingProgess() As Boolean
ClosingProgess = FileCheck(SvrPath.TerminalPath & "DAT Files\" & "DailyClosing.dat")
End Function
Public Sub GetSetAccess(GetManagerName As String)
On Error GoTo ERH
TAgain:
If Check_Option(161) = 0 Then Exit Sub
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
Dim getnewCknum As ADODB.Connection
Dim Records As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & TimeRecordsDb, getnewCknum) = False Then Exit Sub
Sql = "Select * from [Employee Records] Where [Employee Name] = '" & GetManagerName & "'"
Sql = Sql & " and [Time in] > '' and [Time Out] = ''"
ConnectADORS Sql, getnewCknum, Records
If Not Records.EOF Then
 LoggedOnAs(0).EmpAccessLevel = Access_Required(Records![Department Name])
End If
DbClose getnewCknum, Records
Exit Sub
ERH:
Err.Clear
End Sub
Public Function ReturnServerAudit(RegisterNo As Long) As Long
On Error GoTo ErrorHandler
Dim NextNoAv       As Long
Dim LastNumberUsed As Long
Dim getnewCknum    As ADODB.Connection
Dim Records        As ADODB.Recordset
Dim CA             As Integer
Dim ShiftNo        As Long
retry = 0
ReturnServerAudit = RegisterNo
ConnectADODB SvrPath.TerminalPath & "MSaleData", getnewCknum
tryagain:
    Sql = "Select * from [Check Numbers] where cast(([Check No] / 1000) as int) = " & RegisterNo / 1000
  ConnectADORS Sql, getnewCknum, Records
      If Records.EOF Then
       Records.AddNew
       Records.fields("Check No").Value = RegisterNo
      Else
       If Records![Register Name] = MyRegisterName And Records.fields("Servers Name").Value = GetActiveCashier Then
         Records.fields("Audit Status").Value = 0
         Records.Update
         DbClose getnewCknum, Records
         Exit Function
       Else
         Records.AddNew
         Records.fields("Check No").Value = Assign_Check_No(GetActiveCashier, False)
       End If
      End If
      Records.fields("Servers Name").Value = GetActiveCashier
      Records.fields("Time").Value = Format(Str(Time), "HH:MM ampm")
      Records.fields("Start Date").Value = Format(Date, "MM/DD/YY")
      Records.fields("Audit Status").Value = 0
      Records![In Use] = 0
      Records![Register Name] = MyRegisterName
      Records![Register Group] = ""
      Records![Drawer No] = 0
      Records![Work Area] = ""
      Records![Starting Bank] = 0
      Records![Deposits Paid In] = 0
      NextNoAv = Records.fields("Check No").Value
      Records.Update
      DbClose getnewCknum, Records
      ReturnServerAudit = NextNoAv
      Sql = "Select * from [Check Numbers] where [Check No] = " & NextNoAv
      Sql = Sql & " and [Register Name] <> '" & MyRegisterName & "'"
      ConnectADODB SvrPath.TerminalPath & "MSaleData", getnewCknum
      ConnectADORS Sql, getnewCknum, Records
      If Not Records.EOF Then
         SetNothing Records
         Sql = "Delete from [Check Numbers] where [Register Name] = '" & MyRegisterName & "'"
         getnewCknum.Execute Sql
         DbClose getnewCknum, Records
         GoTo tryagain
      End If
     ' Set UDA = Nothing
  DbClose getnewCknum, Records
Exit Function
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Start Server Audit ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function

Public Function GetDayToNumber(DayNo As Integer) As Integer
If DayNo = 0 Then GetDayToNumber = 0: Exit Function
If Trim$(Input_Option(6)) = "Monday" Then
  GetDayToNumber = DayNo
ElseIf Trim(Input_Option(6)) = "Tuesday" Then
  If DayNo = 1 Then GetDayToNumber = 7
  If DayNo = 2 Then GetDayToNumber = 1
  If DayNo = 3 Then GetDayToNumber = 2
  If DayNo = 4 Then GetDayToNumber = 3
  If DayNo = 5 Then GetDayToNumber = 4
  If DayNo = 6 Then GetDayToNumber = 5
  If DayNo = 7 Then GetDayToNumber = 6
ElseIf Trim(Input_Option(6)) = "Wednesday" Then
  If DayNo = 1 Then GetDayToNumber = 6
  If DayNo = 2 Then GetDayToNumber = 7
  If DayNo = 3 Then GetDayToNumber = 1
  If DayNo = 4 Then GetDayToNumber = 2
  If DayNo = 5 Then GetDayToNumber = 3
  If DayNo = 6 Then GetDayToNumber = 4
  If DayNo = 7 Then GetDayToNumber = 5
ElseIf Trim(Input_Option(6)) = "Thursday" Then
  If DayNo = 1 Then GetDayToNumber = 5
  If DayNo = 2 Then GetDayToNumber = 6
  If DayNo = 3 Then GetDayToNumber = 7
  If DayNo = 4 Then GetDayToNumber = 1
  If DayNo = 5 Then GetDayToNumber = 2
  If DayNo = 6 Then GetDayToNumber = 3
  If DayNo = 7 Then GetDayToNumber = 4
ElseIf Trim(Input_Option(6)) = "Friday" Then
  If DayNo = 1 Then GetDayToNumber = 4
  If DayNo = 2 Then GetDayToNumber = 5
  If DayNo = 3 Then GetDayToNumber = 6
  If DayNo = 4 Then GetDayToNumber = 7
  If DayNo = 5 Then GetDayToNumber = 1
  If DayNo = 6 Then GetDayToNumber = 2
  If DayNo = 7 Then GetDayToNumber = 3
ElseIf Trim(Input_Option(6)) = "Saturday" Then
  If DayNo = 1 Then GetDayToNumber = 3
  If DayNo = 2 Then GetDayToNumber = 4
  If DayNo = 3 Then GetDayToNumber = 5
  If DayNo = 4 Then GetDayToNumber = 6
  If DayNo = 5 Then GetDayToNumber = 7
  If DayNo = 6 Then GetDayToNumber = 1
  If DayNo = 7 Then GetDayToNumber = 2
ElseIf Trim(Input_Option(6)) = "Sunday" Then
  If DayNo = 1 Then GetDayToNumber = 2
  If DayNo = 2 Then GetDayToNumber = 3
  If DayNo = 3 Then GetDayToNumber = 4
  If DayNo = 4 Then GetDayToNumber = 5
  If DayNo = 5 Then GetDayToNumber = 6
  If DayNo = 6 Then GetDayToNumber = 7
  If DayNo = 7 Then GetDayToNumber = 1
End If
Exit Function
End Function
Public Function WeekStartsOn() As Byte
Dim Day_OffSet As Byte
   Select Case Trim$(Input_Option(6))
        Case "Monday"
           Day_OffSet = vbMonday
         Case "Tuesday"
           Day_OffSet = vbTuesday
        Case "Wednesday"
           Day_OffSet = vbWednesday
        Case "Thursday"
           Day_OffSet = vbThursday
        Case "Friday"
           Day_OffSet = vbFriday
        Case "Saturday"
           Day_OffSet = vbSaturday
        Case "Sunday"
           Day_OffSet = vbSunday
        Case Else
            WeekStartsOn = vbMonday
      End Select
    WeekStartsOn = Day_OffSet
End Function
Public Function GetValidAcctNo(ByVal IncommingNo As String) As String
Dim VN As String
VN = GetValidationNo(GetLocationName(IncommingNo))
If VN = "True" Then
   GetValidAcctNo = IncommingNo
ElseIf VN = "False" Then
   GetValidAcctNo = "Not Valid"
   See_Msg "Cannot Charge ...", 5, "Card not accepted at this location"
Else
     GetValidAcctNo = IncommingNo
End If
End Function
Public Function EditGiftCards(ByVal CardNo As String, ByVal GiftAmount As Currency, ByVal AccountMemo As String, ByVal AccountType As String, ByVal AccountExpDate As String, ByVal CustName As String) As Boolean
On Error GoTo ErrorHandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim AccountName As String
Dim AccountNo   As String
If CardNo = "" Then
  Trans_Log "No Account information. Amount = " & Format(GiftAmount, "0.00")
  Exit Function
End If
If AccountType = "Membership" Then
  With Assign_Name
      OkToDo = True  'turn off all account load
     .AcctType = "Membership"
     .AcctType.Enabled = False
     .SSCommand15.Enabled = False
     .SSCommand14.Enabled = False
     .Label1(18).Enabled = False
     .UserPrvilage True
     .Scan_Data(3).Text = CardNo
     .AcctType = "Membership"
     .GetAccount
     .AddToExpiration AccountExpDate
     .AExpiration.Enabled = False
  End With
  Assign_Name.Show 1
  OkToDo = False
  Exit Function
 End If
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Function
End If

Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & CardNo & "'"
Sql = Sql & " or [Scan Name] = '" & CardNo & "'"
Sql = Sql & " Or [Spc Interest 1] = '" & CardNo & "' Or [Spc Interest 2] = '" & CardNo & "'"
Sql = Sql & " Or [Spc Interest 3] = '" & CardNo & "' Or [Spc Interest 4] = '" & CardNo & "'"

If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
If Tb.EOF Then
 Tb.AddNew
   If CustName > "" Then
    Tb(0) = CustName
    AccountName = CustName
   Else
    Tb(0) = "Cert. #" & CardNo
    AccountName = "Cert. #" & CardNo
   End If
    AccountNo = CardNo   'need account number for payments
    Tb(1) = 0
    Tb(2) = 0
    Tb![Acct Number] = CardNo
    Tb(4) = 0
    Tb(5) = 0
    Tb(6) = 0
    Tb![Credit Balance] = 1
    Tb![Address] = "Gift Card"
    Tb![City] = ""
    Tb![State] = ""
    Tb![Zip] = ""
    Tb![Phone] = ""
    Tb![Memo] = AccountMemo & ""
    Tb![Account Type] = "Gift Card"
    Tb.fields("Spc Interest 1").Value = ""
    Tb.fields("Spc Interest 2").Value = ""
    Tb.fields("Spc Interest 3").Value = ""
    Tb.fields("Spc Interest 4").Value = ""
    If IsDate(AccountExpDate) Then
      Tb![Account Expires On] = AccountExpDate
    Else
      Tb![Account Expires On] = "None"
    End If
 Tb.Update
Else
  AccountName = Tb(0) & ""
  AccountNo = Tb![Acct Number] & ""
End If
SetNothing Tb
Sql = "Select * from [HousePayments] Where [Account Name] = '" & AccountName & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
   Tb.AddNew
   Tb![Date] = CStr(Format(Date, "MM/DD/YYYY"))
   Tb![Amount] = GiftAmount
   Tb![Account Name] = AccountName
   Tb![Payment Type] = AccountType
   Tb![Location Name] = LocationReset
   Tb![Check Number] = Check_no
   Tb![Account Number] = AccountNo
   Tb.Update
 DbClose Db, Tb
 Trans_Log "House Account Update  Account Number " & AccountNo & " Payment Amount " & Format(GiftAmount, "0.00")
 Call See_Msg(AccountName, 5, "Account Updated ...")
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Description, Err.Number, "Issuse Gift Card " & Sql)
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetSpoolPrinter(UsePrinter As String) As String
On Error GoTo ERH
  If UsePrinter = "None" Or UsePrinter = "" Then GetSpoolPrinter = "None": Exit Function
  GetSpoolPrinter = AllPrinter.Item(UsePrinter).PrinterRate
  If UsePrinter = "Report" Then Exit Function
  Printer_Drivers UsePrinter
Exit Function
ERH:
GetSpoolPrinter = "None"
End Function

