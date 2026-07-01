VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsFlex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.3#0"; "3dabm9u.ocx"
Begin VB.Form AuditDeposit 
   BorderStyle     =   0  'None
   Caption         =   "Edit Audit Deposits"
   ClientHeight    =   10530
   ClientLeft      =   225
   ClientTop       =   690
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   Begin BTNENHLib4.BtnEnh SSCommand9 
      Height          =   1035
      Left            =   12240
      TabIndex        =   0
      Top             =   9240
      Width           =   2550
      _Version        =   589827
      _ExtentX        =   4498
      _ExtentY        =   1826
      _StockProps     =   66
      Caption         =   "Exit"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   4
      Picture         =   "C:\Icons\Exit.ico"
      PictureXOffset  =   15
      PictureYOffset  =   2
      BackColorContainer=   14737632
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      ShadowMode      =   3
      UserData        =   0.1
      dibPicture      =   "AuditDeposit.frx":0000
      textCaption     =   "AuditDeposit.frx":00F8
      textLT          =   "AuditDeposit.frx":0160
      textCT          =   "AuditDeposit.frx":0178
      textRT          =   "AuditDeposit.frx":0190
      textLM          =   "AuditDeposit.frx":01A8
      textRM          =   "AuditDeposit.frx":01C0
      textLB          =   "AuditDeposit.frx":01D8
      textCB          =   "AuditDeposit.frx":01F0
      textRB          =   "AuditDeposit.frx":0208
      colorBack       =   "AuditDeposit.frx":0220
      colorIntern     =   "AuditDeposit.frx":024A
      colorMO         =   "AuditDeposit.frx":0274
      colorFocus      =   "AuditDeposit.frx":029E
      colorDisabled   =   "AuditDeposit.frx":02C8
      colorPressed    =   "AuditDeposit.frx":02F2
      HollowFrame     =   -1  'True
      VistaColorGlossyEffectLower=   13677994
   End
   Begin VSFlex8DAOCtl.VSFlexGrid RegList 
      Height          =   8100
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   14085
      _cx             =   1993826572
      _cy             =   1993816015
      Appearance      =   2
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin BTNENHLib4.BtnEnh SSCommand14 
      Height          =   1005
      Left            =   8760
      TabIndex        =   3
      Top             =   9240
      Width           =   3450
      _Version        =   589827
      _ExtentX        =   6085
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Save Deposit"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   4
      Picture         =   "C:\Icons\MyCash.bmp"
      PictureXOffset  =   69
      PictureYOffset  =   -1
      BackColorContainer=   12648384
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      ShadowMode      =   3
      UserData        =   0.1
      dibPicture      =   "AuditDeposit.frx":031C
      textCaption     =   "AuditDeposit.frx":04A2
      textLT          =   "AuditDeposit.frx":051A
      textCT          =   "AuditDeposit.frx":0532
      textRT          =   "AuditDeposit.frx":054A
      textLM          =   "AuditDeposit.frx":0562
      textRM          =   "AuditDeposit.frx":057A
      textLB          =   "AuditDeposit.frx":0592
      textCB          =   "AuditDeposit.frx":05AA
      textRB          =   "AuditDeposit.frx":05C2
      colorBack       =   "AuditDeposit.frx":05DA
      colorIntern     =   "AuditDeposit.frx":0604
      colorMO         =   "AuditDeposit.frx":062E
      colorFocus      =   "AuditDeposit.frx":0658
      colorDisabled   =   "AuditDeposit.frx":0682
      colorPressed    =   "AuditDeposit.frx":06AC
      HollowFrame     =   -1  'True
      VistaColorGlossyEffectLower=   13677994
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Audits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   13890
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   8850
      Left            =   150
      Top             =   150
      Width           =   14745
   End
End
Attribute VB_Name = "AuditDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
On Error Resume Next
ResizeScreen1024 Me, True
CenterForm Me
Dim Cw As Integer
With RegList
   .Clear
   .Cols = 9
   .ColWidth(0) = 1000 * WidScale800x600
   .ColWidth(1) = 1000 * WidScale800x600
   .ColWidth(2) = 2000 * WidScale800x600
   .ColWidth(3) = 1000 * WidScale800x600
   .ColWidth(4) = 1000 * WidScale800x600
   .ColWidth(5) = 1000 * WidScale800x600
   .ColWidth(6) = 1000 * WidScale800x600
   .ColWidth(7) = 1000 * WidScale800x600
   .ColWidth(8) = 2000 * WidScale800x600
   .FixedRows = 1
   .TextMatrix(0, 0) = "Date"
   .TextMatrix(0, 1) = "Start Time"
   .TextMatrix(0, 2) = "Operator"
   .TextMatrix(0, 3) = "Non Cash"
   .TextMatrix(0, 4) = "Cash In Drw"
   .TextMatrix(0, 5) = "Skim"
   .TextMatrix(0, 6) = "Cash Deposit"
   .TextMatrix(0, 7) = "-/+"
   .TextMatrix(0, 8) = "Register"
   .Width = 100
   For Cw = 0 To 8
   .Width = .Width + .ColWidth(Cw)
   Next Cw
   .Width = .Width + 100 * WidScale800x600
End With
LoadAudits
End Sub
Public Sub LoadAudits()
On Error GoTo errorhandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
Dim DB As ADODB.Connection, Tb As ADODB.Recordset
Dim Adate As String
ConnectADODB CPath.TerminalPath & "Audits.mdb", "Audits.mdb", DB
If System_config = "Express" Then
        Sql = "Select * from [Audits] where [Close Out Day] < 1 order by TRY_CAST([Start time] AS datetime) asc"
    ConnectADORS Sql, DB, Tb
    With RegList
      .Rows = RecordsCount(Tb) + 1
      .Row = 0
      Do While Not Tb.EOF
       .Row = .Row + 1
       .TextMatrix(.Row, 0) = Format(Tb![Date], "MM/DD/YY")
       .TextMatrix(.Row, 1) = Format(Tb![Start Time], "HH:MM ampm")
        .TextMatrix(.Row, 2) = Tb![Operator] & ""
       .TextMatrix(.Row, 3) = Format(Tb![Non Cash Tenders] & "", "0.00")
       .TextMatrix(.Row, 4) = Format(Tb![Cash In Drawer] & "", "0.00")
       .TextMatrix(.Row, 5) = Format(Tb![Skim], "0.00")
       .TextMatrix(.Row, 6) = Format(Tb![Bank Deposit], "0.00")
       .TextMatrix(.Row, 8) = Format(Tb![Total To Account For] - Val(Tb![Paid Outs]), "0.00")
       .TextMatrix(.Row, 7) = Format((Val(Tb![Bank Deposit] & "") + Val(Tb![Skim] & "")) - Abs(Val(.TextMatrix(.Row, 4))), "0.00") '(Val(Tb![Non Cash Tenders] & "") +
       If Tb![Close out day] = -1 Then
        .TextMatrix(.Row, 8) = "On Hold"
       Else
        .TextMatrix(.Row, 8) = Tb.Fields("Register Name").Value
       End If
       Tb.MoveNext
       Loop
      End With
      DbClose DB, Tb
     Dim FA As Integer
     For FA = 0 To NetWork_Data
        ConnectADODB Server_Path(2, FA) & FinancialDb, FinancialDb, DB
              Sql = "Select * from [Audits]" 'Where TRY_CAST([Start Date] AS datetime) + TRY_CAST([start time] AS datetime) Between #" & CDate(Adate) & " " & GetStartTime & "#"
        ConnectADORS Sql, DB, Tb
        With RegList
        .Rows = .Rows + RecordsCount(Tb)
          Do While Not Tb.EOF
              .Row = .Row + 1
              .TextMatrix(.Row, 0) = Format(Tb![Start date], "MM/DD/YY")
              .TextMatrix(.Row, 1) = Format(Tb![Start Time] & "", "HH:MM ampm")
              .TextMatrix(.Row, 2) = Tb![Name] & ""
              .TextMatrix(.Row, 7) = "Open"
              .TextMatrix(.Row, 8) = Tb![Register Name] & ""
            Tb.MoveNext
           Loop
          End With
          DbClose DB, Tb
          If SvrPath.sServerIsOn = 1 And SvrPath.sTerminalAddress > "" Then Exit For
     Next FA
Else
  If Check_System_Access("Show Only Your Audit") >= LoggedOnAs(0).EmpAccessLevel Then
     Sql = "Select * from [Audits] where [Operator] = '" & LoggedOnAs(0).EMPName & "' and [Close Out Day] = 0 order by TRY_CAST([Start time] AS datetime) asc"
  Else
     Sql = "Select * from [Audits] where [Close Out Day] = 0 order by TRY_CAST([Start time] AS datetime) asc"
  End If
  ConnectADORS Sql, DB, Tb
    With RegList
      .Rows = RecordsCount(Tb) + 1
      .Row = 0
      Do While Not Tb.EOF
       .Row = .Row + 1
       .TextMatrix(.Row, 0) = Format(Tb![Date], "MM/DD/YY")
       .TextMatrix(.Row, 1) = Format(Tb![Start Time], "HH:MM ampm")
       .TextMatrix(.Row, 2) = Tb![Operator] & ""
       .TextMatrix(.Row, 3) = Format(Tb![Non Cash Tenders] & "", "0.00")
       .TextMatrix(.Row, 4) = Format(Tb![Cash In Drawer] & "", "0.00")
       .TextMatrix(.Row, 5) = Format(Tb![Skim], "0.00")
       .TextMatrix(.Row, 6) = Format(Tb![Bank Deposit], "0.00")
       .TextMatrix(.Row, 8) = Format(Tb![Total To Account For] - Val(Tb![Paid Outs]), "0.00")
       If Format(Tb![Cash In Drawer] & "", "0.00") < 0 Then
        .TextMatrix(.Row, 7) = "0.00"
       Else
        .TextMatrix(.Row, 7) = Format((Val(Tb![Bank Deposit] & "") + Val(Tb![Skim] & "")) - Abs(Val(.TextMatrix(.Row, 4))), "0.00") '(Val(Tb![Non Cash Tenders] & "") +
       End If
       .TextMatrix(.Row, 8) = Tb.Fields("Register Name").Value
       Tb.MoveNext
       Loop
      End With
      DbClose DB, Tb
End If
Exit Sub
errorhandler:
ErrorRespose = Main_Error_Handler(Err.Description, Err.Number, "Edit Audits")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub CalTotals()
Dim T As Integer
For T = 1 To RegList.Rows - 1
 With RegList
  If .TextMatrix(T, 8) = "" Or .TextMatrix(T, 7) = "Open" Then Exit Sub
  RegList.TextMatrix(T, 7) = Format((Val(.TextMatrix(T, 6)) + (Val(.TextMatrix(T, 5)))) - Abs(Val(.TextMatrix(T, 4))), "0.00") '(Val(.TextMatrix(T, 3)) +
 End With
Next T
End Sub
Private Sub RegList_Click()
If RegList.TextMatrix(RegList.Row, 7) = "Open" Then Exit Sub
Dim ReturnedValue As Variant
If CreditInUse = False Then
  If RegList.Col = 3 Then
   ReturnedValue = ShowNumberPad("Edit Tender", False, 0, 0)
   If Val(ReturnedValue) > 0 Then RegList.TextMatrix(RegList.Row, 3) = Format(ReturnedValue, "0.00")
  ElseIf RegList.Col = 4 Then
   ReturnedValue = ShowNumberPad("Edit Cash", False, 0, 0)
   If Val(ReturnedValue) > 0 Then RegList.TextMatrix(RegList.Row, 4) = Format(ReturnedValue, "0.00")
  Else
   ReturnedValue = ShowNumberPad("Enter Deposit", False, 0, 0)
   If Val(ReturnedValue) > 0 Then RegList.TextMatrix(RegList.Row, 6) = Format(ReturnedValue, "0.00")
  End If
Else
   ReturnedValue = ShowNumberPad("Enter Deposit", False, 0, 0)
   If Val(ReturnedValue) > 0 Then RegList.TextMatrix(RegList.Row, 6) = Format(ReturnedValue, "0.00")
End If
CalTotals
End Sub
Private Sub SSCommand14_Click()
Dim R As Integer
Dim DB As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB CPath.TerminalPath & "Audits.mdb", "Audits.mdb", DB
With RegList
    For R = 1 To .Rows - 1
        If .TextMatrix(R, 7) = "Open" Then Exit For
        If Not IsDate(.TextMatrix(R, 0)) Then Exit For
        Sql = "Select * from [Audits] where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(.TextMatrix(R, 0))), "yyyy-mm-dd hh:nn:ss") & "'"
        Sql = Sql & " and TRY_CAST([Start Time] AS datetime) = # " & CDate(.TextMatrix(R, 1)) & "#"
        Sql = Sql & " and [Operator] = '" & .TextMatrix(R, 2) & "'"
        Sql = Sql & " and [Register Name] = '" & .TextMatrix(R, 8) & "'"
       ConnectADORS Sql, DB, Tb
       If Not Tb.EOF Then
         If CreditInUse = False Then
         Tb![Non Cash Tenders] = Format(.TextMatrix(R, 3), "0.00")
         Tb![Cash In Drawer] = Format(.TextMatrix(R, 4), "0.00")
         End If
         Tb![Bank Deposit] = Format(.TextMatrix(R, 6), "0.00")
         Tb.Update
       End If
    Next R
  DbClose DB, Tb
  See_Msg "Saved.....", 5, "Changes Saved."
End With
Exit Sub
errorhandler:
ErrorRespose = Main_Error_Handler(Err.Description, Err.Number, "Save Audits")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand9_Click()
Call MemoryCheck(Me)
End Sub
