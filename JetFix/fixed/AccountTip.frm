VERSION 5.00
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form AccountTip 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Account Tip"
   ClientHeight    =   6675
   ClientLeft      =   3150
   ClientTop       =   1890
   ClientWidth     =   5445
   Icon            =   "AccountTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox HTip 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Index           =   6
      Left            =   3330
      TabIndex        =   17
      Top             =   2775
      Width           =   1709
   End
   Begin VB.TextBox HTip 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   5
      Left            =   2925
      TabIndex        =   13
      Top             =   4200
      Width           =   1965
   End
   Begin VB.TextBox HTip 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   365
      Index           =   4
      Left            =   3330
      TabIndex        =   12
      Top             =   3240
      Width           =   1709
   End
   Begin VB.TextBox HTip 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Index           =   3
      Left            =   3330
      TabIndex        =   11
      Top             =   2340
      Width           =   1709
   End
   Begin VB.TextBox HTip 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Index           =   2
      Left            =   3330
      TabIndex        =   10
      Top             =   1875
      Width           =   1709
   End
   Begin VB.TextBox HTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Index           =   1
      Left            =   3330
      TabIndex        =   9
      Top             =   1410
      Width           =   1709
   End
   Begin VB.TextBox HTip 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Index           =   0
      Left            =   1950
      TabIndex        =   7
      Top             =   960
      Width           =   3091
   End
   Begin VB.ListBox Charges 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7245
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   6765
   End
   Begin BTNENHLib4.BtnEnh Quit 
      Height          =   825
      Left            =   2880
      TabIndex        =   14
      Top             =   5400
      Width           =   2295
      _Version        =   589828
      _ExtentX        =   4048
      _ExtentY        =   1455
      _StockProps     =   66
      Caption         =   "Exit"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      Surface         =   15
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "AccountTip.frx":1CCA
      textLT          =   "AccountTip.frx":1D32
      textCT          =   "AccountTip.frx":1D4A
      textRT          =   "AccountTip.frx":1D62
      textLM          =   "AccountTip.frx":1D7A
      textRM          =   "AccountTip.frx":1D92
      textLB          =   "AccountTip.frx":1DAA
      textCB          =   "AccountTip.frx":1DC2
      textRB          =   "AccountTip.frx":1DDA
      colorBack       =   "AccountTip.frx":1DF2
      colorIntern     =   "AccountTip.frx":1E1C
      colorMO         =   "AccountTip.frx":1E46
      colorFocus      =   "AccountTip.frx":1E70
      colorDisabled   =   "AccountTip.frx":1E9A
      colorPressed    =   "AccountTip.frx":1EC4
      HollowFrame     =   -1  'True
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   2377800
      Office2007ColorNorthBottom=   2377800
      Office2007ColorSouthTop=   2245700
      Office2007ColorSouthBottom=   2245700
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   2377800
      Office2007ColorNorthBottomMouseOver=   2377800
      Office2007ColorSouthTopMouseOver=   2377800
      Office2007ColorSouthBottomMouseOver=   2377800
      Office2007ColorInnerBorderPressed=   16777215
      Office2007ColorOuterBorderPressed=   16777215
      Office2007ColorNorthTopPressed=   3236450
      Office2007ColorNorthBottomPressed=   3104350
      Office2007ColorSouthTopPressed=   2245700
      Office2007ColorSouthBottomPressed=   0
      Office2003ColorNorth=   128
      Office2003ColorSouth=   64
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   128
      Office2003ColorSouthMouseOver=   128
      Office2003ColorBorderMouseOver=   1812708
      Office2003ColorNorthPressed=   128
      Office2003ColorSouthPressed=   64
      Office2003ColorBorderPressed=   1812708
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand14 
      Height          =   825
      Left            =   240
      TabIndex        =   15
      Top             =   5400
      Width           =   2370
      _Version        =   589828
      _ExtentX        =   4180
      _ExtentY        =   1455
      _StockProps     =   66
      Caption         =   "Save Tip"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      Surface         =   15
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "AccountTip.frx":1EEE
      textLT          =   "AccountTip.frx":1F5E
      textCT          =   "AccountTip.frx":1F76
      textRT          =   "AccountTip.frx":1F8E
      textLM          =   "AccountTip.frx":1FA6
      textRM          =   "AccountTip.frx":1FBE
      textLB          =   "AccountTip.frx":1FD6
      textCB          =   "AccountTip.frx":1FEE
      textRB          =   "AccountTip.frx":2006
      colorBack       =   "AccountTip.frx":201E
      colorIntern     =   "AccountTip.frx":2048
      colorMO         =   "AccountTip.frx":2072
      colorFocus      =   "AccountTip.frx":209C
      colorDisabled   =   "AccountTip.frx":20C6
      colorPressed    =   "AccountTip.frx":20F0
      HollowFrame     =   -1  'True
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   2377800
      Office2007ColorNorthBottom=   2377800
      Office2007ColorSouthTop=   2245700
      Office2007ColorSouthBottom=   2245700
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   2377800
      Office2007ColorNorthBottomMouseOver=   2377800
      Office2007ColorSouthTopMouseOver=   2377800
      Office2007ColorSouthBottomMouseOver=   2377800
      Office2007ColorInnerBorderPressed=   16777215
      Office2007ColorOuterBorderPressed=   16777215
      Office2007ColorNorthTopPressed=   3236450
      Office2007ColorNorthBottomPressed=   3104350
      Office2007ColorSouthTopPressed=   2245700
      Office2007ColorSouthBottomPressed=   0
      Office2003ColorNorth=   3302400
      Office2003ColorSouth=   2642176
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   3302400
      Office2003ColorSouthMouseOver=   3302400
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   3302400
      Office2003ColorSouthPressed=   2642176
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VB.Label CreditBalanceOnly 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Balance Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Index           =   6
      Left            =   405
      TabIndex        =   16
      Top             =   3240
      Width           =   2850
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Tip Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   4
      Left            =   555
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   825
      Left            =   405
      Top             =   4050
      Width           =   4620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tip Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   405
      TabIndex        =   8
      Top             =   2805
      Width           =   2850
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   5
      Top             =   2340
      Width           =   2850
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   405
      TabIndex        =   4
      Top             =   1875
      Width           =   2850
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   405
      TabIndex        =   3
      Top             =   1410
      Width           =   2850
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      TabIndex        =   2
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "House / Gift Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   345
      Width           =   2475
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   4650
      Left            =   240
      Top             =   645
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   6300
      Left            =   105
      Top             =   180
      Width           =   5235
   End
End
Attribute VB_Name = "AccountTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'processes tips for house accounts and local gift cards.

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
            HTip(5).Text = HTip(6).Text
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
    Sql = Sql & " Order By TRY_CAST([Date] AS date) DESC"   ' adjust to your actual date column name
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    If Tb.EOF Then
        BeginBal = 0
    Else
        BeginBal = CDbl(Nz(Tb![Beginning Balance], 0))
    End If

    HTip(4).Text = Format(BeginBal + TotalCharges - Payments, "0.00")
    DbClose Db, Tb
End Sub
Private Sub Form_Load()
'' handles tips for house charges

ResizeScreen1024 Me
CenterForm Me
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
On Error GoTo ErrorHandler
If LoadCardData.GiftCheckNo > 0 Then
  LoadRecords LoadCardData.GiftCheckNo, LoadCardData.GiftCheckAcct
  If GetCreditBalance(HTip(0).Text & "") = 0 Then
    CreditBalanceOnly.Visible = False
  Else
    CreditBalanceOnly.Visible = True
  End If
  Charges.Visible = False
Else
  Sql = "Select * from [HouseCharges] Where [Day of Charge] = 0 and [voided] = 0"
  If ConnectADODB(HouseAcct, Db) = False Then
    Exit Sub
  End If
  If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
  Charges.Clear
   Do While Not RS.EOF
     Charges.AddItem "CHK NO " & RS![Check Number] & " Acct. " & Trim$(RS![Scan Number] & "") & "-"
     RS.MoveNext
   Loop
   DbClose Db, RS
   If Charges.ListCount = 0 Then
     See_Msg ReturnLabel("No Charges Found"), 5, ""
   End If
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Tip house charge ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub HTip_Click(Index As Integer)
Dim ReturnedBal As Currency
If Index = 5 Then
  HTip(5).Text = ShowNumberPad(ReturnLabel("Enter Tip Amount"), False, 0, 0)
  HTip(5).Text = PlaceDec(HTip(5).Text)
  ReturnedBal = GetAccountBalance(HTip(0).Text)
  If Val(HTip(5).Text) > ReturnedBal And CreditBalanceOnly.Visible Then
    See_Msg "Over Account Balance " & Format(ReturnedBal, "0.00"), 0, "Cannot Accept."
    HTip(5).Text = 0
  End If
End If
End Sub
Private Sub HTip_KeyPress(Index As Integer, KeyAscii As Integer)
If Index < 5 Then KeyAscii = 0
End Sub
Private Sub Quit_Click()
Call MemoryCheck(AccountTip)
End Sub
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
    'Sql = Sql & " and [Charged To] = '" & HTip(0).Text & "'"   ' tighten match with account name

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

