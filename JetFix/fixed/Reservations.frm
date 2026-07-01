VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Reservations 
   BorderStyle     =   0  'None
   Caption         =   "Reservations"
   ClientHeight    =   10530
   ClientLeft      =   90
   ClientTop       =   450
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin BTNENHLib4.BtnEnh BtnEnh8 
      Height          =   615
      Left            =   9120
      TabIndex        =   27
      Top             =   2520
      Width           =   2175
      _Version        =   589828
      _ExtentX        =   3836
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Show Customer Information"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8454016
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":0000
      textLT          =   "Reservations.frx":0092
      textCT          =   "Reservations.frx":00AA
      textRT          =   "Reservations.frx":00C2
      textLM          =   "Reservations.frx":00DA
      textRM          =   "Reservations.frx":00F2
      textLB          =   "Reservations.frx":010A
      textCB          =   "Reservations.frx":0122
      textRB          =   "Reservations.frx":013A
      colorBack       =   "Reservations.frx":0152
      colorIntern     =   "Reservations.frx":017C
      colorMO         =   "Reservations.frx":01A6
      colorFocus      =   "Reservations.frx":01D0
      colorDisabled   =   "Reservations.frx":01FA
      colorPressed    =   "Reservations.frx":0224
      Orientation     =   3
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
      Office2003ColorNorth=   0
      Office2003ColorSouth=   0
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   0
      Office2003ColorSouthMouseOver=   0
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   0
      Office2003ColorSouthPressed=   0
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   7
      Left            =   6120
      TabIndex        =   24
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   6
      Left            =   6120
      TabIndex        =   22
      Top             =   1320
      Width           =   2775
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   855
      Left            =   10920
      TabIndex        =   14
      Top             =   3360
      Width           =   2055
      _Version        =   589828
      _ExtentX        =   3625
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Save "
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":024E
      textLT          =   "Reservations.frx":02B8
      textCT          =   "Reservations.frx":02D0
      textRT          =   "Reservations.frx":02E8
      textLM          =   "Reservations.frx":0300
      textRM          =   "Reservations.frx":0318
      textLB          =   "Reservations.frx":0330
      textCB          =   "Reservations.frx":0348
      textRB          =   "Reservations.frx":0360
      colorBack       =   "Reservations.frx":0378
      colorIntern     =   "Reservations.frx":03A2
      colorMO         =   "Reservations.frx":03CC
      colorFocus      =   "Reservations.frx":03F6
      colorDisabled   =   "Reservations.frx":0420
      colorPressed    =   "Reservations.frx":044A
      Orientation     =   3
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
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Index           =   5
      Left            =   9120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   10200
      TabIndex        =   12
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   6120
      TabIndex        =   11
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   6120
      TabIndex        =   10
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   6120
      TabIndex        =   9
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Data 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   2775
   End
   Begin VSFlex8Ctl.VSFlexGrid Reserve 
      Height          =   6015
      Left            =   600
      TabIndex        =   1
      Top             =   4200
      Width           =   14055
      _cx             =   1973706967
      _cy             =   1973692786
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      BackColorBkg    =   16777215
      BackColorAlternate=   14342874
      GridColor       =   -2147483633
      GridColorFixed  =   13290186
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      DataMember      =   ""
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
   Begin MSComCtl2.MonthView Rdate 
      Height          =   3210
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   5662
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   135856130
      TitleBackColor  =   16749608
      TitleForeColor  =   16777215
      TrailingForeColor=   12632256
      CurrentDate     =   38455
   End
   Begin BTNENHLib4.BtnEnh BtnEnh2 
      Height          =   855
      Left            =   7560
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Cancel Reservation"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":0474
      textLT          =   "Reservations.frx":04F8
      textCT          =   "Reservations.frx":0510
      textRT          =   "Reservations.frx":0528
      textLM          =   "Reservations.frx":0540
      textRM          =   "Reservations.frx":0558
      textLB          =   "Reservations.frx":0570
      textCB          =   "Reservations.frx":0588
      textRB          =   "Reservations.frx":05A0
      colorBack       =   "Reservations.frx":05B8
      colorIntern     =   "Reservations.frx":05E2
      colorMO         =   "Reservations.frx":060C
      colorFocus      =   "Reservations.frx":0636
      colorDisabled   =   "Reservations.frx":0660
      colorPressed    =   "Reservations.frx":068A
      Orientation     =   3
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
      Office2003ColorNorth=   187
      Office2003ColorSouth=   128
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   187
      Office2003ColorSouthMouseOver=   187
      Office2003ColorBorderMouseOver=   1812708
      Office2003ColorNorthPressed=   187
      Office2003ColorSouthPressed=   128
      Office2003ColorBorderPressed=   1812708
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh3 
      Height          =   855
      Left            =   13080
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Exit"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      Picture         =   "C:\Icons\Exit.ico"
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPicture      =   "Reservations.frx":06B4
      textCaption     =   "Reservations.frx":06CC
      textLT          =   "Reservations.frx":0734
      textCT          =   "Reservations.frx":074C
      textRT          =   "Reservations.frx":0764
      textLM          =   "Reservations.frx":077C
      textRM          =   "Reservations.frx":0794
      textLB          =   "Reservations.frx":07AC
      textCB          =   "Reservations.frx":07C4
      textRB          =   "Reservations.frx":07DC
      colorBack       =   "Reservations.frx":07F4
      colorIntern     =   "Reservations.frx":081E
      colorMO         =   "Reservations.frx":0848
      colorFocus      =   "Reservations.frx":0872
      colorDisabled   =   "Reservations.frx":089C
      colorPressed    =   "Reservations.frx":08C6
      Orientation     =   3
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
   Begin BTNENHLib4.BtnEnh BtnEnh4 
      Height          =   855
      Left            =   5880
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Find Reservation"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":08F0
      textLT          =   "Reservations.frx":0970
      textCT          =   "Reservations.frx":0988
      textRT          =   "Reservations.frx":09A0
      textLM          =   "Reservations.frx":09B8
      textRM          =   "Reservations.frx":09D0
      textLB          =   "Reservations.frx":09E8
      textCB          =   "Reservations.frx":0A00
      textRB          =   "Reservations.frx":0A18
      colorBack       =   "Reservations.frx":0A30
      colorIntern     =   "Reservations.frx":0A5A
      colorMO         =   "Reservations.frx":0A84
      colorFocus      =   "Reservations.frx":0AAE
      colorDisabled   =   "Reservations.frx":0AD8
      colorPressed    =   "Reservations.frx":0B02
      Orientation     =   3
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
      Office2003ColorNorth=   0
      Office2003ColorSouth=   0
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   0
      Office2003ColorSouthMouseOver=   0
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   0
      Office2003ColorSouthPressed=   0
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh5 
      Height          =   855
      Left            =   9240
      TabIndex        =   20
      Top             =   3360
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Print Reservation Sheet"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":0B2C
      textLT          =   "Reservations.frx":0BBA
      textCT          =   "Reservations.frx":0BD2
      textRT          =   "Reservations.frx":0BEA
      textLM          =   "Reservations.frx":0C02
      textRM          =   "Reservations.frx":0C1A
      textLB          =   "Reservations.frx":0C32
      textCB          =   "Reservations.frx":0C4A
      textRB          =   "Reservations.frx":0C62
      colorBack       =   "Reservations.frx":0C7A
      colorIntern     =   "Reservations.frx":0CA4
      colorMO         =   "Reservations.frx":0CCE
      colorFocus      =   "Reservations.frx":0CF8
      colorDisabled   =   "Reservations.frx":0D22
      colorPressed    =   "Reservations.frx":0D4C
      Orientation     =   3
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
      Office2003ColorNorth=   15790320
      Office2003ColorSouth=   14211288
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   15790320
      Office2003ColorSouthPressed=   14211288
      Office2003ColorBorderPressed=   12632256
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh6 
      Height          =   615
      Left            =   12000
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
      _Version        =   589828
      _ExtentX        =   3836
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Edit Order"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8454016
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":0D76
      textLT          =   "Reservations.frx":0DEA
      textCT          =   "Reservations.frx":0E02
      textRT          =   "Reservations.frx":0E1A
      textLM          =   "Reservations.frx":0E32
      textRM          =   "Reservations.frx":0E4A
      textLB          =   "Reservations.frx":0E62
      textCB          =   "Reservations.frx":0E7A
      textRB          =   "Reservations.frx":0E92
      colorBack       =   "Reservations.frx":0EAA
      colorIntern     =   "Reservations.frx":0ED4
      colorMO         =   "Reservations.frx":0EFE
      colorFocus      =   "Reservations.frx":0F28
      colorDisabled   =   "Reservations.frx":0F52
      colorPressed    =   "Reservations.frx":0F7C
      Orientation     =   3
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
      Office2003ColorNorth=   0
      Office2003ColorSouth=   0
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   0
      Office2003ColorSouthMouseOver=   0
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   0
      Office2003ColorSouthPressed=   0
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh7 
      Height          =   855
      Left            =   12600
      TabIndex        =   26
      Top             =   240
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Reserve Section"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Surface         =   15
      BackColorContainer=   8454016
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Reservations.frx":0FA6
      textLT          =   "Reservations.frx":1024
      textCT          =   "Reservations.frx":103C
      textRT          =   "Reservations.frx":1054
      textLM          =   "Reservations.frx":106C
      textRM          =   "Reservations.frx":1084
      textLB          =   "Reservations.frx":109C
      textCB          =   "Reservations.frx":10B4
      textRB          =   "Reservations.frx":10CC
      colorBack       =   "Reservations.frx":10E4
      colorIntern     =   "Reservations.frx":110E
      colorMO         =   "Reservations.frx":1138
      colorFocus      =   "Reservations.frx":1162
      colorDisabled   =   "Reservations.frx":118C
      colorPressed    =   "Reservations.frx":11B6
      Orientation     =   3
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
      Office2003ColorNorth=   0
      Office2003ColorSouth=   0
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   0
      Office2003ColorSouthMouseOver=   0
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   0
      Office2003ColorSouthPressed=   0
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   25
      Top             =   2280
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   21
      Top             =   1320
      Width           =   1000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Touch Box To Add Guest Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label LDate 
      BackStyle       =   0  'Transparent
      Caption         =   "MyDate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   6
      Top             =   480
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Party #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   2760
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      Top             =   840
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   360
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   1800
      Width           =   1000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   4920
      Top             =   120
      Width           =   9495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   10755
      Left            =   120
      Top             =   45
      Width           =   15015
   End
End
Attribute VB_Name = "Reservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TotalRows As Integer
Private HoldSection As Boolean
Private Sub BtnEnh1_Click()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Tb As ADODB.Recordset
Dim P  As Integer
On Error GoTo ErrorHandler
If Not IsDate(Data(0).Text) Then
  See_Msg "Invalid Date", 5, ""
  Exit Sub
End If
If Not IsDate(Data(1).Text) Then
  See_Msg "Invalid Time", 5, ""
  Exit Sub
End If
If Data(6).Text = "" Then
  See_Msg "Phone Number Required", 5, ""
  Exit Sub
End If
If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Sub
If ConnectADODB(SvrPath.TerminalPath & TableLayoutDb, Db) = False Then Exit Sub
 Sql = "Select * from [Reservations] Where [Name] = '" & Data(2).Text & "'"
 Sql = Sql & " and [Phone] = '" & Data(6).Text & "'"
 Sql = Sql & " and TRY_CAST([Date] AS datetime) = '" & CDate(Data(0).Text) & "'"
If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
   If RS.EOF Then RS.AddNew
   RS![Date] = Format(Data(0).Text, "M/D/YYYY")
   RS![Time] = Format(Data(1).Text, "H:MM ampm")
   RS![Name] = Data(2).Text
   RS![Number in Party] = Val(Data(3).Text)
   RS![Seating] = Data(4).Text & ""
   RS![Memo] = Data(5).Text & " "
   RS![Phone] = Data(6).Text & ""
   RS![Email] = Data(7).Text & ""
 RS.Update
DbClose Db, RS
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
Sql = "Select * from [General Data] where  [Phone] = '" & Data(6).Text & "'"
Sql = Sql & " and [Scan Name] = '" & Data(2).Text & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
   If Tb.EOF Then
    Tb.AddNew
    Tb.fields("Scan Name").Value = Data(2).Text
    Tb(1) = 0
    Tb(2) = 0
    Tb(3) = Data(6).Text & ""
    Tb(4) = 0
    Tb(5) = 0
    Tb(6) = 0
    Tb![Credit Balance] = 1
    Tb![Address] = ""
    Tb![City] = ""
    Tb![State] = ""
    Tb![Zip] = ""
    Tb![Phone] = Data(6).Text & ""
    Tb![Account Type] = GetConfigInfo("Functions", "Reservations Report To", "Frequent Diner", "FUNCTIONS.INI") & ""
    Tb![Cell] = Data(6).Text & ""
    Tb![Email] = Data(7).Text & ""
    Tb![Fax] = ""
    Tb![Birthday] = ""
    Tb![Anniversary] = ""
    Tb![Account Expires On] = ""
    Tb![Spc Interest 1] = ""
    Tb![Spc Interest 2] = ""
    Tb![Spc Interest 3] = ""
    Tb![Spc Interest 4] = ""
    Tb![Memo] = Data(5).Text & " "
    Tb![Meal Limit] = 0
    Tb![Discount] = "None"
    Tb![Card Number] = "Not Used"
    Tb![Card Expires] = "Not Used"
    Tb![Customer Id] = 0
    Tb.Update
  End If
DbClose Db, RS
If Check_Option(209) = 1 Then
 depositpaid = True
 PaidIn.Show 1
 depositpaid = False
 If DepositAmtPaid > 0 Then
  EditGiftCards Data(6).Text & "", DepositAmtPaid, Data(5).Text & " ", "Reservation Deposit", "Never", Data(2).Text
  PrintGiftChit Data(6).Text & "", DepositAmtPaid, "Reservation Deposit"
  DepositAmtPaid = 0
 End If
End If
If Data(7).Text > "" Then
   If See_Msg("EMail Confirmation?", 4, "Send Email") = 6 Then
    Dim MsgTo As String
    MsgTo = "We have you confirmed for " & vbCrLf
    MsgTo = MsgTo & Data(0).Text & " @ " & Data(1).Text & vbCrLf
    MsgTo = MsgTo & "See You Then." & vbCrLf
    MsgTo = MsgTo & "Thank You" & vbCrLf
    MsgTo = MsgTo & StoreName & vbCrLf
    ReportView.SendEmailTo Data(7).Text, Data(2).Text, MsgTo
   End If
End If

If Check_Option(208) = 1 Then
  If See_Msg("Do you Want to Place an Order ?", 4, "") = 6 Then
    Table_Number = Data(2).Text
    CallorderInfo = "Reservation for " & Data(0).Text & " @ " & Data(1).Text
    ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
        Sql = "Select * from [Customer Data] where [Customer Name] = '" & Data(2).Text & "'"
        Sql = Sql & " AND [Customer Phone] = '" & Data(6).Text & "'"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    If Tb.EOF Then
       Tb.AddNew
    End If
        Tb![Customer Name] = Data(2).Text & ""
        Tb![Customer Phone] = Data(6).Text & ""
        Tb![Customer Address] = ""
        Tb![Customer City] = ""
        Tb![Customer State] = ""
        Tb![Customer Zip] = ""
        Tb![Instructions] = Data(5).Text & "" 'max field len 255
        Tb![Customer Discount] = ""
        Tb![Email] = ""
    Tb.Update
    DbClose Db, Tb
    Set Db = Nothing
    Check_no = CallInNumber 'Advance_Check(SvrPath.TerminalPath, CallInOrderDb, False)
    Set AllOrders = Nothing
    'Set EachOrder = Nothing
    Set AllOrders = New AllCallins
        AllOrders.add Data(6).Text
        With AllOrders.Item(Data(6).Text)
            .CallerName = Trim$(Data(2).Text)
            .CallerPhone = Data(6).Text
            .callerAddress = ""
            .City = ""
            .State = ""
            .Zip = ""
            .PTime = IsPickUpDate(Data(1).Text, PTime)
            .PDate = IsPickUpDate(Data(0).Text, PDate)
            .Deposits = 0   'Val(CallerInfo(8).Text)
            .PickUpNo = Check_no
            .Instruction = Data(5).Text & ""
            .OrderType = "Reservation"
            .Email = "" 'CallerInfo(10).Text
            .AccountOnFile = ""
            .ExpiresOnFile = ""
        End With
    If AllOrders Is Nothing Then Exit Sub
    ReDim NextRow(0)
    Number_In_Party = 1
    DT_Order_Status = MyOrderStatus(ParkCheck)
    Check_Status = MyOrderStatus(LoadCallInCheck)
    ReDim_Array
    Sales_Totals
    Order.LoadOrderScreen "Reservation BtnEnh1"
    Order.Amount_Tend_Update
    DT_Order_Status = MyOrderStatus(ParkCheck)
    Check_Status = MyOrderStatus(LoadCallInCheck)
    Order.CheckNo = Check_no
    OrderIs = "Reservation"
    DepositAmtPaid = 0
    Order.Show 1
    If DepositAmtPaid <= 0 And Check_Option(209) = 1 Then
        If ConnectADODB(SvrPath.TerminalPath & TableLayoutDb, Db) = False Then Exit Sub
        Sql = "Delete from [Reservations] Where [Name] = '" & Data(2).Text & "'"
        Sql = Sql & " and [Phone] = '" & Data(6).Text & "'"
        Sql = Sql & " and TRY_CAST([Date] AS datetime) = '" & CDate(Data(0).Text) & "'"
        Sql = Sql & " and TRY_CAST([Time] AS datetime) = '" & Data(1).Text & "'"
        Db.Execute Sql
        See_Msg "Cannot Save Reservation", 5, ""
    End If
  End If
End If

Sql = "Select * from [Reservations] Where TRY_CAST([date] AS datetime) = '" & CDate(Data(0).Text) & "'"
Sql = Sql & " Order by [Time],[Name] asc"
LoadGrid Sql, False
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "RESERVATIONS BTN1 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh2_Click()
If See_Msg("Delete Reservation ?", 4, "Cancel") <> 6 Then Exit Sub
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim P  As Integer
If Not IsDate(Data(0).Text) Then
  See_Msg "Invalid Date", 5, ""
  Exit Sub
End If
If Not IsDate(Data(1).Text) Then
  See_Msg "Invalid Time", 5, ""
  Exit Sub
End If
HoldSection = False
Sql = "Delete from [Reservations] Where [Name] = '" & Data(2).Text & "'"
Sql = Sql & " and [Phone] = '" & Data(6).Text & "'"
Sql = Sql & " and TRY_CAST([Date] AS datetime) = '" & CDate(Data(0).Text) & "'"
Sql = Sql & " and [Time] = '" & Data(1).Text & "'"
If ConnectADODB(SvrPath.TerminalPath & TableLayoutDb, Db) = False Then Exit Sub
Db.Execute Sql
  Sql = "Select * from [Reservations] Where TRY_CAST([date] AS datetime) = '" & CDate(Data(0).Text) & "'"
  Sql = Sql & " Order by [Time],[Name] asc"
Data(0).Text = ""
Data(1).Text = ""
Data(3).Text = ""
Data(4).Text = ""
Data(5).Text = ""
Data(6).Text = ""
LoadGrid Sql, False
End Sub
Private Sub BtnEnh3_Click()
Call MemoryCheck(Me)
End Sub
Private Sub BtnEnh4_Click()
Dim SearchStr As String
SearchStr = ShowKeyBoard("Name or Phone")
If SearchStr = "" Then Exit Sub
Sql = "Select * from [Reservations] "
Sql = Sql & " Where [Name] Like '" & SearchStr & "%'"
Sql = Sql & " Or [Phone] like '" & SearchStr & "%'"
Sql = Sql & " Order by TRY_CAST([DAte] AS datetime), TRY_CAST([Time] AS datetime),[Name] asc"
LoadGrid Sql, True
End Sub

Private Sub BtnEnh5_Click()
PrintReport
End Sub
Private Sub BtnEnh6_Click()
On Error GoTo ErrorHandler
Dim R       As Integer
Dim Sql     As String
Dim ErrMsg  As String
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim RS      As ADODB.Recordset
If Trim(Data(2).Text) = "" Then
    R = See_Msg("Please Enter Name", 5, "")
    Exit Sub
End If
If Trim(Data(6).Text) = "" Then
    R = See_Msg("Please Enter Phone", 5, "")
    Exit Sub
End If
Load_Reservation_Order Data(6).Text, Data(2).Text, Data(0).Text, Data(1).Text
If ConnectADODB(SvrPath.TerminalPath & TableLayoutDb, Db) = False Then Exit Sub
    Sql = "Select * from [Reservations] Where [Name] = '" & Data(2).Text & "'"
    Sql = Sql & " and [Phone] = '" & Data(6).Text & "'"
    Sql = Sql & " and TRY_CAST([Date] AS datetime) = '" & CDate(Data(0).Text) & "'"
    Sql = Sql & " and TRY_CAST([Time] AS datetime) = '" & Data(1).Text & "'"
If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
 If RS.EOF Then RS.AddNew
   RS![Date] = Format(Data(0).Text, "M/D/YYYY")
   RS![Time] = Format(Data(1).Text, "H:MM ampm")
   RS![Name] = Data(2).Text
   RS![Number in Party] = Val(Data(3).Text)
   RS![Seating] = Data(4).Text & ""
   RS![Memo] = Data(5).Text & ""
   RS![Phone] = Data(6).Text & ""
 RS.Update
DbClose Db, RS
Sql = "Select * from [Reservations] Where TRY_CAST([date] AS datetime) = '" & CDate(Data(0).Text) & "'"
Sql = Sql & " Order by [Time],[Name] asc"
LoadGrid Sql, False
BtnEnh1.Visible = True
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reserveration Place Order ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh7_Click()
On Error GoTo ErrorHandler
If Not IsDate(Data(1).Text) Then
   See_Msg "Invalid Time", 5, ""
   Exit Sub
End If
If Not IsDate(Data(0).Text) Then
   See_Msg "Invalid Date", 5, ""
   Exit Sub
End If
If Data(2).Text = "" Then
   See_Msg "Name Required", 5, ""
   Exit Sub
End If
ReserveDate = Data(0).Text
ReserveTime = Data(1).Text
Dim Db As ADODB.Connection
Dim Ts As ADODB.Recordset
Dim RS As ADODB.Recordset
HoldSection = True
Load_What = "Change Sections"
Find_What = "Choose a Section"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
If See_Msg("Reserve Section?", 4, "") <> 6 Then Exit Sub
Data(5).Text = Data(5).Text & "Area Reserved " & GetSelected
If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Sub
If ConnectADODB(SvrPath.TerminalPath & TableLayoutDb, Db) = False Then Exit Sub
 Sql = "Select * from [Layout] Where [Layout Name] = '" & GetSelected & "'"
If ConnectADORS(Sql, Db, Ts) = False Then Exit Sub
 Do While Not Ts.EOF
 If ReserveTable(Ts.fields("Table Name").Value) = True Then
     See_Msg "Section not available at this time", 5, ""
     DbClose Db, Ts
     Exit Sub
 End If
 Ts.MoveNext
Loop
If Not Ts.BOF Then Ts.MoveFirst
Do While Not Ts.EOF
 Sql = "Select * from [Reservations] Where [Name] = '" & Data(2).Text & "'"
 Sql = Sql & " and [Phone] = '" & Data(6).Text & "'"
 Sql = Sql & " and [Seating] = '" & Ts.fields("Table Name").Value & "" & "'"
 Sql = Sql & " and TRY_CAST([Date] AS datetime) = '" & CDate(Data(0).Text) & "'"
    If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
       If RS.EOF Then RS.AddNew
       RS![Date] = Format(Data(0).Text, "M/D/YYYY")
       RS![Time] = Format(Data(1).Text, "H:MM ampm")
       RS![Name] = Data(2).Text
       RS![Number in Party] = Val(Data(3).Text)
       RS![Seating] = Ts.fields("Table Name").Value & ""
       RS![Memo] = Data(5).Text & " "
       RS![Phone] = Data(6).Text & ""
       RS![Email] = Data(7).Text & ""
     RS.Update
     SetNothing RS
   Ts.MoveNext
 Loop
DbClose Db, Ts
If Check_Option(208) = 1 Then
  If See_Msg("Do you Want to Place an Order ?", 4, "") = 6 Then
    Table_Number = Data(2).Text
    CallorderInfo = "Reservation for " & Data(0).Text & " @ " & Data(1).Text
    ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
        Sql = "Select * from [Customer Data] where [Customer Name] = '" & Data(2).Text & "'"
        Sql = Sql & " AND [Customer Phone] = '" & Data(6).Text & "'"
    ConnectADORS Sql, Db, Ts
    If Ts.EOF Then
       Ts.AddNew
    End If
        Ts![Customer Name] = Data(2).Text & ""
        Ts![Customer Phone] = Data(6).Text & ""
        Ts![Customer Address] = ""
        Ts![Customer City] = ""
        Ts![Customer State] = ""
        Ts![Customer Zip] = ""
        Ts![Instructions] = Data(5).Text & "" 'max field len 255
        Ts![Customer Discount] = ""
        Ts![Email] = ""
    Ts.Update
    SetNothing Db
    Check_no = CallInNumber 'Advance_Check(SvrPath.TerminalPath, CallInOrderDb, False)
    Set AllOrders = Nothing
    'Set EachOrder = Nothing
    Set AllOrders = New AllCallins
        AllOrders.add Data(6).Text
        With AllOrders.Item(Data(6).Text)
            .CallerName = Trim$(Data(2).Text)
            .CallerPhone = Data(6).Text
            .callerAddress = ""
            .City = ""
            .State = ""
            .Zip = ""
            .PTime = IsPickUpDate(Data(1).Text, PTime)
            .PDate = IsPickUpDate(Data(0).Text, PDate)
            .Deposits = 0   'Val(CallerInfo(8).Text)
            .PickUpNo = Check_no
            .Instruction = Data(5).Text & ""
            .OrderType = "Reservation"
            .Email = "" 'CallerInfo(10).Text
            .AccountOnFile = ""
            .ExpiresOnFile = ""
        End With
    If AllOrders Is Nothing Then Exit Sub
    ReDim NextRow(0)
    Number_In_Party = 1
    DT_Order_Status = MyOrderStatus(ParkCheck)
    Check_Status = MyOrderStatus(LoadCallInCheck)
    ReDim_Array
    Sales_Totals
    Order.LoadOrderScreen "Reservation BtnEnh7"
    Order.Amount_Tend_Update
    DT_Order_Status = MyOrderStatus(ParkCheck)
    Check_Status = MyOrderStatus(LoadCallInCheck)
    Order.CheckNo = Check_no
    OrderIs = "Reservation"
    DepositAmtPaid = 0
    Order.Show 1
   
  End If
End If
BtnEnh1_Click
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "RESERVATIONS BTN1 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub BtnEnh8_Click()
      Assign_Name.UserPrvilage True
      Assign_Name.Scan_Data(0).Text = Data(2).Text
      Assign_Name.GetAccount
      If Assign_Name.Scan_Data(3).Text = "" Then
        Unload Assign_Name
        Exit Sub
      End If
      Assign_Name.Scan_Data(0).DataChanged = False
      Assign_Name.Show 1

End Sub

Private Sub Data_Click(Index As Integer)
Dim DataStr    As String
Dim TempAccess As Long
Dim TempName   As String
Select Case Index
   Case Is = 1
      'If Check_Option(19) = 1 Then Exit Sub
      Dim PKT As DateOrTime
      Set PKT = New DateOrTime
        DataStr = PKT.ReturnTime
         If DataStr > "" Then
           Data(Index).Text = Format(DataStr, "H:MM ampm")
         End If
      Set PKT = Nothing
   Case Is = 2, 5, 6, 7
    'If Check_Option(19) = 1 Then Exit Sub
        Dim MyText As String
        Dim GK As MSaleKeyboard.MyKeyboard
        Set GK = New MSaleKeyboard.MyKeyboard
            DataStr = GK.GetName()
        Set GK = Nothing
        'DataStr = ShowKeyBoard("Enter Information")
        
        If DataStr > "" Then
          If Index = 6 Then
             Data(6).Text = FormatPhoneNumber(DataStr)
             
             If Data(6).Text = "" Then Exit Sub
          ElseIf Index = 2 Then
            If See_Msg(GetConfigInfo("Reservations", "Name Format", "First and Last", "FUNCTIONS.INI"), 4, DataStr & " Is This Correct") = 6 Then
              Data(Index).Text = DataStr
            Else
              Exit Sub
            End If
          Else
             Data(Index).Text = DataStr
          End If
        End If
        If DataStr > "" And Index = 6 Then FindGoodCustomer Data(6).Text
   Case Is = 3
    'If Check_Option(19) = 1 Then Exit Sub
    DataStr = CStr(GetCovers)
    If Val(DataStr) > 0 Then
      Data(Index).Text = DataStr
    End If
   Case Is = 4
    If Check_Option(171) = 0 And Check_Option(89) = 1 Then
     ReserveDate = Data(0).Text
     ReserveTime = Data(1).Text
     If Not IsDate(ReserveTime) Then
       See_Msg "Invalid Time", 5, ""
       Exit Sub
     End If
     If Not IsDate(ReserveDate) Then
       See_Msg "Invalid Date", 5, ""
       Exit Sub
     End If
     TempAccess = LoggedOnAs(0).EmpAccessLevel
     TempName = LoggedOnAs(0).EMPName
     LoggedOnAs(0).EmpAccessLevel = 5
     LoggedOnAs(0).EMPName = "Reservations"
     OpenTableOnly = True
     TableLayout.LoadTableScreen
     TableLayout.Show 1
     LoggedOnAs(0).EmpAccessLevel = TempAccess
     LoggedOnAs(0).EMPName = TempName ''
     Data(4).Text = Table_Number
     Call MemoryCheck(TableLayout)
     Set TableLayout = Nothing
   Else
      Load_What = "Preferences"
      Find_What = "Select Preference"
      Select_What.Show 1
     If GetSelected = "" Then Exit Sub
     If GetSelected = "Keyboard" Then
      DataStr = ShowKeyBoard("Type Preference")
     Else
      DataStr = GetSelected
     End If
     If DataStr > "" Then
      Data(Index).Text = DataStr
     End If
   End If
End Select
End Sub
Public Function FindReservationOrder(PHONENO As String, PickUpTime As String, PickUpDate As String) As Boolean
On Error GoTo ermsg
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
   Sql = "Select * from [Open Checks] Where [Customer Phone] = '" & PHONENO & "'"
   If CallInOrderDb = "Call In Orders.mdb" Then
     Sql = Sql & " and TRY_CAST([Pick Up time] AS datetime) = '" & Format(CDate(PickUpTime), "yyyy-mm-dd hh:nn:ss") & "' and TRY_CAST([Pick Up Date] AS datetime) = '" & Format(CDate(PickUpDate), "yyyy-mm-dd hh:nn:ss") & "'"
   Else
     Sql = Sql & " and TRY_CAST([Pick Up time] AS datetime) = '" & PickUpTime & "' and TRY_CAST([Pick Up Date] AS date) = '" & PickUpDate & "'"
   End If
   Sql = Sql & " and [Order Destination] = 'Reservation'"
ConnectADORS Sql, Db, Tb
If Not Tb.EOF Then
   FindReservationOrder = True
Else
  FindReservationOrder = False
End If
DbClose Db, Tb
Exit Function
ermsg:
Error_Message "Find Reservation", Err.Description
Resume Next
End Function
Public Sub Load_Reservation_Order(RPhone As String, RName As String, Rdate As String, rTime As String)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim RP1 As String
Dim RP2 As String
Dim RP3 As String
Dim RP4 As String
Dim RP5 As String
Dim RP6 As String
Dim IsPresented As Integer
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
   Sql = "Select * from [Open Checks] Where [Customer Phone] = '" & RPhone & "'"
   Sql = Sql & " and [Customer Name] = '" & RName & "'"
   If CallInOrderDb = "Call In Orders.mdb" Then
     Sql = Sql & " and TRY_CAST([Pick Up time] AS datetime) = '" & Format(CDate(rTime), "yyyy-mm-dd hh:nn:ss") & "' and TRY_CAST([Pick Up Date] AS datetime) = '" & Format(CDate(Rdate), "yyyy-mm-dd hh:nn:ss") & "'"
   Else
     Sql = Sql & " and TRY_CAST([Pick Up time] AS datetime) = '" & rTime & "' and TRY_CAST([Pick Up Date] AS date) = '" & Rdate & "'"
   End If
   Sql = Sql & " and [Order Destination] = 'Reservation'"
ConnectADORS Sql, Db, Tb
        If Tb.EOF Then
           DbClose Db, Tb
           See_Msg "Order Not Found", 5, ""
           Exit Sub
         Else
                Order.LoadOrderScreen "Reservation Load Order"
                Table_Number = Tb(1)
                Check_no = Val(Check_Null(Tb(2)))
                Number_In_Party = Tb(3)
                IsPresented = Val(Check_Null(Tb(5)))
                OrderIs = Tb(18)
                Set AllOrders = Nothing
                'Set EachOrder = Nothing
                Set AllOrders = New AllCallins
                If Trim(Tb![Customer Phone]) = "" Then
                  Tb![Customer Phone] = "None"  'add if phone number is empty will allow call in to close
                  Tb.Update
                End If
                AllOrders.add Tb![Customer Phone]
                   With AllOrders.Item(Tb![Customer Phone])
                    .Deposits = Tb![Deposit]
                    .CallerName = Trim$(Tb![Customer Name])  'was set to table number
                    .PDate = IsPickUpDate(Tb(24) & "", PDate)
                    .PTime = IsPickUpDate(Tb(25) & "", PTime)
                    .PickUpNo = Check_no
                    .CallerPhone = Tb![Customer Phone] & ""
                    .callerAddress = Tb![Street] & ""
                    .DriverName = Tb![Server Name] & ""
                    .AccountOnFile = ""
                    .ExpiresOnFile = ""
                 End With
            End If
       SetNothing Tb
Top:
   Clear_Grid Order.Ticket(0)
   ReDim_Array
   Selection_Number = Abs(GetLastSelectionNumber(Check_no, Db))
   Sql = "Select * from [Check detail] "
   Sql = Sql & " where [Check Number] = " & Check_no
   Sql = Sql & " and [customer Name] = '" & RName & "'"
   Sql = Sql & " Order By Abs([Selection Number]) , [Pos] asc "
   ConnectADORS Sql, Db, Tb
      ReDim NextRow(Number_In_Party)
       If Tb.EOF Then
          SetNothing Tb
          Sql = "Select * from [Open Checks] where [Close Out Day] = 0 and [Check Number] = " & Check_no
          Sql = Sql & " and [Customer Name] = '" & RName & "'"
          ConnectADORS Sql, Db, Tb
          If Not Tb.EOF Then
             Tb.delete
             DbClose Db, Tb
          End If
          Exit Sub
       End If
       Dim Cs As Integer, ShowItem As Integer
       Do While Not Tb.EOF
            Cs = Tb(10)
            ShowItem = 0   'Tb(8)
                  If ShowItem = 0 Then
                     If Cs <> Tb(10) Then Cs = Tb(10)
                        If Cs > Number_In_Party Then
                          Tb(10) = Number_In_Party
                          Tb.Update
                        End If
                        If Tb(2) = 0 Then
                           Grid_array(Tb(10), NextRow(Tb(10)), 0) = ""
                        Else
                           Grid_array(Tb(10), NextRow(Tb(10)), 0) = Tb(2)
                        End If
                        If Val(Tb(8)) > 0 Then
                           Grid_array(Tb(10), NextRow(Tb(10)), 1) = Chr(118) & "oid-> " & Tb(3)
                        ElseIf IsPresented = 0 Then
                           Grid_array(Tb(10), NextRow(Tb(10)), 1) = Tb(3)
                        Else
                           Grid_array(Tb(10), NextRow(Tb(10)), 1) = "* " & Tb(3)
                        End If
                       If Tb(4) = 0 Then
                           Grid_array(Tb(10), NextRow(Tb(10)), 2) = ""
                        Else
                          ' If Check_Option(43) = 0 Then
                          '   Grid_array(Tb(10), NextRow(Tb(10)), 2) = Align(7, Tb(4) + Tb(7))
                          ' Else
                             Grid_array(Tb(10), NextRow(Tb(10)), 2) = Align(7, Tb(4))
                          ' End If
                        End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 3) = GetExpandPrint(Tb(3))
                           Grid_array(Tb(10), NextRow(Tb(10)), 10) = Tb(5)
                           Grid_array(Tb(10), NextRow(Tb(10)), 11) = Tb(6)
                           Grid_array(Tb(10), NextRow(Tb(10)), 12) = Tb(7)
                        If Tb(8) = 1 Then
                           Grid_array(Tb(10), NextRow(Tb(10)), 13) = -1 * Tb(8)
                        Else
                           Grid_array(Tb(10), NextRow(Tb(10)), 13) = Tb(8)
                        End If
                        If IsPresented = 0 Then
                           Grid_array(Tb(10), NextRow(Tb(10)), 14) = Abs(Tb(13))  'hold selection number for voided items
                        Else
                           Grid_array(Tb(10), NextRow(Tb(10)), 14) = Tb(13)  'hold selection number for voided items
                        End If
                          Grid_array(Tb(10), NextRow(Tb(10)), 3) = Tb![Expand Print]
                          Grid_array(Tb(10), NextRow(Tb(10)), 4) = Tb![Remote Print]
                          Grid_array(Tb(10), NextRow(Tb(10)), 5) = Tb![Remote 1]
                          Grid_array(Tb(10), NextRow(Tb(10)), 6) = Tb![Remote 2]
                          Grid_array(Tb(10), NextRow(Tb(10)), 7) = Tb![Remote 3]
                          Grid_array(Tb(10), NextRow(Tb(10)), 8) = Tb![Remote 4]
                          Grid_array(Tb(10), NextRow(Tb(10)), 9) = Tb![Remote 5]
                          Grid_array(Tb(10), NextRow(Tb(10)), 30) = Tb(3) & ""
                          Grid_array(Tb(10), NextRow(Tb(10)), 31) = Tb![Remote 6]
                          
                        If Check_Option(189) = 1 Then
                           If Not AllMenuItems(CStr(Tb(3))) Is Nothing Then
                             RP1 = AllMenuItems(CStr(Tb(3))).Remote1
                           End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 5) = RP1
                           If Not AllMenuItems(CStr(Tb(3))) Is Nothing Then
                             RP2 = AllMenuItems(CStr(Tb(3))).Remote2
                           End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 6) = RP2
                           
                           If Not AllMenuItems(CStr(Tb(3))) Is Nothing Then
                             RP3 = AllMenuItems(CStr(Tb(3))).Remote3
                           End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 7) = RP3
                           
                           If Not AllMenuItems(CStr(Tb(3))) Is Nothing Then
                             RP4 = AllMenuItems(CStr(Tb(3))).Remote4
                           End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 8) = RP4
                           
                           If Not AllMenuItems(CStr(Tb(3))) Is Nothing Then
                             RP5 = AllMenuItems(CStr(Tb(3))).Remote5
                           End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 9) = RP5
                           
                           If Not AllMenuItems(CStr(Tb(3))) Is Nothing Then
                             RP6 = AllMenuItems(CStr(Tb(3))).Remote6
                           End If
                           Grid_array(Tb(10), NextRow(Tb(10)), 31) = RP6
                        End If
                          Grid_array(Tb(10), NextRow(Tb(10)), 15) = Tb(11)
                          Grid_array(Tb(10), NextRow(Tb(10)), 16) = Tb(12)
                          Grid_array(Tb(10), NextRow(Tb(10)), 17) = Tb(14)
                          Grid_array(Tb(10), NextRow(Tb(10)), 18) = Tb(15)
                          Grid_array(Tb(10), NextRow(Tb(10)), 19) = Val(Tb![Record Type] & "")
                          Grid_array(Tb(10), NextRow(Tb(10)), 23) = Tb(26)
                          Grid_array(Tb(10), NextRow(Tb(10)), 24) = Tb(10)
                          Grid_array(Tb(10), NextRow(Tb(10)), 25) = Tb![Menu Number]
                          Grid_array(Tb(10), NextRow(Tb(10)), 26) = Tb![Department Report]
                          Grid_array(Tb(10), NextRow(Tb(10)), 32) = Val(Check_Null(Tb![Inclusive Tax]))
                          NextRow(Tb(10)) = NextRow(Tb(10)) + 1
                    End If
                Tb.MoveNext
           Loop
DbClose Db, Tb
Order.TableNumber.Caption = "For " & Table_Number
Order.Emp_Name.Caption = LoggedOnAs(0).EMPName
'Changes Deposits Paid to Open Check Number for Phone Order
Clear_Grid Order.Ticket(0)
Order.Load_Grid
Order.Amount_Tend_Update
DT_Order_Status = MyOrderStatus(ParkCheck)
Check_Status = MyOrderStatus(CallIncheck)
OrderIs = "Reservation"
Order.Show 1
'Order.Show 1
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Call In Order ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub Form_Load()
ResizeScreen1024 Me
Me.Top = 0:
Me.Left = 0
Me.WindowState = vbMaximized
Rdate.Value = Date
LoggedOnAs(0).EMPName = "Reservation"
LoggedOnAs(0).EmpAccessLevel = 1
Data(0).Text = Format(Date, "MM/DD/YYYY")
LDate.Caption = "Reservations For " & Format(Date, "MM/DD/YYYY")
With Reserve
   .cols = 8
   .rows = 25
   .ColWidth(0) = 1500 * WidScale1024
   .ColWidth(1) = 1300 * WidScale1024
   .ColWidth(2) = 3600 * WidScale1024
   .ColWidth(3) = 1300 * WidScale1024
   .ColWidth(4) = 4500 * WidScale1024
   .ColWidth(5) = 0
   .ColWidth(6) = 2000
   .ColWidth(7) = 0
   .FixedRows = 1
   .FixedCols = 0
   .TextMatrix(0, 0) = "Time"
   .TextMatrix(0, 1) = "Guests"
   .TextMatrix(0, 2) = "Name"
   .TextMatrix(0, 3) = "Seating"
   .TextMatrix(0, 4) = "Note"
   .TextMatrix(0, 5) = "Reservation Date"
   .TextMatrix(0, 6) = "Phone"
   .TextMatrix(0, 7) = ""
End With
    Sql = "Select * from [Reservations] Where TRY_CAST([date] AS datetime) = '" & CDate(Date) & "'"
    Sql = Sql & " Order by [Time],[Name] asc"
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
LoadGrid Sql, False
BtnEnh1.Caption = "Save"
End Sub
Private Sub PrintReport()
Reserve.rows = TotalRows + 1
ReportView.View.Orientation = orLandscape
SetPage ReportView.View, LDate.Caption
ReportView.View.PaperSize = pprLetter
ReportView.View.RenderControl = Reserve.hwnd
ReportView.View.EndDoc
ReportView.Caption = "Reservation List"
ReportView.Show 1
Call MemoryCheck(ReportView)
End Sub
Private Sub FindGoodCustomer(PHONENO As String)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
    Sql = "Select * from [General Data] where  [Phone] = '" & PHONENO & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
   Data(2).Text = Tb(0) & ""
   Data(5).Text = "**" & vbCrLf
   Data(7).Text = Tb.fields("EMail").Value & ""
   If Tb![Memo] & "" > "" Then
     Data(5).Text = Data(5).Text & Tb![Memo] & "" & vbCrLf
   End If
   If Tb![Birthday] & "" > "" Then
      Data(5).Text = Data(5).Text & "Birthday " & Format(Tb![Birthday] & "", "MM/DD") & vbCrLf
   End If
   If Tb![Anniversary] & "" > "" Then
      Data(5).Text = Data(5).Text & "Anniversary " & Format(Tb![Anniversary] & "", "MM/DD") & vbCrLf
   End If
Else
  DbClose Db, Tb
  If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then Exit Sub
    Sql = "Select * from [Customer Data] where  [Customer Phone] = '" & PHONENO & "'"
  If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
  If Not Tb.EOF Then
    If Tb.fields("Instructions").Value & "" > "" Then
      Data(5).Text = "**" & vbCrLf
      Data(5).Text = Data(5).Text & Tb.fields("Instructions").Value & ""
    End If
    Data(7).Text = Tb.fields("EMail").Value & ""
    Data(2).Text = Tb.fields("Customer Name").Value & ""
  End If
End If
DbClose Db, Tb
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "FIND GOOD CUST ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub ClearData()
Dim C As Integer
For C = 1 To 7
   Data(C).Text = ""
Next C
End Sub
Private Sub LoadGrid(SortHow As String, ShowAll As Boolean)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim P  As Integer
On Error GoTo ErrorHandler
BtnEnh6.Visible = False
BtnEnh1.Visible = True

If ShowAll Then
   
   With Reserve
      .ColWidth(5) = 3100 * WidScale800x600
      .ColWidth(4) = 0
   End With
Else
   With Reserve
      .ColWidth(4) = 3100 * WidScale800x600
      .ColWidth(5) = 0
   End With
End If
Reserve.Refresh
If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Sub
If ConnectADODB(SvrPath.TerminalPath & TableLayoutDb, Db) = False Then Exit Sub
If ConnectADORS(SortHow, Db, RS) = False Then Exit Sub
P = 0
Reserve.Clear flexClearScrollable
Do While Not RS.EOF
 P = P + 1
  With Reserve
    .rows = .rows + P
    .TextMatrix(P, 0) = Format(RS![Time] & "", "H:MM ampm")
    .TextMatrix(P, 1) = Format(Val(RS![Number in Party] & ""), "#")
    .TextMatrix(P, 2) = RS![Name] & ""
    .TextMatrix(P, 3) = RS![Seating] & ""
    .TextMatrix(P, 4) = RS![Memo] & ""
    .TextMatrix(P, 5) = Format(RS![Date] & "", "M/D/YYYY")
    .TextMatrix(P, 6) = RS![Phone] & ""
    .TextMatrix(P, 7) = RS![Email] & ""
  End With
 RS.MoveNext
Loop
DbClose Db, RS
TotalRows = P
ClearData
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "LOAD GRID RESERVATIONS ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub Rdate_DateClick(ByVal DateClicked As Date)
Data(0).Text = Format(DateClicked, "MM/DD/YYYY")
LDate.Caption = "Reservations For " & Format(DateClicked, "MM/DD/YYYY")
Sql = "Select * from [Reservations] Where TRY_CAST([date] AS datetime) = '" & CDate(Data(0).Text) & "'"
Sql = Sql & " Order by  TRY_CAST([Time] AS datetime),[Name] asc"
LoadGrid Sql, False
End Sub
Private Sub Reserve_Click()
Reserve.HighLight = flexHighlightNever
With Reserve
   If .TextMatrix(.Row, 5) = "" Then Exit Sub
   Data(1).Text = .TextMatrix(.Row, 0)
   Data(2).Text = .TextMatrix(.Row, 2)
   Data(3).Text = .TextMatrix(.Row, 1)
   Data(4).Text = .TextMatrix(.Row, 3)
   Data(5).Text = .TextMatrix(.Row, 4)
   Data(0).Text = .TextMatrix(.Row, 5)
   Data(6).Text = .TextMatrix(.Row, 6)
   Data(7).Text = .TextMatrix(.Row, 7)
End With
If Check_Option(208) = 1 Then
    If FindReservationOrder(Data(6).Text, Data(1).Text, Data(0).Text) = True Then
     BtnEnh6.Visible = True
     BtnEnh1.Visible = False
    Else
     BtnEnh6.Visible = False
     BtnEnh1.Visible = True
    End If
End If
End Sub

