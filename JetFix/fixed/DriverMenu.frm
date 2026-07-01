VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form DriverMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00352F2F&
   BorderStyle     =   0  'None
   Caption         =   "Driver Report"
   ClientHeight    =   11130
   ClientLeft      =   3975
   ClientTop       =   450
   ClientWidth     =   5535
   FillColor       =   &H00352F2F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "DriverMenu.frx":0000
      Top             =   8160
      Width           =   4815
   End
   Begin BTNENHLib4.BtnEnh exit 
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   9960
      Width           =   2295
      _Version        =   589828
      _ExtentX        =   4048
      _ExtentY        =   1508
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
      textCaption     =   "DriverMenu.frx":0050
      textLT          =   "DriverMenu.frx":00B8
      textCT          =   "DriverMenu.frx":00D0
      textRT          =   "DriverMenu.frx":00E8
      textLM          =   "DriverMenu.frx":0100
      textRM          =   "DriverMenu.frx":0118
      textLB          =   "DriverMenu.frx":0130
      textCB          =   "DriverMenu.frx":0148
      textRB          =   "DriverMenu.frx":0160
      colorBack       =   "DriverMenu.frx":0178
      colorIntern     =   "DriverMenu.frx":01A2
      colorMO         =   "DriverMenu.frx":01CC
      colorFocus      =   "DriverMenu.frx":01F6
      colorDisabled   =   "DriverMenu.frx":0220
      colorPressed    =   "DriverMenu.frx":024A
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
   Begin VSFlex8DAOCtl.VSFlexGrid Ticket 
      Height          =   7575
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4890
      _cx             =   1998201265
      _cy             =   1998206001
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   4194304
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65535
      ForeColorSel    =   0
      BackColorBkg    =   3485487
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   4
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   600
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      ComboSearch     =   0
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
   Begin BTNENHLib4.BtnEnh SSCommand21 
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   9000
      Width           =   2295
      _Version        =   589828
      _ExtentX        =   4048
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Cash Drop"
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
      textCaption     =   "DriverMenu.frx":0274
      textLT          =   "DriverMenu.frx":02E6
      textCT          =   "DriverMenu.frx":02FE
      textRT          =   "DriverMenu.frx":0316
      textLM          =   "DriverMenu.frx":032E
      textRM          =   "DriverMenu.frx":0346
      textLB          =   "DriverMenu.frx":035E
      textCB          =   "DriverMenu.frx":0376
      textRB          =   "DriverMenu.frx":038E
      colorBack       =   "DriverMenu.frx":03A6
      colorIntern     =   "DriverMenu.frx":03D0
      colorMO         =   "DriverMenu.frx":03FA
      colorFocus      =   "DriverMenu.frx":0424
      colorDisabled   =   "DriverMenu.frx":044E
      colorPressed    =   "DriverMenu.frx":0478
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
      Office2003ColorNorth=   12870144
      Office2003ColorSouth=   8404992
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   12870144
      Office2003ColorSouthMouseOver=   12870144
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   12870144
      Office2003ColorSouthPressed=   8404992
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh2 
      Height          =   855
      Left            =   2880
      TabIndex        =   3
      Top             =   9000
      Width           =   2295
      _Version        =   589828
      _ExtentX        =   4048
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Collect Delivery Charges"
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
      textCaption     =   "DriverMenu.frx":04A2
      textLT          =   "DriverMenu.frx":0532
      textCT          =   "DriverMenu.frx":054A
      textRT          =   "DriverMenu.frx":0562
      textLM          =   "DriverMenu.frx":057A
      textRM          =   "DriverMenu.frx":0592
      textLB          =   "DriverMenu.frx":05AA
      textCB          =   "DriverMenu.frx":05C2
      textRB          =   "DriverMenu.frx":05DA
      colorBack       =   "DriverMenu.frx":05F2
      colorIntern     =   "DriverMenu.frx":061C
      colorMO         =   "DriverMenu.frx":0646
      colorFocus      =   "DriverMenu.frx":0670
      colorDisabled   =   "DriverMenu.frx":069A
      colorPressed    =   "DriverMenu.frx":06C4
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
      Office2003ColorNorth=   12870144
      Office2003ColorSouth=   8404992
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   12870144
      Office2003ColorSouthMouseOver=   12870144
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   12870144
      Office2003ColorSouthPressed=   8404992
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   9960
      Width           =   2295
      _Version        =   589828
      _ExtentX        =   4048
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Print Report"
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
      textCaption     =   "DriverMenu.frx":06EE
      textLT          =   "DriverMenu.frx":0766
      textCT          =   "DriverMenu.frx":077E
      textRT          =   "DriverMenu.frx":0796
      textLM          =   "DriverMenu.frx":07AE
      textRM          =   "DriverMenu.frx":07C6
      textLB          =   "DriverMenu.frx":07DE
      textCB          =   "DriverMenu.frx":07F6
      textRB          =   "DriverMenu.frx":080E
      colorBack       =   "DriverMenu.frx":0826
      colorIntern     =   "DriverMenu.frx":0850
      colorMO         =   "DriverMenu.frx":087A
      colorFocus      =   "DriverMenu.frx":08A4
      colorDisabled   =   "DriverMenu.frx":08CE
      colorPressed    =   "DriverMenu.frx":08F8
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
   Begin VB.Label DVRName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00352F2F&
      FillStyle       =   0  'Solid
      Height          =   10815
      Left            =   120
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "DriverMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DeliveryFees      As Currency
Private TotalOWED         As Currency
Private AllDeliveryChgs   As Currency
Private StartingAuditNo   As Long
Private AuditStarted      As String
Private Sub BtnEnh1_Click()
On Error GoTo Printerr
Dim Db              As ADODB.Connection
Dim OP              As ADODB.Recordset
Dim PDB             As ADODB.Connection
Dim PrinterPort     As String
Dim DeliveryCount   As Long
Dim SvrCharges      As Currency
Dim DeliveryTotal   As Currency
Dim OCHECKSales     As Currency
Dim CCHECKSales     As Currency
Dim R               As Integer
Dim Net_Disc        As Currency
Dim Net_STax        As Currency
Dim Net_SPTax       As Currency
Dim Net_LTax        As Currency
Dim NetSale         As Currency
Dim checkInfo()     As String
Dim CheckT          As Currency
Dim PF              As Integer
Dim Pa              As Integer
Dim Heading         As String
Dim ExpPrint        As Integer
Dim DataReturned    As String
PF = FreeFile
ConnectADODB CheckStatDb, PDB
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
  Sql = "Select * from [Open Checks] where [Server Name] = '" & DVRName.Caption & "' Order by [Server Name],[Check Number] asc"
ConnectADORS Sql, Db, OP
Dim GD As GeneralLib.Utilities
Set GD = New GeneralLib.Utilities
ReDim checkInfo(4, RecordsCount(OP))
R = -1
OCHECKSales = 0
If Not OP.EOF Then
    R = R + 1
    checkInfo(0, R) = "Open Checks"
    Do While Not OP.EOF
       R = R + 1
       checkInfo(0, R) = Left$(OP![Street] & "", 30)  'table & vbcrlf
       checkInfo(1, R) = OP(2)  'Check
       checkInfo(2, R) = OP(4) & ""  'time
       Net_Disc = OP(10) - OP(15)
       Net_STax = OP(7) - OP(12)
       Net_SPTax = OP(8) - OP(13)
       Net_LTax = OP(9) - OP(14)
       NetSale = OP(6) - OP(11)
       CheckT = (NetSale + Net_STax + Net_SPTax + Net_LTax) - Net_Disc
       checkInfo(3, R) = CheckT
       checkInfo(4, R) = 0
       OCHECKSales = OCHECKSales + CheckT
       OP.MoveNext
    Loop
End If
DbClose Db, OP
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Select * from [Closed Checks] where [Server Name] = '" & DVRName.Caption & "'"
  Sql = Sql & " and [Close Out Day] = 0 and [Check Total] > 0"
  Sql = Sql & " and [Order Destination] = 'Delivery'"
ConnectADORS Sql, Db, OP
DeliveryCount = 0
CCHECKSales = 0
If Not OP.EOF Then
   ReDim Preserve checkInfo(4, RecordsCount(OP) + (R + 1))
    R = R + 1
    checkInfo(0, R) = "Closed Checks"
    Do While Not OP.EOF
        R = R + 1
        DeliveryCount = DeliveryCount + 1
        checkInfo(0, R) = GetDeliveryAddress("", OP![Check Number])
        checkInfo(1, R) = OP![Check Number] 'Check
        checkInfo(2, R) = Format(OP![Close Time] & "", "H:M ampm") 'time
        checkInfo(3, R) = Format(OP![Check Total] & "", "0.00")
        checkInfo(4, R) = Val(GD.ReturnTipAmount(PDB, CStr(checkInfo(1, R))))
        CCHECKSales = CCHECKSales + Format(OP![Check Total] & "", "0.00")
       OP.MoveNext
    Loop
End If
DbClose Db, OP
PDB.Close
Set GD = Nothing
Heading = "Driver Report"
On Error GoTo Printerr
Printer_Drivers "Receipt"
tryagain:
Open "Report" For Output Shared As #PF
Print #PF, EC(1); EC(3); EC(8); Heading; EC(4); EC(7)
Print #PF, EC(1); EC(3); EC(8); "Delivery Tickets"; EC(4); EC(7)
Print #PF, EC(9); EC(7)
Print #PF, ; EC(7)
Print #PF, EC(2); EC(3); "Driver Name : "; DVRName.Caption; EC(2); EC(7)
Print #PF, Format(Time, "HH:MM AmPm");
Print #PF, Tab(24); "Date  "; Trim(Date); EC(7)
Print #PF, String(40, 45); EC(7)
Print #PF, "Address"; EC(7)
Print #PF, "Chk#       Time     Chk Ttl    Tip"; EC(7)
For Pa = 0 To R
 If checkInfo(0, Pa) = "Open Checks" Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Open Check List"; EC(7)
    Print #PF, String(40, 45); EC(7)
 ElseIf checkInfo(0, Pa) = "Closed Checks" Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Closed Check List"; EC(7)
    Print #PF, String(40, 45); EC(7)
 Else
    If checkInfo(0, Pa) > "" Then Print #PF, checkInfo(0, Pa) & EC(7)
    Print #PF, checkInfo(1, Pa); Tab(10); checkInfo(2, Pa);
    Print #PF, Tab(20); Format(Abs(checkInfo(3, Pa)), "0.00") & " "; Tab(32); Format(Val(checkInfo(4, Pa)), "0.00"); EC(7)
 End If
Next Pa
Print #PF, String(40, 45); EC(7)
For Pa = 0 To Ticket(0).rows
 If Ticket(0).TextMatrix(Pa, 2) = "*" Then
   ExpPrint = 1
 Else
   ExpPrint = 2
 End If
 If Ticket(0).TextMatrix(Pa, 0) = "" Then Exit For
 If Ticket(0).TextMatrix(Pa, 1) = "" Then
  Print #PF, EC(ExpPrint); EC(3); EC(8); Ticket(0).TextMatrix(Pa, 0); EC(4); EC(7)
 Else
 If Ticket(0).TextMatrix(Pa, 0) = "Total Deliveries" Or Ticket(0).TextMatrix(Pa, 0) = "Total Credit Card" Or Ticket(0).TextMatrix(Pa, 0) = "Total PrePaid" Then
  Print #PF, EC(ExpPrint); EC(3); Ticket(0).TextMatrix(Pa, 0); Tab(38); Val(Ticket(0).TextMatrix(Pa, 1)); EC(4); EC(7)
 Else
  Print #PF, EC(ExpPrint); EC(3); Ticket(0).TextMatrix(Pa, 0); Tab(35); Align(7, Val(Ticket(0).TextMatrix(Pa, 1))); EC(4); EC(7)
 End If
 End If
Next Pa
Print #PF, EC(1); String(40, 45); EC(7)
Print #PF, EC(6); EC(6); EC(5)
Close #PF
DataReturned = SendJobToPrinter("Report", "Receipt", True)
SaveReport DataReturned, DVRName.Caption, 1, Format(Date, "MM/DD/YYYY"), Format(Time, "HH:MM ampm")
'add 07.07.2022 for Tracy
Exit Sub
Printerr:
Call See_Msg(Err.Description, 0, "Error..")
Close
Exit Sub
End Sub
Private Sub BtnEnh2_Click()
If Can_Continue(BtnEnh2.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim CollectFees As Currency
Dim Db                As ADODB.Connection
Dim OP                As ADODB.Recordset
CollectFees = ShowNumberPad("Enter Delivery Charge Amt", False, 0, 0)
If CollectFees = 0 Then Exit Sub
If CollectFees > DeliveryFees Then
   See_Msg "Cannot Collect More Than Total Charge Amount", 5, "Invalid Amount"
   Exit Sub
End If
If See_Msg("New Amount Is " & Format(CollectFees, "0.00"), 4, "Ok To Save ?") <> 6 Then Exit Sub
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
Sql = "Select * from [Deposits Paid In] where [Driver Name] = '" & DVRName.Caption & "'"
ConnectADORS Sql, Db, OP
 If OP.EOF Then OP.AddNew
 OP![Driver Name] = DVRName.Caption
 OP![Delivery Charge Collected] = Val(OP![Delivery Charge Collected] & "") + CollectFees
 If IsNull(OP![Cash Deposit]) Then OP![Cash Deposit] = 0
 OP.Update
 DbClose Db, OP
 ShowDriverSales DVRName.Caption, GetStartingTime(DVRName.Caption), GetStartingCheck(DVRName.Caption)
 Exit Sub
ErrorHandler:
See_Msg Err.Description, 0, "Error"
Error_Log Err.Number, Err.Description, "Drive Menu", "None"
Err.Clear
End Sub
Private Sub Exit_Click()
Call MemoryCheck(Me)
End Sub

Private Sub Form_Activate()
'AuditStarted = GetStartingTime(DVRName.Caption)
ShowDriverSales GetSelected, GetStartingTime(DVRName.Caption), GetStartingCheck(DVRName.Caption)
End Sub

Private Sub Form_Load()
ResizeScreen1024 Me
CenterForm Me
With Ticket(0)
    .HighLight = flexHighlightNever
    .rows = 200
    .cols = 3
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColWidth(0) = (3600 * WidScale1024)
    .ColWidth(1) = (1250 * WidScale1024)
    .ColWidth(2) = 0
End With
End Sub
Private Function GetStartingTime(Driver As String) As String
Dim Db                As ADODB.Connection
Dim ATB               As ADODB.Recordset
GetStartingTime = "" ' AuditStarted = Now
Dim StartAudit  As String
If System_config = ExpressService Then
    
    ConnectADODB SvrPath.TerminalPath & DebitDb, Db
     Sql = "Select * From [Debit] where [Server Name] = '" & Driver & "' and [close out day] = 0"
     Sql = Sql & " Order by TRY_CAST([Date] AS datetime) , TRY_CAST([Time] AS datetime) asc"
    ConnectADORS Sql, Db, ATB
    If Not ATB.EOF Then
       StartAudit = ATB![Date] & " " & ATB![Time] & ""
    Else
      StartAudit = ""
    End If
    DbClose Db, ATB
    
    ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
     Sql = "Select * From [DayPayments] where [Server Name] = '" & Driver & "'"
     Sql = Sql & " Order by TRY_CAST([Recorded Date] AS datetime) , TRY_CAST([Recorded Time] AS datetime) asc"
    ConnectADORS Sql, Db, ATB
     If Not ATB.EOF Then
      GetStartingTime = ATB![Recorded Date] & " " & Format(ATB![Recorded Time] & "", "HH:MM ampm")
      If IsDate(StartAudit) Then
       If CDate(StartAudit) < CDate(GetStartingTime) Then
        GetStartingTime = StartAudit
       End If
      End If
     Else
      If IsDate(StartAudit) Then
       GetStartingTime = StartAudit
      Else
        Call ConnectADODB(CPath.TerminalPath & CallInOrderDb, Db)
          Sql = "Select * from [Open Checks] Where [Driver Assigned] = '" & Driver & "'"
          Sql = Sql & " and isdate([Pick Up date] + [Time Assigned]) = 0"
          Sql = Sql & "Order by TRY_CAST([Pick Up Date] AS datetime) , TRY_CAST([Time assigned] AS datetime) asc"
          Call ConnectADORS(Sql, Db, ATB)
           If Not ATB.EOF Then
             GetStartingTime = ATB![Pick Up Date] & " " & ATB![Time assigned] & ""
           Else
             GetStartingTime = Now
             See_Msg "No payments found for ", 5, Driver & ""
             Trans_Log "No Payments Audit Start Time " & Driver & " " & GetStartingTime
           End If
          DbClose Db, ATB
      End If
     End If
Else
    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    Sql = "Select * From [Check Numbers] where [Servers Name] = '" & Driver & "' and [Audit Status] = 0"
    ConnectADORS Sql, Db, ATB
    If Not ATB.EOF Then
       GetStartingTime = Format(ATB![Start date], "MM/DD/YYYY") & " " & ATB![Time] & ""
    End If
End If
DbClose Db, ATB
Trans_Log "Audit Start Time " & Driver & " " & GetStartingTime
End Function
Private Function GetStartingCheck(Driver As String) As Long
Dim Db                As ADODB.Connection
Dim ATB               As ADODB.Recordset
GetStartingCheck = 0
If System_config = ExpressService Then
    
Else
    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    Sql = "Select * From [Check Numbers] where [Servers Name] = '" & Driver & "' and [Audit Status] = 0"
    ConnectADORS Sql, Db, ATB
    If Not ATB.EOF Then
      GetStartingCheck = ATB![Check No]
    End If
    DbClose Db, ATB
End If
End Function
Public Sub ShowDriverSales(ForWho As String, AuditStart As String, StartCheck As Long)
On Error GoTo ErrorHandler
Dim Db                As ADODB.Connection
Dim OP                As ADODB.Recordset
Dim PrinterPort       As String
Dim DeliveryCount     As Long
Dim SvrCharges        As Currency
Dim DeliveryTotal     As Currency
Dim Net_Disc          As Currency
Dim Net_STax          As Currency
Dim Net_SPTax         As Currency
Dim Net_LTax          As Currency
Dim NetSale           As Currency
Dim R                 As Integer
Dim L                 As Integer
Dim CCHECKSales       As Currency
Dim DepositsPaidIn    As Currency
Dim DvrChgCollect     As Currency
Dim DeliverySalesAmt  As Currency
Dim DeliveryAmountDue As Currency
Dim C_Charges         As Currency
Dim C_Tips            As Currency
Dim C_Count           As Long
Dim DrvGrat           As Currency
Dim PP_Charges        As Currency
Dim PP_Tips           As Currency
Dim PP_Count          As Integer
Dim Ol_Tip            As Currency
Dim Ol_Count          As Integer
Dim TotalPPTips       As Currency
DeliveryCount = GetDeliveryCount(ForWho, AuditStart)
DeliverySalesAmt = GetDriverSales(ForWho, AuditStart)
SvrCharges = GetServiceCharges(ForWho, StartCheck, AuditStart)
DrvGrat = GetDriverGratuity(ForWho, StartCheck, AuditStart)

   If SvrCharges > 0 Then
     If Val(Input_Option(59)) > 0 Then
       If InStr(Input_Option(59), ".") = 0 Then  'used for Per of delivery
          DeliveryAmountDue = SvrCharges * (Val(Input_Option(59) / 100))
       Else
          SvrCharges = 0
       End If
     End If
   End If
   If SvrCharges = 0 Then
     If Val(Input_Option(59)) > 0 Then
       If InStr(Input_Option(59), ".") = 0 Then
         DeliveryAmountDue = DeliverySalesAmt * (Val(Input_Option(59) / 100))
       End If
     End If
   End If
   If SvrCharges = 0 Then
     If Val(Input_Option(59)) > 0 Then
       If InStr(Input_Option(59), ".") > 0 Then
         DeliveryAmountDue = Format(Val(DeliveryCount) * Val(Input_Option(59)), "0.00")
       End If
     End If
   End If
'If EMVDeviceIsPresent Then
 ConnectADODB SvrPath.TerminalPath & DebitDb, Db
  Sql = "Select * from [debit] where [Server Name] = '" & ForWho & "'"
  If StartCheck = 0 Then
   If GetConfigInfo("Phone Orders", "Search EMV Card By Close Out Day", "No", "Functions.INI") = "ENABLED" Then
   Else
      Sql = Sql & " and TRY_CAST([Date] AS datetime) + TRY_CAST([Time] AS datetime) >= '" & AuditStart & "'"
   End If
  Else
    Sql = Sql & " and [Check Number] >= " & StartCheck
  End If
  Sql = Sql & " and [Close out day] = 0"
 ConnectADORS Sql, Db, OP
 C_Count = 0
    Do While Not OP.EOF
       C_Charges = C_Charges + Format(OP![Amount] + OP![tip], "0.00")
       C_Tips = C_Tips + Format(OP![tip], "0.00")
       C_Count = C_Count + 1
      OP.MoveNext
    Loop
'End If
DbClose Db, OP

ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
  Sql = "Select [Server Name], Sum([Payment Amount]) as TSale, Sum([Tip Amount]) as TGrat, Count([Payment Amount]) as TCrds from [DayPayments] where [Server Name] = '" & ForWho & "'"
  Sql = Sql & " and TRY_CAST([Recorded Date] AS datetime) + TRY_CAST([Recorded Time] AS datetime) >= '" & AuditStart & "'"
  Sql = Sql & " and [Payment Type] <> 'Cash'"
  Sql = Sql & " and [Check Number] <> -9999"
  Sql = Sql & " and [Tip Paid] = 0 Group by [Server Name]"
  
ConnectADORS Sql, Db, OP
PP_Charges = 0: PP_Tips = 0: PP_Count = 0
 Do While Not OP.EOF
    PP_Charges = PP_Charges + Format(OP![TSale] + OP![TGrat], "0.00")
    PP_Tips = PP_Tips + Format(OP![TGrat], "0.00")
    PP_Count = PP_Count + Format(OP![TCrds], "0.00")
    OP.MoveNext
 Loop
DbClose Db, OP
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Select [Servers Name], Sum([Price]) as TTip, Count([Price]) as TCrds from [Check Detail] where [Servers Name] = '" & ForWho & "'"
  Sql = Sql & " and [Department Report] = 'Online Tip' and [Close out Day] = 0"
  Sql = Sql & " and TRY_CAST([Check Date] AS datetime) + TRY_CAST([Recorded Time] AS datetime) >= '" & AuditStart & "'"
  Sql = Sql & " Group by [Servers Name]"
ConnectADORS Sql, Db, OP
 Ol_Tip = 0: Ol_Count = 0
 Do While Not OP.EOF
    Ol_Tip = Ol_Tip + Format(OP![TTip], "0.00")
    Ol_Count = Ol_Count + Format(OP![TCrds], "0.00")
    OP.MoveNext
 Loop
DbClose Db, OP

ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
   Sql = "Select * from [Open Checks] where [Server Name] = '" & ForWho & "'"
   'Sql = Sql & " and [Driver Assigned] = '" & ForWho & "'"
   
ConnectADORS Sql, Db, OP
Net_Disc = 0: Net_STax = 0: Net_SPTax = 0: Net_LTax = 0: NetSale = 0
If Not OP.EOF Then
    Do While Not OP.EOF
       Net_Disc = Net_Disc + (OP(10) - OP(15))
       Net_STax = Net_STax + (OP(7) - OP(12))
       Net_SPTax = Net_SPTax + (OP(8) - OP(13))
       Net_LTax = Net_LTax + (OP(9) - OP(14))
       NetSale = NetSale + (OP(6) - OP(11))
       OP.MoveNext
    Loop
End If
SetNothing OP
Sql = "Select * from [Deposits Paid In] where [Driver Name] = '" & ForWho & "'"
ConnectADORS Sql, Db, OP
DvrChgCollect = 0: DepositsPaidIn = 0
Do While Not OP.EOF
   DvrChgCollect = DvrChgCollect + Format(Val(OP![Delivery Charge Collected] & ""), "0.00")
   DepositsPaidIn = DepositsPaidIn + Format(Val(OP![Cash Deposit] & ""), "0.00")
  OP.MoveNext
Loop
L = 0
Dim TotalSales As Currency
TotalSales = ((((NetSale + (Net_STax + (Net_SPTax + Net_LTax)))) - Net_Disc))
With Ticket(0)
    .TextMatrix(L, 0) = "*** Open Phone Orders ***"
    .TextMatrix(L, 1) = ""
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Net Open Sales"
    .TextMatrix(L, 1) = Format(NetSale, "0.00")
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Open Discounts"
    .TextMatrix(L, 1) = Format(Net_Disc, "0.00")
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Open Sales Tax"
    .TextMatrix(L, 1) = Format(((Net_STax + Net_SPTax) + Net_LTax), "0.00")
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Total Open Sales"
    .TextMatrix(L, 1) = Format(TotalSales, "0.00")
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "*** Delivery Sales ***"
    .TextMatrix(L, 1) = ""
    .TextMatrix(L, 2) = ""
    If TotalSales = 0 Then
     L = L + 1
     .TextMatrix(L, 0) = "Total Delivery Sales"
     .TextMatrix(L, 1) = Format(DeliverySalesAmt, "0.00")
     .TextMatrix(L, 2) = ""
    Else
     L = L + 1
     .TextMatrix(L, 0) = "*** Total Delivery Sales ***"
     .TextMatrix(L, 1) = ""
     .TextMatrix(L, 2) = ""
     L = L + 1
     .TextMatrix(L, 0) = "Total Sales Open & Closed"
     .TextMatrix(L, 1) = Format(DeliverySalesAmt + TotalSales, "0.00")
     .TextMatrix(L, 2) = "*"
    End If
    L = L + 1
    .TextMatrix(L, 0) = "Cash Drop"
    .TextMatrix(L, 1) = Format(DepositsPaidIn, "0.00")
    .TextMatrix(L, 2) = ""
    L = L + 1
   .TextMatrix(L, 0) = "Credit Card Charges"
   .TextMatrix(L, 1) = Format(C_Charges - C_Tips, "0.00")
   .TextMatrix(L, 2) = "*"
    L = L + 1
   .TextMatrix(L, 0) = "Credit Card Tips"
   .TextMatrix(L, 1) = Format(C_Tips, "0.00")
   .TextMatrix(L, 2) = "*"
    L = L + 1
   .TextMatrix(L, 0) = "PrePaid Charges"
   .TextMatrix(L, 1) = Format(PP_Charges - PP_Tips, "0.00")
   .TextMatrix(L, 2) = "*"
    L = L + 1
   .TextMatrix(L, 0) = "PrePaid Charge Tips"
   .TextMatrix(L, 1) = Format(PP_Tips + Ol_Tip, "0.00")
   .TextMatrix(L, 2) = "*"
    TotalPPTips = PP_Tips + Ol_Tip
    L = L + 1
    C_Charges = C_Charges + PP_Charges
    'C_Charges = C_Charges + (PP_Charges - (PP_Tips + Ol_Tip))
   .TextMatrix(L, 0) = "Total Charges"
   .TextMatrix(L, 1) = Format(C_Charges, "0.00")
   .TextMatrix(L, 2) = "*"
    L = L + 1
   .TextMatrix(L, 0) = "Delivery Charges Paid"
   .TextMatrix(L, 1) = Format(DvrChgCollect, "0.00")
   .TextMatrix(L, 2) = ""
    L = L + 1
   .TextMatrix(L, 0) = "Total Delivery Charges"
   .TextMatrix(L, 1) = Format(DeliveryAmountDue, "0.00")
   .TextMatrix(L, 2) = "*"
    AllDeliveryChgs = (DeliveryAmountDue)
    If DepositsPaidIn > 0 Then
     L = L + 1
     .TextMatrix(L, 0) = "Amount Due Less Drop"
     .TextMatrix(L, 1) = Format((DeliverySalesAmt + TotalSales) - DepositsPaidIn, "0.00")
    End If
    L = L + 1
    .TextMatrix(L, 0) = "Net Del. Chg Due Driver"
    .TextMatrix(L, 1) = Format(DeliveryAmountDue - DvrChgCollect, "0.00")
    DeliveryFees = Val(.TextMatrix(L, 1))
    .TextMatrix(L, 2) = "*"
    If DrvGrat > 0 Then
     L = L + 1
    .TextMatrix(L, 0) = "Net Grat. Due Driver"
    .TextMatrix(L, 1) = Format(DrvGrat, "0.00")
    .TextMatrix(L, 2) = "*"
    End If
    L = L + 1
    .TextMatrix(L, 0) = "Net Cash Due (Restaurant)"
    .TextMatrix(L, 1) = Format((((DeliverySalesAmt + TotalSales) - DepositsPaidIn) - (((DeliveryAmountDue + DrvGrat) + TotalPPTips) - DvrChgCollect) - C_Charges), "0.00")
    TotalOWED = Val(.TextMatrix(L, 1))
    .TextMatrix(L, 2) = "*"
    L = L + 1
    .TextMatrix(L, 0) = "Total Deliveries"
    .TextMatrix(L, 1) = DeliveryCount
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Total Credit Card"
    .TextMatrix(L, 1) = C_Count
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Total PrePaid"
    .TextMatrix(L, 1) = PP_Count
    .TextMatrix(L, 2) = ""
    L = L + 1
    .TextMatrix(L, 0) = "Amount Due Driver"
    .TextMatrix(L, 1) = (DrvGrat + AllDeliveryChgs) + ((C_Tips + PP_Tips) + Ol_Tip)
    .TextMatrix(L, 2) = ""

End With
DbClose Db, OP
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "driver report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub ManagerEdit(DropsOrSvrCharge As String, TotalIn As Currency)
On Error GoTo ErrorHandler
Dim CollectFees As Currency
Dim Db                As ADODB.Connection
Dim OP                As ADODB.Recordset
CollectFees = ShowNumberPad("Enter New Amount", False, 0, 0)
If CollectFees = 0 Then Exit Sub
If CollectFees > TotalIn Then
   See_Msg "Cannot Deposit More Than Total Amount", 5, "Invalid Amount"
   Exit Sub
End If
If See_Msg("New Amount Is " & Format(CollectFees, "0.00"), 4, "Ok To Save ?") <> 6 Then Exit Sub
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
  Sql = "Select * from [Deposits Paid In] where [Driver Name] = '" & DVRName.Caption & "'"
ConnectADORS Sql, Db, OP
 If OP.EOF Then OP.AddNew
 OP![Driver Name] = DVRName.Caption
 If DropsOrSvrCharge = "Delivery Charges Paid" Then
  OP![Delivery Charge Collected] = CollectFees
 Else
  OP![Cash Deposit] = CollectFees
 End If
 OP.Update
 DbClose Db, OP
 ShowDriverSales DVRName.Caption, GetStartingTime(DVRName.Caption), GetStartingCheck(DVRName.Caption)
 Exit Sub
ErrorHandler:
See_Msg Err.Description, 0, "Error"
Error_Log Err.Number, Err.Description, "Drive Menu", "None"
Err.Clear

End Sub
Private Sub SSCommand21_Click()
If Can_Continue(SSCommand21.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim CollectFees As Currency
Dim Db                As ADODB.Connection
Dim OP                As ADODB.Recordset
CollectFees = ShowNumberPad("Enter Cash Amount", False, 0, 0)
If CollectFees = 0 Then Exit Sub
If CollectFees > TotalOWED Then
   See_Msg "Cannot Deposit More Than Total Amount", 5, "Invalid Amount"
   Exit Sub
End If
If See_Msg("New Amount Is " & Format(CollectFees, "0.00"), 4, "Ok To Save ?") <> 6 Then Exit Sub
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
  Sql = "Select * from [Deposits Paid In] where [Driver Name] = '" & DVRName.Caption & "'"
ConnectADORS Sql, Db, OP
 If OP.EOF Then OP.AddNew
 OP![Driver Name] = DVRName.Caption
 If IsNull(OP![Delivery Charge Collected]) Then OP![Delivery Charge Collected] = 0
 OP![Cash Deposit] = Val(OP![Cash Deposit] & "") + CollectFees
 OP.Update
 DbClose Db, OP
 ShowDriverSales DVRName.Caption, GetStartingTime(DVRName.Caption), GetStartingCheck(DVRName.Caption)
 Exit Sub
ErrorHandler:
See_Msg Err.Description, 0, "Error"
Error_Log Err.Number, Err.Description, "Drive Menu", "None"
Err.Clear
End Sub
Private Sub Ticket_Click(Index As Integer)
If Can_Continue("Managers Menu") = False Then Exit Sub
With Ticket(0)
   If .TextMatrix(.Row, 0) = "Cash Drop" Then
       ManagerEdit .TextMatrix(.Row, 0), TotalOWED
   ElseIf .TextMatrix(.Row, 0) = "Delivery Charges Paid" Then
      ManagerEdit .TextMatrix(.Row, 0), AllDeliveryChgs
   End If
End With
End Sub
