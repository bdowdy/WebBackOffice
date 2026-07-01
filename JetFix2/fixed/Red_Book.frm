VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Object = "{84E5CF37-E467-4AC2-89C4-C6002FFB5055}#25.1#0"; "CHARTV~1.OCX"
Begin VB.Form Red_Book 
   Appearance      =   0  'Flat
   BackColor       =   &H00302821&
   BorderStyle     =   0  'None
   Caption         =   "Red Book"
   ClientHeight    =   10680
   ClientLeft      =   195
   ClientTop       =   465
   ClientWidth     =   15195
   FillColor       =   &H00302821&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00302821&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10680
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox HourlyData 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   6360
      ScaleHeight     =   3225
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   4575
      Begin BTNENHLib4.BtnEnh BtnEnh28 
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4095
         _Version        =   589828
         _ExtentX        =   7223
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Week To Date"
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
         Surface         =   15
         BackColorContainer=   12632256
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0000
         textLT          =   "Red_Book.frx":0078
         textCT          =   "Red_Book.frx":0090
         textRT          =   "Red_Book.frx":00A8
         textLM          =   "Red_Book.frx":00C0
         textRM          =   "Red_Book.frx":00D8
         textLB          =   "Red_Book.frx":00F0
         textCB          =   "Red_Book.frx":0108
         textRB          =   "Red_Book.frx":0120
         colorBack       =   "Red_Book.frx":0138
         colorIntern     =   "Red_Book.frx":0162
         colorMO         =   "Red_Book.frx":018C
         colorFocus      =   "Red_Book.frx":01B6
         colorDisabled   =   "Red_Book.frx":01E0
         colorPressed    =   "Red_Book.frx":020A
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
         Office2003ColorNorth=   12632256
         Office2003ColorSouth=   12632256
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   12632256
         Office2003ColorSouthMouseOver=   10485760
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   8421504
         Office2003ColorSouthPressed=   4194304
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh26 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   4335
         _Version        =   589828
         _ExtentX        =   7646
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
         Surface         =   15
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0234
         textLT          =   "Red_Book.frx":029C
         textCT          =   "Red_Book.frx":02B4
         textRT          =   "Red_Book.frx":02CC
         textLM          =   "Red_Book.frx":02E4
         textRM          =   "Red_Book.frx":02FC
         textLB          =   "Red_Book.frx":0314
         textCB          =   "Red_Book.frx":032C
         textRB          =   "Red_Book.frx":0344
         colorBack       =   "Red_Book.frx":035C
         colorIntern     =   "Red_Book.frx":0386
         colorMO         =   "Red_Book.frx":03B0
         colorFocus      =   "Red_Book.frx":03DA
         colorDisabled   =   "Red_Book.frx":0404
         colorPressed    =   "Red_Book.frx":042E
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
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh27 
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   4095
         _Version        =   589828
         _ExtentX        =   7223
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Hourly Sales"
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
         Surface         =   15
         BackColorContainer=   16744576
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0458
         textLT          =   "Red_Book.frx":04D0
         textCT          =   "Red_Book.frx":04E8
         textRT          =   "Red_Book.frx":0500
         textLM          =   "Red_Book.frx":0518
         textRM          =   "Red_Book.frx":0530
         textLB          =   "Red_Book.frx":0548
         textCB          =   "Red_Book.frx":0560
         textRB          =   "Red_Book.frx":0578
         colorBack       =   "Red_Book.frx":0590
         colorIntern     =   "Red_Book.frx":05BA
         colorMO         =   "Red_Book.frx":05E4
         colorFocus      =   "Red_Book.frx":060E
         colorDisabled   =   "Red_Book.frx":0638
         colorPressed    =   "Red_Book.frx":0662
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
         Office2003ColorNorth=   12632256
         Office2003ColorSouth=   12632256
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   12632256
         Office2003ColorSouthMouseOver=   10485760
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   12632256
         Office2003ColorSouthPressed=   4194304
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hourly Sales By Terminal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1935
         Left            =   120
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.Frame Frame20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   11090
      TabIndex        =   80
      Top             =   720
      Width           =   3900
      Begin VSFlex8DAOCtl.VSFlexGrid HourlyGrid 
         Height          =   3850
         Left            =   80
         TabIndex        =   81
         Top             =   480
         Width           =   3750
         _cx             =   1990138327
         _cy             =   1990138503
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   15132390
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16316664
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   16777215
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   2
         Rows            =   200
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Red_Book.frx":068C
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
      Begin VB.Shape Shape8 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   0
         Top             =   350
         Width           =   3900
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales by the Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   110
         Width           =   2775
      End
   End
   Begin VB.Frame Frame19 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4655
      Left            =   11090
      TabIndex        =   78
      Top             =   5600
      Width           =   3900
      Begin CDChartViewer.ChartViewer MealPeriods 
         Height          =   1815
         Left            =   0
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3201
         Appearance      =   1
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Periods"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   110
         Width           =   1455
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   0
         Top             =   345
         Width           =   3900
      End
   End
   Begin VB.Frame Frame18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   7140
      TabIndex        =   77
      Top             =   720
      Width           =   3900
      Begin CDChartViewer.ChartViewer OrderDest 
         Height          =   1815
         Left            =   0
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
         Appearance      =   1
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   0
         Top             =   350
         Width           =   3975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   0
         Left            =   0
         Top             =   345
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Destination"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   110
         Width           =   1935
      End
   End
   Begin VB.Frame Frame17 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3475
      TabIndex        =   76
      Top             =   4520
      Width           =   3615
      Begin CDChartViewer.ChartViewer LD 
         Height          =   2415
         Left            =   0
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         Appearance      =   1
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor vs. Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   110
         Width           =   2055
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   0
         Top             =   350
         Width           =   4095
      End
   End
   Begin VB.Frame Frame16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4655
      Left            =   7140
      TabIndex        =   75
      Top             =   5600
      Width           =   3900
      Begin CDChartViewer.ChartViewer LaborBreakdown 
         Height          =   2415
         Left            =   0
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   2  'Horizontal Line
         Height          =   15
         Left            =   -720
         Top             =   350
         Width           =   5000
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   110
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7140
      TabIndex        =   62
      Top             =   4800
      Width           =   3900
      Begin VB.Shape Shape6 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label OrderForWhere 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dine In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   70
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Destination 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   69
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Destination 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3050
         TabIndex        =   68
         Top             =   120
         Width           =   735
      End
      Begin VB.Label OrderForWhere 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Go Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   67
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label OrderForWhere 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   66
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Destination 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3050
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label OrderForWhere 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Thru"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Destination 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   63
         Top             =   360
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame Title 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2900
      Left            =   3475
      TabIndex        =   49
      Top             =   7360
      Width           =   3615
      Begin VB.Shape Shape7 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   0
         Top             =   350
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions and Averages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   495
         Left            =   120
         TabIndex        =   87
         Top             =   110
         Width           =   3255
      End
      Begin VB.Label TableTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2400
         TabIndex        =   61
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Average Table Time"
         ForeColor       =   &H00706B66&
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   2040
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label WinTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2400
         TabIndex        =   59
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Average Time @ Window"
         ForeColor       =   &H00706B66&
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   2280
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label G_Data 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   57
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Average"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   56
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Average"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   55
         Top             =   960
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Customers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   1320
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Transactions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   53
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label G_Data 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   52
         Top             =   600
         Width           =   855
      End
      Begin VB.Label G_Data 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   51
         Top             =   960
         Width           =   855
      End
      Begin VB.Label G_Data 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   50
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00492C1C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   15255
      Begin BTNENHLib4.BtnEnh SSCommand7 
         Height          =   525
         Left            =   5280
         TabIndex        =   43
         Top             =   45
         Width           =   1995
         _Version        =   589828
         _ExtentX        =   3519
         _ExtentY        =   926
         _StockProps     =   66
         Caption         =   "Hourly Sales"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":06F2
         textLT          =   "Red_Book.frx":076A
         textCT          =   "Red_Book.frx":0782
         textRT          =   "Red_Book.frx":079A
         textLM          =   "Red_Book.frx":07B2
         textRM          =   "Red_Book.frx":07CA
         textLB          =   "Red_Book.frx":07E2
         textCB          =   "Red_Book.frx":07FA
         textRB          =   "Red_Book.frx":0812
         colorBack       =   "Red_Book.frx":082A
         colorIntern     =   "Red_Book.frx":0854
         colorMO         =   "Red_Book.frx":087E
         colorFocus      =   "Red_Book.frx":08A8
         colorDisabled   =   "Red_Book.frx":08D2
         colorPressed    =   "Red_Book.frx":08FC
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
         Office2003ColorNorth=   5851449
         Office2003ColorSouth=   5851449
         Office2003ColorBorder=   13399296
         Office2003ColorNorthMouseOver=   10322534
         Office2003ColorSouthMouseOver=   10322534
         Office2003ColorBorderMouseOver=   16760704
         Office2003ColorNorthPressed=   10322534
         Office2003ColorSouthPressed=   10322534
         Office2003ColorBorderPressed=   16760704
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   525
         Left            =   3240
         TabIndex        =   44
         Top             =   45
         Width           =   1995
         _Version        =   589828
         _ExtentX        =   3519
         _ExtentY        =   926
         _StockProps     =   66
         Caption         =   "Daily Waste"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0926
         textLT          =   "Red_Book.frx":099C
         textCT          =   "Red_Book.frx":09B4
         textRT          =   "Red_Book.frx":09CC
         textLM          =   "Red_Book.frx":09E4
         textRM          =   "Red_Book.frx":09FC
         textLB          =   "Red_Book.frx":0A14
         textCB          =   "Red_Book.frx":0A2C
         textRB          =   "Red_Book.frx":0A44
         colorBack       =   "Red_Book.frx":0A5C
         colorIntern     =   "Red_Book.frx":0A86
         colorMO         =   "Red_Book.frx":0AB0
         colorFocus      =   "Red_Book.frx":0ADA
         colorDisabled   =   "Red_Book.frx":0B04
         colorPressed    =   "Red_Book.frx":0B2E
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
         Office2003ColorNorth=   5851449
         Office2003ColorSouth=   5851449
         Office2003ColorBorder=   13399296
         Office2003ColorNorthMouseOver=   10322534
         Office2003ColorSouthMouseOver=   10322534
         Office2003ColorBorderMouseOver=   16760704
         Office2003ColorNorthPressed=   10322534
         Office2003ColorSouthPressed=   10322534
         Office2003ColorBorderPressed=   16760704
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand6 
         Height          =   525
         Left            =   7320
         TabIndex        =   45
         Top             =   45
         Width           =   1995
         _Version        =   589828
         _ExtentX        =   3519
         _ExtentY        =   926
         _StockProps     =   66
         Caption         =   "Print Sale Report"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0B58
         textLT          =   "Red_Book.frx":0BDA
         textCT          =   "Red_Book.frx":0BF2
         textRT          =   "Red_Book.frx":0C0A
         textLM          =   "Red_Book.frx":0C22
         textRM          =   "Red_Book.frx":0C3A
         textLB          =   "Red_Book.frx":0C52
         textCB          =   "Red_Book.frx":0C6A
         textRB          =   "Red_Book.frx":0C82
         colorBack       =   "Red_Book.frx":0C9A
         colorIntern     =   "Red_Book.frx":0CC4
         colorMO         =   "Red_Book.frx":0CEE
         colorFocus      =   "Red_Book.frx":0D18
         colorDisabled   =   "Red_Book.frx":0D42
         colorPressed    =   "Red_Book.frx":0D6C
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
         Office2003ColorNorth=   5851449
         Office2003ColorSouth=   5851449
         Office2003ColorBorder=   13399296
         Office2003ColorNorthMouseOver=   10322534
         Office2003ColorSouthMouseOver=   10322534
         Office2003ColorBorderMouseOver=   16760704
         Office2003ColorNorthPressed=   10322534
         Office2003ColorSouthPressed=   10322534
         Office2003ColorBorderPressed=   16760704
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand8 
         Height          =   525
         Left            =   9360
         TabIndex        =   46
         Top             =   45
         Width           =   1995
         _Version        =   589828
         _ExtentX        =   3519
         _ExtentY        =   926
         _StockProps     =   66
         Caption         =   "Print Meal Period"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0D96
         textLT          =   "Red_Book.frx":0E18
         textCT          =   "Red_Book.frx":0E30
         textRT          =   "Red_Book.frx":0E48
         textLM          =   "Red_Book.frx":0E60
         textRM          =   "Red_Book.frx":0E78
         textLB          =   "Red_Book.frx":0E90
         textCB          =   "Red_Book.frx":0EA8
         textRB          =   "Red_Book.frx":0EC0
         colorBack       =   "Red_Book.frx":0ED8
         colorIntern     =   "Red_Book.frx":0F02
         colorMO         =   "Red_Book.frx":0F2C
         colorFocus      =   "Red_Book.frx":0F56
         colorDisabled   =   "Red_Book.frx":0F80
         colorPressed    =   "Red_Book.frx":0FAA
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
         Office2003ColorNorth=   5851449
         Office2003ColorSouth=   5851449
         Office2003ColorBorder=   13399296
         Office2003ColorNorthMouseOver=   10322534
         Office2003ColorSouthMouseOver=   10322534
         Office2003ColorBorderMouseOver=   16760704
         Office2003ColorNorthPressed=   10322534
         Office2003ColorSouthPressed=   10322534
         Office2003ColorBorderPressed=   16760704
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand9 
         Height          =   525
         Left            =   11400
         TabIndex        =   47
         Top             =   50
         Width           =   1995
         _Version        =   589828
         _ExtentX        =   3519
         _ExtentY        =   926
         _StockProps     =   66
         Caption         =   "Print All"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":0FD4
         textLT          =   "Red_Book.frx":1046
         textCT          =   "Red_Book.frx":105E
         textRT          =   "Red_Book.frx":1076
         textLM          =   "Red_Book.frx":108E
         textRM          =   "Red_Book.frx":10A6
         textLB          =   "Red_Book.frx":10BE
         textCB          =   "Red_Book.frx":10D6
         textRB          =   "Red_Book.frx":10EE
         colorBack       =   "Red_Book.frx":1106
         colorIntern     =   "Red_Book.frx":1130
         colorMO         =   "Red_Book.frx":115A
         colorFocus      =   "Red_Book.frx":1184
         colorDisabled   =   "Red_Book.frx":11AE
         colorPressed    =   "Red_Book.frx":11D8
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
         Office2003ColorNorth=   5851449
         Office2003ColorSouth=   5851449
         Office2003ColorBorder=   13399296
         Office2003ColorNorthMouseOver=   10322534
         Office2003ColorSouthMouseOver=   10322534
         Office2003ColorBorderMouseOver=   16760704
         Office2003ColorNorthPressed=   10322534
         Office2003ColorSouthPressed=   10322534
         Office2003ColorBorderPressed=   16760704
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand3 
         Height          =   525
         Left            =   13440
         TabIndex        =   48
         Top             =   45
         Width           =   1655
         _Version        =   589828
         _ExtentX        =   2910
         _ExtentY        =   926
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Red_Book.frx":1202
         textLT          =   "Red_Book.frx":126A
         textCT          =   "Red_Book.frx":1282
         textRT          =   "Red_Book.frx":129A
         textLM          =   "Red_Book.frx":12B2
         textRM          =   "Red_Book.frx":12CA
         textLB          =   "Red_Book.frx":12E2
         textCB          =   "Red_Book.frx":12FA
         textRB          =   "Red_Book.frx":1312
         colorBack       =   "Red_Book.frx":132A
         colorIntern     =   "Red_Book.frx":1354
         colorMO         =   "Red_Book.frx":137E
         colorFocus      =   "Red_Book.frx":13A8
         colorDisabled   =   "Red_Book.frx":13D2
         colorPressed    =   "Red_Book.frx":13FC
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
         Office2003ColorSouth=   102
         Office2003ColorBorder=   255
         Office2003ColorNorthMouseOver=   128
         Office2003ColorSouthMouseOver=   128
         Office2003ColorBorderMouseOver=   128
         Office2003ColorNorthPressed=   128
         Office2003ColorSouthPressed=   128
         Office2003ColorBorderPressed=   128
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Live Sales Dashboard"
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
         Height          =   255
         Left            =   200
         TabIndex        =   90
         Top             =   175
         Width           =   3015
      End
   End
   Begin VB.Frame Frame14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1045
      Left            =   50
      TabIndex        =   39
      Top             =   5580
      Width           =   3375
      Begin VB.Shape Shape17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   1030
         Width           =   3175
      End
      Begin VB.Label CheckSales 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label CheckSales 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Closed Check Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   40
         Top             =   175
         Width           =   3375
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   50
      TabIndex        =   34
      Top             =   6600
      Width           =   3375
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gratuity Collected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   38
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   37
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Non Sales Revenue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1045
      Left            =   50
      TabIndex        =   31
      Top             =   4535
      Width           =   3375
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Open Check Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   375
         Left            =   0
         TabIndex        =   91
         Top             =   175
         Width           =   3375
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label CheckSales 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label CheckSales 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Open Check Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   0
         Left            =   -9840
         TabIndex        =   32
         Top             =   4680
         Width           =   3375
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   3475
      TabIndex        =   15
      Top             =   720
      Width           =   3615
      Begin VB.Frame Frame21 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1045
         Left            =   120
         TabIndex        =   84
         Top             =   2620
         Width           =   3375
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "WTD Discounts"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00706B66&
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   93
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label sales 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   0
            TabIndex        =   92
            Top             =   480
            Width           =   1695
         End
         Begin VB.Shape Shape13 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00E0E0E0&
            Height          =   735
            Index           =   4
            Left            =   1680
            Top             =   120
            Width           =   15
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "WTD Voids"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00706B66&
            Height          =   255
            Index           =   15
            Left            =   1680
            TabIndex        =   86
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label sales 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   13
            Left            =   1680
            TabIndex        =   85
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1045
         Left            =   120
         TabIndex        =   72
         Top             =   460
         Width           =   3375
         Begin VB.Shape Shape13 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00E0E0E0&
            Height          =   15
            Index           =   2
            Left            =   0
            Top             =   960
            Width           =   3375
         End
         Begin VB.Shape Shape13 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00E0E0E0&
            Height          =   15
            Index           =   1
            Left            =   0
            Top             =   1710
            Width           =   3375
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "WTD Net Sales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00706B66&
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   74
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label sales 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8300&
            Height          =   375
            Index           =   10
            Left            =   0
            TabIndex        =   73
            Top             =   480
            Width           =   3375
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1045
         Left            =   120
         TabIndex        =   71
         Top             =   1540
         Width           =   3375
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "WTD Net (inc. Tax)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00706B66&
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   95
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label sales 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H003A2B2B&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8300&
            Height          =   375
            Index           =   11
            Left            =   0
            TabIndex        =   94
            Top             =   480
            Width           =   3360
         End
         Begin VB.Shape Shape13 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00E0E0E0&
            Height          =   15
            Index           =   3
            Left            =   0
            Top             =   960
            Width           =   3615
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "This Week (Week to date)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B1618&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   3015
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00FD8A24&
         Height          =   15
         Left            =   0
         Top             =   375
         Width           =   4215
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2465
      Left            =   50
      TabIndex        =   14
      Top             =   7800
      Width           =   3375
      Begin VB.Shape Shape4 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   375
         Width           =   3135
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8300&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   30
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sales W/ Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1705
         Width           =   1935
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquor Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   27
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   25
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label sales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8300&
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   145
         Width           =   1695
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   1935
         Width           =   3135
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1045
      Left            =   50
      TabIndex        =   13
      Top             =   3490
      Width           =   3375
      Begin VB.Shape Shape15 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   960
         Width           =   3137
      End
      Begin VB.Label sales 
         Alignment       =   2  'Center
         BackColor       =   &H003A00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   19
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Voided Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   175
         Width           =   3375
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1100
      Left            =   50
      TabIndex        =   12
      Top             =   2430
      Width           =   3375
      Begin VB.Shape Shape14 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label sales 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Discounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   16
         Top             =   175
         Width           =   3375
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1FFF0&
      Height          =   1770
      Left            =   50
      TabIndex        =   9
      Top             =   720
      Width           =   3375
      Begin VB.Shape Shape18 
         BorderColor     =   &H00E0E0E0&
         Height          =   15
         Left            =   120
         Top             =   1700
         Width           =   3135
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   15
         Index           =   0
         Left            =   0
         Top             =   1710
         Width           =   3375
      End
      Begin VB.Label sales 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8300&
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   775
         Width           =   3375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00706B66&
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      Caption         =   "Payment Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   15120
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   4095
      Begin VSFlex8DAOCtl.VSFlexGrid PaymentGrid 
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
         _cx             =   1973688976
         _cy             =   1973684531
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   0
         BackColorFixed  =   8454016
         ForeColorFixed  =   16711680
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   100
         Cols            =   4
         FixedRows       =   1
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Serving Times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   15120
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   3135
      Begin VSFlex8DAOCtl.VSFlexGrid AvgSerTime 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2895
         _cx             =   1973687282
         _cy             =   1973684954
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   0
         BackColorFixed  =   8454143
         ForeColorFixed  =   16711680
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   200
         Cols            =   4
         FixedRows       =   1
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
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1020
      Left            =   5400
      Top             =   7920
      Width           =   8295
   End
End
Attribute VB_Name = "Red_Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum HrlSaleData
       DayByRegister
       WTDByRegister
       ALLDay
       ALLWeek
End Enum
Private Type RestLabor
        JobTitle As String
        JobRate  As Currency
        JobHours As Single
End Type
Private Jobs() As RestLabor

Private Cd As ChartDirector.API
'Cd.setLicenseCode "RDST-247P-CD3B-ZSV7-96F8-4DDA"
Private Total_depts             As Integer
Public NextTab                  As Integer
Private MyTimer                 As Long
Private Sub BtnEnh26_Click()
HourlyData.Visible = False
HourlyData.Left = -5000
End Sub
Private Sub BtnEnh27_Click()
Erase Hourly, MealPeriod
ReDim Hourly(1, 24, 3)
ReDim MealPeriod(1, 5, 4)
HourlySales Hourly(), MealPeriod(), MyRegisters(0).regName, 0
'// 0 is hourly sales for today
Print_Hourly_Sales True, Format(Date, "MM/DD/YYYY")
Erase Hourly, MealPeriod
BtnEnh26_Click
'// clear the selection box from the screen
End Sub
Private Sub BtnEnh28_Click()
Erase Hourly, MealPeriod
ReDim Hourly(1, 24, 3)
ReDim MealPeriod(1, 5, 4)
HourlySales Hourly(), MealPeriod(), MyRegisters(0).regName, -3
'// -3 is week to date hourly sales
Print_Hourly_Sales True, Format(Date, "MM/DD/YYYY")
Erase Hourly, MealPeriod
BtnEnh26_Click
End Sub
Private Sub Exit_Click(Index As Integer)
Call MemoryCheck(Red_Book)
End Sub
Private Sub Load_Meal_Periods()
On Error GoTo ERH
Dim R       As Integer
Dim Data()
Dim Labels()
Dim Colors()
Dim Td      As Integer
Colors = Array(&H66AAEE, &HEEBB22, &HBBBBBB, &H8844FF, &HDD2222, &H9900)
Td = -1
For R = 1 To 5
  If Nz(MealPeriod(0, R, 1)) > 0 Then
   Td = Td + 1
  End If
Next R
ReDim Data(Td)
ReDim Labels(Td)
ReDim Colors(Td)
Td = 0
For R = 1 To 5
   If Nz(MealPeriod(0, R, 1)) > 0 Then
     Labels(Td) = MealPeriod(0, R, 0) ' MealP(R)
     Data(Td) = Format(MealPeriod(0, R, 1), "0.00")
     Td = Td + 1
   End If
 Next R
If Td > 0 Then
 createChart MealPeriods, Data(), Labels(), Colors(), "Meal Period Sales"
End If
Exit Sub
ERH:
Err.Clear
End Sub
Public Sub createChart(viewer As Object, Data() As Variant, Labels() As Variant, Colors() As Variant, ChartTitle As String)
    On Error GoTo ERH
    Dim C As PieChart
    Dim T As ChartDirector.TextBox
    If ChartTitle = "ORDERDEST" Then
        Set C = Cd.PieChart(350 * WidScale1024, 350 * TopScale1024)
        Call C.setDonutSize(140 * WidScale1024, 100 * TopScale1024, 50 * WidScale1024, 25 * WidScale1024)
        Set T = C.setLabelStyle("arialbd.ttf", 10, Cd.SameAsMainColor)
        Call T.setBackground(Cd.Transparent)
        'Set the border color of the sector the same color as the fill color. Set the
        'line color of the join line to black (0x0)
        Call C.setLineColor(Cd.SameAsMainColor, &HFFFFFF)
        Call C.setLabelFormat("{label}<*br*>${value}")
     ElseIf ChartTitle = "Sales vs. Labor" Then
        Set C = Cd.PieChart(250 * WidScale1024, 250 * TopScale1024)
        Call C.setDonutSize(110 * WidScale1024, 100 * TopScale1024, 45 * WidScale1024, 22 * WidScale1024)
        Call C.setLabelFormat("{label}<*br*>{percent}%")
        Call C.setLineColor(Cd.SameAsMainColor, &HFFFFFF)
        Set T = C.setLabelStyle("arialbd.ttf", 10, &H0)
        Call T.setBackground(Cd.Transparent)
     Else
        Set C = Cd.PieChart(350 * WidScale1024, 350 * TopScale1024)
        Call C.setDonutSize(140 * WidScale1024, 100 * TopScale1024, 50 * WidScale1024, 25 * WidScale1024)
        Call C.setLabelFormat("{label}<*br*>${value}")
        Call C.setLineColor(Cd.SameAsMainColor, &HFFFFFF)
        Set T = C.setLabelStyle("arialbd.ttf", 10, Cd.SameAsMainColor)
        Call T.setBackground(Cd.Transparent)
    End If
    Call C.setExplode(0)
    Call C.setLabelLayout(Cd.SideLayout, 2)
    Call C.SetData(Data, Labels)
    Call C.setColors2(Cd.DataColor, Array(&H3A5AFF, &H3AA5FF, &H3AE6FF, &H8844FF, &HDD2222, &H9900))
    Set viewer.Picture = C.makePicture()
        'include tool tip for the chart
    viewer.ImageMap = C.getHTMLImageMap("clickable", "", _
        "title='{label}: ${value} ({percent}%)'")
Exit Sub
ERH:
 Error_Log Err.Number, Err.Description, "red book create chart", "None"
End Sub
Private Sub CalLabor(BusinessDay As String)
On Error GoTo ErrorHandler
Dim PRTb        As ADODB.Recordset
Dim JCS         As Integer
Dim EDb         As ADODB.Connection
Dim Id          As ADODB.Recordset
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim dw          As Integer
Dim Cw          As Integer
Dim WrkTime     As Single
Dim PRate       As Currency
Dim Total_Time  As Single
Dim Ttl_P       As Single
Dim Ttl_L       As Single
Dim TWRKTIME    As Single
Dim EmpTwt      As Single
Dim Pl          As Integer
Dim LabPer      As Single

TWRKTIME = 0
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
If BusinessDay = "" Then Exit Sub
OpenDb SvrPath.LocalTerminalPath & "Employee.mdb", EDb
  Sql = "Select * from [General Data] Where ISNULL(TRY_CONVERT(int,[Terminated]),0) = 0 order by [Last Name] asc"
ConnectADORS Sql, EDb, Id
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
Do While Not Id.EOF
    SearchName = Id(0) & " " & Id(2)
    Pen = 0: EmpTwt = 0
    Sql = "Select * From [Job Codes] Where [First Name] = '" & Id(0) & "'"
    Sql = Sql & " and [Last Name] = '" & Id(2) & "'"
   ConnectADORS Sql, EDb, PRTb
    ReDim PayRate(RecordsCount(PRTb))
    JCS = 0
    Do While Not PRTb.EOF
         PayRate(JCS).JobCodes = PRTb![Job Code] & ""
         PayRate(JCS).JobRate = Nz(PRTb![rate])
         PayRate(JCS).JobHourly = Nz(PRTb![Hourly])
         JCS = JCS + 1
      PRTb.MoveNext
    Loop
SetNothing PRTb
    Sql = "Select * From [Employee Records] Where [Employee Name] = '" & SearchName & "'"
    Sql = Sql & " AND TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BusinessDay) & " " & GetStartTime & "'"
    Sql = Sql & " and '" & SetEndingDate(CDate(BusinessDay)) & " " & SetEndingTime & "'"
    
    ConnectADORS Sql, Db, Tb
       'If Not Tb.BOF Then Tb.MoveFirst
       Do While Not Tb.EOF
         Call CalculateTimeRecord(CStr(Tb(2)) & "", CStr(Tb(7)) & "", CStr(Tb(3)) & "", CStr(Tb![Break Out] & ""), CStr(Tb![Break In] & ""), CStr(Tb(6)) & "", Total_Time)
         'HW = Hoursworked
          For Ck = 0 To JCS - 1
            If Trim(Tb(1)) = Trim(PayRate(Ck).JobCodes) Then
             PRate = PayRate(Ck).JobRate 'PayRate(Ck).JobRate
             Exit For
            End If
          Next Ck
       WrkTime = ConvertTimeToMinutes(Total_Time)
       CalDayLabor Tb(1) & "", PRate, WrkTime, Abs(PayRate(Ck).JobHourly), 1, OvertimeHrs(EmpTwt, WrkTime)
       TWRKTIME = TWRKTIME + Nz(WrkTime)
       EmpTwt = EmpTwt + Nz(WrkTime)
       WrkTime = 0
       Tb.MoveNext
     Loop
  Id.MoveNext
Loop
DbClose Db, Tb
DbClose EDb, Id
Dim TotalSales As Currency
If Check_Option(91) = 1 Then
   TotalSales = TotalNetSales + NetDiscounts
Else
   TotalSales = TotalNetSales
End If

For Pl = 0 To Total_depts - 1
  If Trim(Jobs(Pl).JobTitle) > "" And Jobs(Pl).JobRate > 0 And TotalSales > 0 Then
    LabPer = (Jobs(Pl).JobRate / (TotalSales + OpenCk)) * 100
  End If
  Ttl_L = Ttl_L + Jobs(Pl).JobRate
Next Pl
If Nz(Ttl_L) > 0 And TotalSales > 0 Then Ttl_P = Ttl_L / (TotalSales + OpenCk) * 100
If Nz(Ttl_L & "") <> 0 Then
 Dim Data()
 Dim Labels()
    Dim Colors()
    Colors = Array(&H66AAEE, &HEEBB22, &HBBBBBB, &H8844FF, &HDD2222, &H9900)
  With LaborData
    Data = Array(Ttl_L, ((TotalSales + OpenCk) - Ttl_L))
    Labels = Array("Labor", "Net Sales")
  End With
  createChart LD, Data(), Labels(), Colors(), "Sales vs. Labor"
End If
Exit Sub
ErrorHandler:
If Err.Number = 9 Then
   Err.Clear
   Exit Sub
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Calculate Labor RedBook")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub CalDayLabor(WorkDept As String, Dept_Pay As Currency, WorkTime As Single, HowPaid As Integer, DaysWorked As Integer, OverTime As Single)
    On Error GoTo ERH
    Dim Cld As Integer
    For Cld = 0 To Total_depts
       If Trim(WorkDept) = Trim(Jobs(Cld).JobTitle) Then
        If HowPaid = 1 Then
            Jobs(Cld).JobRate = Jobs(Cld).JobRate + (Dept_Pay * WorkTime) + CalculateOT(Val(OverTime), CCur(Dept_Pay))
            Jobs(Cld).JobHours = Jobs(Cld).JobHours + WorkTime + CalculateOT(Val(OverTime), CCur(Dept_Pay))
        End If
        Exit For
       End If
    Next Cld
ERH:
End Sub
Private Sub BtnEnh3_Click()
MyTimer = 0
CurrentDayWasteReport
End Sub
Private Sub ETime_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
End Sub
Private Sub ChartViewer1_ClickHotSpot(hotSpot As Collection, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Cd As New ChartDirector.API

    ' The data for the bar chart
    Dim Data()
    Data = Array(85, 156, 179.5, 211, 123)

    ' The labels for the bar chart
    Dim Labels()
    Labels = Array("Mon", "Tue", "Wed", "Thu", "Fri")

    ' Create a XYChart object of size 250 x 250 pixels
    Dim C As XYChart
    Set C = Cd.XYChart(250, 250)

    ' Set the plotarea at (30, 20) and of size 200 x 200 pixels
    Call C.setPlotArea(30, 20, 200, 200)

    ' Add a bar chart layer using the given data
    Call C.addBarLayer(Data)

    ' Set the labels on the x axis.
    Call C.xAxis().setLabels(Labels)

    ' Output the chart
    Set viewer.Picture = C.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = C.getHTMLImageMap("clickable", "", "title='{xLabel}: {value} GBytes'")
End Sub

Private Sub Form_Load()
ResizeScreen1024 Me
CenterForm Me
Set Cd = New ChartDirector.API
Cd.setLicenseCode "RDST-247P-CD3B-ZSV7-96F8-4DDA"
On Error GoTo ErrorHandler
Dim R    As Integer
Dim Pa   As Integer
Dim TC   As Integer
Dim HSal As Currency
MyTimer = 0
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
 Sql = "Select * from [Job Titles] order by [Job Title] asc"
ConnectADORS Sql, Db, Tb
  Total_depts = RecordsCount(Tb)
  ReDim Jobs(Total_depts)
   DN = -1
   Do While Not Tb.EOF
      DN = DN + 1
      Jobs(DN).JobTitle = Tb.fields("Job Title").Value & ""   ' Dept Name
      Tb.MoveNext
   Loop
DbClose Db, Tb
SalesReading True


HourlyGrid.ColAlignment(1) = flexAlignRightCenter
HourlyGrid.ColAlignment(2) = flexAlignRightCenter
HourlyGrid.ColAlignment(3) = flexAlignRightCenter
PaymentGrid.ColAlignment(1) = 1
PaymentGrid.ColAlignment(2) = 1
PaymentGrid.ColAlignment(3) = 3
PaymentGrid.ColWidth(0) = 1100
PaymentGrid.ColWidth(1) = 1100
PaymentGrid.ColWidth(2) = 800
PaymentGrid.ColWidth(3) = 800
With AvgSerTime
   .Col = 3
   .ColAlignment(0) = 1
   .ColAlignment(1) = 3
   .ColAlignment(2) = 3
   .ColWidth(0) = 2000
   .ColWidth(1) = 850
   .ColWidth(2) = 850
End With
Dim SrTime() As ServerTimeData
Dim lg As Integer

lg = TimeByServer(SrTime(), "")
With AvgSerTime
.TextMatrix(0, 0) = "Name"
.TextMatrix(0, 1) = "Mins"
.TextMatrix(0, 2) = "Trans"
For lg = 0 To lg
  .TextMatrix(lg + 1, 0) = SrTime(lg).Name
  .TextMatrix(lg + 1, 1) = SrTime(lg).ETime
  .TextMatrix(lg + 1, 2) = SrTime(lg).ERcd
Next lg
End With
If System_config = ExpressService Then
   WinTime.Caption = GetDriveThruTime
   If Val(WinTime.Caption) > 0 Then
    Label8.Visible = True
    WinTime.Visible = True
   End If
Else
    TableTime.Caption = avgTime
    If Val(TableTime.Caption) > 0 Then
       Label4.Visible = True
       TableTime.Visible = True
    End If
End If
   Hourly_Sales_Data ""
   With HourlyGrid
        .ColAlignment(0) = 1
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .FontSize = 9
        .TextMatrix(0, 0) = "Hours"
        .TextMatrix(0, 1) = "Net Sales"
        .TextMatrix(0, 2) = "# Trans."
        .TextMatrix(0, 3) = "# Cust."
        For HRS = 6 To 23
            If Val(Hourly(0, HRS, 1)) <> 0 Then
                 R = R + 1
                .TextMatrix(R, 0) = Format(Hourly(0, HRS, 0), "H:MM ampm")
                .TextMatrix(R, 1) = Format(Hourly(0, HRS, 1), "0.00")
                .TextMatrix(R, 2) = Format(Hourly(0, HRS, 2), "##")
                .TextMatrix(R, 3) = Format(Hourly(0, HRS, 3), "##")
            End If
        Next HRS
        For HRS = 0 To 5
            If Val(Hourly(0, HRS, 1)) <> 0 Then
                 R = R + 1
                .TextMatrix(R, 0) = Format(Hourly(0, HRS, 0), "H:MM ampm")
                .TextMatrix(R, 1) = Format(Hourly(0, HRS, 1), "0.00")
                .TextMatrix(R, 2) = Format(Hourly(0, HRS, 2), "##")
                .TextMatrix(R, 3) = Format(Hourly(0, HRS, 3), "##")
            End If
        Next HRS
  End With
Dim Payments As FinancialItems
Set Payments = New FinancialItems
SortCreditCards Payments, 0, "", 0, "", ""
With PaymentGrid
   .Row = 0
   .TextMatrix(0, 1) = "Payment"
   .TextMatrix(0, 1) = "Amount"
   .TextMatrix(0, 2) = "Tips"
   .TextMatrix(0, 3) = "Count"
    For R = 1 To Payments.count
     .TextMatrix(R, 0) = Payments.Item(R).Tender
     .TextMatrix(R, 1) = Payments.Item(R).Amount
     .TextMatrix(R, 2) = Payments.Item(R).TipAmount
     .TextMatrix(R, 3) = Payments.Item(R).ItemQty
   Next R
End With
Set Payments = Nothing
Sleep 0
Load_Meal_Periods
Dim Td As Integer
Dim DeptReports() As DeptSale
Dim Data()
Dim Labels()
Dim Colors()
 Colors = Array(&H3A5AFF, &H2ADE1, &HA808EC, &H1BF7F0, &HE66291, &H9EA6CD)
Td = CurrentDeptSales(DeptReports())
On Error Resume Next
On Error GoTo ErrorHandler

If (Td - 1) > -1 Then
  ReDim Data(Td - 1)
  ReDim Labels(Td - 1)
  ReDim Colors(Td - 1)
  For Printall = 0 To Td - 1
      Labels(Printall) = DeptReports(Printall).Department
      Data(Printall) = Format(DeptReports(Printall).DollarsSold, "0.00")
  Next Printall
  createChart LaborBreakdown, Data(), Labels(), Colors(), "Department Sales"
  'createChart LaborBreakdown, Data(), Labels(), Colors(), "Department Sales"
End If
Td = -1
For R = 0 To 3
   If Val(Destination(R).Caption) > 0 Then
    Td = Td + 1
   End If
Next R
If (Td - 1) > -1 Then
 ReDim Data(Td)
 ReDim Labels(Td)
 ReDim Colors(Td)
 Td = 0
 For R = 0 To 3
   If Val(Destination(R).Caption) > 0 Then
      Labels(Td) = OrderForWhere(R).Caption
      Data(Td) = Destination(R).Caption
      Td = Td + 1
    End If
 Next R
 createChart OrderDest, Data(), Labels(), Colors(), "ORDERDEST"
End If
CalLabor GetBusinessDay
Timer1.Enabled = True
Me.Refresh
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Red Book Load Form ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Public Sub SalesReading(ShowallSales As Boolean)
On Error GoTo ErrorHandler
Dim Display As Integer
Dim Db      As ADODB.Connection
Dim OP      As ADODB.Recordset

Sales_Breakdown ShowallSales, 0

Title.Caption = "Total Sales"

For Display = 0 To 13
  sales(Display).Caption = Format(Reg_Sales(0, Display), "0.00")
Next Display
Sleep 0

For Display = 0 To 3
 CheckSales(Display).Visible = False
Next Display

If RegCounts.Transaction > 0 And TotalNetSales > 0 Then
 RegCounts.TransAvg = Format(TotalNetSales / RegCounts.Transaction, "0.00")
End If
If RegCounts.Customer > 0 And TotalNetSales > 0 Then
 RegCounts.CustAvg = Format(TotalNetSales / RegCounts.Customer, "0.00")
End If

G_Data(0).Caption = RegCounts.Transaction
G_Data(1).Caption = Format(RegCounts.TransAvg, "0.00")
G_Data(2).Caption = RegCounts.Customer
G_Data(3).Caption = Format(RegCounts.CustAvg, "0.00")

For Display = 0 To 3
    Destination(Display).Caption = Format(RegisterSales(Display).NetSaleAmount - RegisterSales(Display).NetDiscAmount, "0.00")
Next Display
'Delivery Orders
If RegisterSales(0).NetSaleAmount > 0 Then
    If Destination(2).Caption > 0 Then
      Destination(2).Visible = True
      OrderForWhere(2).Visible = True
    End If
    If RegisterSales(3).NetSaleAmount + RegisterSales(4).NetSaleAmount > 0 Then
      Destination(3).Caption = Format(RegisterSales(3).NetSaleAmount + RegisterSales(4).NetSaleAmount, "0.00")
      Destination(3).Visible = True      'Pickup is 3 Delivery = 4
      OrderForWhere(3).Visible = True
    End If
End If
OpenCk = 0: OSTaxes = 0

  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  ConnectADORS "Select * from [Open Checks] Where [Close Out Day] = 0", Db, OP
    Do While Not OP.EOF
       OCNetDiscount = OP.fields("Gross Discounts").Value - OP.fields("Voided Discounts").Value
       NetSale = (OP.fields("Gross Sales").Value - OP.fields("Voided Sales").Value) - OCNetDiscount
       OpenCk = Nz(OpenCk) + Nz(NetSale)
       NST = OP.fields("Gross Sales Tax").Value - OP.fields("Voided Sales Tax").Value
       NSPT = OP.fields("Gross Special Tax").Value - OP.fields("Voided Special Tax").Value
       NLT = OP.fields("Gross Liquor Tax").Value ' - OP(14) voided liquor tax obsolete
       OSTaxes = OSTaxes + (NST + (NSPT + NLT))
       OP.MoveNext
    Loop
    DbClose Db, OP
    For C = 0 To 3
     CheckSales(C).Visible = True
    Next C
    OpenCk = OpenCk + GetPhoneOrderSales
    CheckSales(1).Caption = Format(OpenCk, "0.00")
    CheckSales(3).Caption = Format(sales(5).Caption, "0.00")
   ' sales(0).Caption = Format(sales(0).Caption + OpenCk, "0.00")
   ' sales(5).Caption = Format(sales(5).Caption + OpenCk, "0.00")  'Net sales + OpenCk
    'sales(1).Caption = Format(sales(1).Caption + OSTaxes, "0.00")
    'sales(6).Caption = Format((sales(6).Caption + (OpenCk + OSTaxes)), "0.00") ' net with tax
    'adds open check sales to Net sales
    If OpenCk > 0 Then
        CheckSales(1).Visible = True
        CheckSales(1).Caption = Format(OpenCk, "0.00")
        CheckSales(0).Visible = True
        CheckSales(0).Caption = "Open Sales Orders"
        sales(0).Caption = Format(sales(0).Caption + OpenCk, "0.00")
        sales(1).Caption = Format(sales(1).Caption + OSTaxes, "0.00")
        sales(5).Caption = Format(sales(5).Caption + OpenCk, "0.00")   'Net sales + OpenCk
        sales(6).Caption = Format((sales(6).Caption + (OpenCk + OSTaxes)), "0.00") ' net with tax
    End If

Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Sales Reading ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub KVSSVRRPT_Click()

End Sub

Private Sub SPEEDOFSVR_Click()
End Sub

Private Sub Label10_Click()

End Sub

Private Sub FullService_Click()

End Sub

Private Sub SSCommand3_Click()
Set Cd = Nothing
Call MemoryCheck(Red_Book)
End Sub
Private Sub SSCommand4_Click()
End Sub
Private Sub SSCommand5_Click()
End Sub
Private Sub SSCommand6_Click()
MyTimer = 0

Sales_Breakdown True, 0
Sales_Report -1, 0, Format(Date, "MM/DD/YYYY")
End Sub
Private Sub SSCommand7_Click()
CenterObject Me, HourlyData
HourlyData.Visible = True
HourlyData.ZOrder 0
End Sub
Private Sub SSCommand8_Click()
MyTimer = 0
Print_Meal_Period -1
End Sub
Private Sub CurrentDayWasteReport()
On Error GoTo ErrorHandler
Dim SalesData() As Variant
WasteReport 0, SalesData()
Dim PrinterPort As String
If UBound(SalesData, 2) = 0 Then
 See_Msg "No Waste Found", 5, ""
 Exit Sub
End If
Dim TQty As Long, TSold As Currency
If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    CenterText = EC(8)
    LeftText = EC(9)
    'RightText = Chr(27) + Chr(97) + Chr(2)
    Dim PF As Integer
    PF = FreeFile
    PrinterPort = "Waste Sales"
OneMoreTime:
    Open PrinterPort For Output Shared As #PF
End If
'Val(sales(5).Caption))  '=Net Sales
If Check_Option(96) = 0 Then
    Header PF
    Print #PF, EC(1); EC(3); CenterText; "Daily Waste Report"; EC(4); EC(7)
    Print #PF, LeftText; EC(7)
    Print #PF, ; EC(7)
    Print #PF, Format(Time, "HH:MM AmPm");
    Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
  Else
    SetPage ReportView.View, "Daily Waste Report"
    PrintCLine "Daily Waste Report", -1, 0
  End If
  If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Name                     Amt    Percent"; EC(7)
    Print #PF, String(40, 45); EC(7)
  Else
    DrawLine
    PrintLine "Name", 3200: PrintLine "Amount", 4800: PrintLine "Percent", -1
    DrawLine
  End If
Dim Cl As Integer
Dim WP As Single
Dim TW As Currency
TW = 0
For Cl = 0 To UBound(SalesData, 2) - 1
  If Val(SalesData(1, Cl)) > 0 And Val(sales(5).Caption) > 0 Then
      WP = Val(SalesData(1, Cl)) / Val(sales(5).Caption)
      TW = TW + Val(SalesData(1, Cl))
  Else
      WP = 0
  End If
  If Check_Option(96) = 0 Then
    Print #PF, SalesData(0, Cl); Tab(25); Format(SalesData(1, Cl), "0.00"); Tab(33);
    Print #PF, Align(6, WP); EC(7)
  Else
    PrintLine SalesData(0, Cl), 3500: PrintLine SalesData(1, Cl), 5200
    PrintLine Format(WP, "0.00"), -1
  End If
Next Cl
If TW > 0 And Val(sales(5).Caption) > 0 Then
   WP = TW / Val(sales(5).Caption)
Else
   WP = 0
End If
If Check_Option(96) = 0 Then
 Print #PF, String(40, 45); EC(7)
 Print #PF, "Total Waste"; Tab(25); Format(TW, "0.00"); Tab(33); Align(6, WP); EC(7)
 Print #PF, "Net Sales"; Tab(25); ""; Tab(33); Align(6, Val(sales(5).Caption)); EC(7)
 
Else
 DrawLine
 PrintLine "Total", 3500: PrintLine Format(TW, "0.00"), 5200: PrintLine Format(WP, "0.00"), -1
 PrintLine "Net Sales", 3500: PrintLine "", 5200: PrintLine Format(Val(sales(5).Caption), "0.00"), -1
 
End If
 If Check_Option(96) = 0 Then
    Print #PF, EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
 Else
    FinishJob 1
 End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Print Hourly Sales ")
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
MyTimer = 0
Sales_Breakdown True, 0
Sales_Report -1, 0, Format(Date, "MM/DD/YYYY")
Print_Meal_Period -1
PrintPayments
PrintServeTime -1, Format(Date, "MM/DD/YYYY")
'PrintInterval 0, "All Terminals"
End Sub

Private Sub PrintPayments()
On Error GoTo ErrorHandler
Dim PrinterPort  As String
Dim OnLine       As String
Dim LN           As Integer
Dim DrawerSales  As Currency
Dim TotalPayment As Currency
Dim ChargeCheck  As Currency
Dim ChargeTip    As Currency
Dim Gratuity     As Currency
Dim PF           As Integer
ChargeTip = 0: TotalPayment = 0: ChargeCheck = 0
Heading = "Server Report"
If Check_Option(96) = 0 Then
 Printer_Drivers "Receipt"
 PF = FreeFile
 PrinterPort = "Print Payments"
tryagain:
Open PrinterPort For Output Shared As #PF
 CenterText = EC(8)
 LeftText = EC(9)
 'RightText = Chr(27) + Chr(97) + Chr(2)
 Print #PF, EC(0) 'Int printer
End If
If Check_Option(96) = 0 Then
  Print #PF, LeftText; EC(7)
  Print #PF, EC(1); CenterText; "Payment Report"; EC(4); EC(7)
  Print #PF, LeftText; EC(7)
  Print #PF, ; EC(7)
  Print #PF, EC(1); "Report Time : "; Format(Time, "HH:MM ampm"); EC(7)
  Print #PF, EC(1); "Report Date : "; Trim(Date); EC(4); EC(7)
  Print #PF, String(40, 45); EC(7)
End If
If Check_Option(96) = 0 Then
 Print #PF, "Payment Type           Amount      No. "
 Print #PF, String(40, 45); EC(7)
End If
''
Dim ListRefunds As FinancialItems
Set ListRefunds = New FinancialItems
ListAllReturns ListRefunds, 0
'****************************
Dim Payments As FinancialItems
Set Payments = New FinancialItems
SortCreditCards Payments, 0, "", 0, "", ""
TotalQty = 0: TotalPayment = 0: TotalPayment = 0
For Pa = 1 To Payments.count
   If Check_Option(96) = 0 Then
     Print #PF, Payments.Item(Pa).Tender; Tab(23); Align(6, Payments.Item(Pa).Amount); Tab(33); Format(Payments.Item(Pa).ItemQty, "##"); EC(7)
   End If
   TotalPayment = TotalPayment + Payments.Item(Pa).Amount
   TotalQty = TotalQty + Payments.Item(Pa).ItemQty
   TotalTip = TotalTip + Payments.Item(Pa).TipAmount
Next Pa
If Check_Option(96) = 0 Then
   Print #PF, String(40, 45); EC(7)
   Print #PF, "Total Collected: "; Tab(23); Align(6, Val(TotalPayment)); Tab(33); Format(Val(TotalQty), "##"); EC(7)
   Print #PF, "This Total Includes Tips"; EC(7)
   Print #PF, "Total Charged Tips : "; Tab(23); Align(6, Val(TotalTip)); EC(7)
End If
Dim ChargeBack As Currency
Dim QtyCb      As Integer
Dim GA         As Integer
ChargeBack = 0
For GA = 1 To Payments.count
  ChargeBack = ChargeBack + CalChargeBack(Payments.Item(Pa).Tender, Payments.Item(Pa).TipAmount)
Next GA
If ChargeBack > 0 Then
 If Check_Option(96) = 0 Then
   Print #PF, "Tip Charge Back    : "; Tab(23); Align(6, Val(ChargeBack)); EC(7)
 End If
End If
Set Payments = Nothing
If Gratuity > 0 Then
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Gratuity Charged   : "; Tab(23); Align(6, Val(Gratuity)); EC(7)
    End If
End If
If Check_Option(96) = 0 Then
    Print #PF, EC(6); EC(5)
    Close
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    Printer.EndDoc
End If
Exit Sub
ErrorHandler:
Err.Clear
Resume Next
End Sub
Private Sub Terms_Click()

End Sub
Private Sub Timer1_Timer()
MyTimer = MyTimer + 1
If MyTimer >= 15 Then
  Call MemoryCheck(Me)
End If
End Sub

