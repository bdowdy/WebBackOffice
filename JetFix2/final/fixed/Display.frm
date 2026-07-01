VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Display 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Daily Sales Sheet"
   ClientHeight    =   10575
   ClientLeft      =   210
   ClientTop       =   765
   ClientWidth     =   15120
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Display.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CompareHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   -6840
      ScaleHeight     =   5865
      ScaleWidth      =   8025
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton SSCommand13 
         Caption         =   "Print Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   14
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox CompareByTotal 
         Height          =   330
         ItemData        =   "Display.frx":1CCA
         Left            =   6840
         List            =   "Display.frx":1CD4
         TabIndex        =   9
         Text            =   "Day"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox SearchHow 
         Height          =   330
         ItemData        =   "Display.frx":1CE3
         Left            =   6840
         List            =   "Display.frx":1CF0
         TabIndex        =   7
         Text            =   "Day"
         Top             =   1320
         Width           =   975
      End
      Begin VSFlex8DAOCtl.VSFlexGrid Compare 
         Height          =   5415
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   6495
         _cx             =   1972186304
         _cy             =   1972184399
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
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   13816530
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
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
         Cols            =   5
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
      Begin VB.CommandButton SSCommand10 
         Caption         =   "Move Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   5
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton SSCommand11 
         Caption         =   "Move Foward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   6
         Top             =   3000
         Width           =   975
      End
      Begin BTNENHLib4.BtnEnh SScommand9 
         Height          =   1020
         Left            =   6720
         TabIndex        =   34
         Top             =   4440
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1799
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
         Picture         =   "C:\Icons\Exit.ico"
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "Display.frx":1D05
         textCaption     =   "Display.frx":1D1D
         textLT          =   "Display.frx":1D85
         textCT          =   "Display.frx":1D9D
         textRT          =   "Display.frx":1DB5
         textLM          =   "Display.frx":1DCD
         textRM          =   "Display.frx":1DE5
         textLB          =   "Display.frx":1DFD
         textCB          =   "Display.frx":1E15
         textRB          =   "Display.frx":1E2D
         colorBack       =   "Display.frx":1E45
         colorIntern     =   "Display.frx":1E6F
         colorMO         =   "Display.frx":1E99
         colorFocus      =   "Display.frx":1EC3
         colorDisabled   =   "Display.frx":1EED
         colorPressed    =   "Display.frx":1F17
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
      Begin VB.Label CompareDisp 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Trends"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label compareby 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales By"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Compare By"
         Height          =   255
         Left            =   6840
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog Touch 
      Left            =   480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "Display.frx":1F41
      Top             =   2780
      Width           =   2775
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2520
      Left            =   360
      TabIndex        =   40
      Top             =   240
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   12632256
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   14737632
      StartOfWeek     =   136970242
      TitleBackColor  =   4210752
      TitleForeColor  =   -2147483637
      CurrentDate     =   41277
   End
   Begin VB.PictureBox Memo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
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
      Height          =   3975
      Left            =   -9480
      ScaleHeight     =   3945
      ScaleWidth      =   10425
      TabIndex        =   15
      Top             =   10200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton SSCommand15 
         Caption         =   "Save Memo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7800
         TabIndex        =   20
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox SumMemo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   9975
      End
      Begin VB.CommandButton SSCommand14 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label MemoDate 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10095
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2775
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10215
      End
   End
   Begin VB.PictureBox ServerList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   4695
      Left            =   5280
      ScaleHeight     =   4665
      ScaleWidth      =   3345
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ListBox SList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   2895
      End
      Begin BTNENHLib4.BtnEnh SSCommand23 
         Height          =   735
         Left            =   120
         TabIndex        =   38
         Top             =   3840
         Width           =   3135
         _Version        =   589828
         _ExtentX        =   5530
         _ExtentY        =   1296
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
         textCaption     =   "Display.frx":1F77
         textLT          =   "Display.frx":1FDF
         textCT          =   "Display.frx":1FF7
         textRT          =   "Display.frx":200F
         textLM          =   "Display.frx":2027
         textRM          =   "Display.frx":203F
         textLB          =   "Display.frx":2057
         textCB          =   "Display.frx":206F
         textRB          =   "Display.frx":2087
         colorBack       =   "Display.frx":209F
         colorIntern     =   "Display.frx":20C9
         colorMO         =   "Display.frx":20F3
         colorFocus      =   "Display.frx":211D
         colorDisabled   =   "Display.frx":2147
         colorPressed    =   "Display.frx":2171
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
      Begin VB.Shape Shape4 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3375
         Left            =   120
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Server List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000001&
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "Display.frx":219B
      Top             =   4080
      Width           =   2775
   End
   Begin VSFlex8DAOCtl.VSFlexGrid ShowEdit 
      Height          =   5535
      Left            =   3960
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
      _cx             =   1972181013
      _cy             =   1972184611
      Appearance      =   1
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   12648447
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      Cols            =   2
      FixedRows       =   0
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
   Begin VB.PictureBox NewDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
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
      Height          =   1695
      Left            =   7440
      ScaleHeight     =   1665
      ScaleWidth      =   4185
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox StoreDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton SSCommand5 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton SSCommand6 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton KeyCal 
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComCtl2.MonthView Review 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   3360
      Left            =   1800
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   5927
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12582912
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   136970241
      CurrentDate     =   37165
   End
   Begin VB.ComboBox HistoryDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin BTNENHLib4.BtnEnh SSCommand2 
      Height          =   1020
      Left            =   9360
      TabIndex        =   23
      Top             =   9480
      Width           =   855
      _Version        =   589828
      _ExtentX        =   1508
      _ExtentY        =   1799
      _StockProps     =   66
      Caption         =   ">>"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Display.frx":21F3
      textLT          =   "Display.frx":2257
      textCT          =   "Display.frx":226F
      textRT          =   "Display.frx":2287
      textLM          =   "Display.frx":229F
      textRM          =   "Display.frx":22B7
      textLB          =   "Display.frx":22CF
      textCB          =   "Display.frx":22E7
      textRB          =   "Display.frx":22FF
      colorBack       =   "Display.frx":2317
      colorIntern     =   "Display.frx":2341
      colorMO         =   "Display.frx":236B
      colorFocus      =   "Display.frx":2395
      colorDisabled   =   "Display.frx":23BF
      colorPressed    =   "Display.frx":23E9
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
      Office2003ColorNorth=   14408667
      Office2003ColorSouth=   14211288
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   14408667
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   14408667
      Office2003ColorSouthPressed=   14211288
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand3 
      Height          =   1020
      Left            =   8400
      TabIndex        =   24
      Top             =   9480
      Width           =   855
      _Version        =   589828
      _ExtentX        =   1508
      _ExtentY        =   1799
      _StockProps     =   66
      Caption         =   "<<"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18.75
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
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Display.frx":2413
      textLT          =   "Display.frx":2477
      textCT          =   "Display.frx":248F
      textRT          =   "Display.frx":24A7
      textLM          =   "Display.frx":24BF
      textRM          =   "Display.frx":24D7
      textLB          =   "Display.frx":24EF
      textCB          =   "Display.frx":2507
      textRB          =   "Display.frx":251F
      colorBack       =   "Display.frx":2537
      colorIntern     =   "Display.frx":2561
      colorMO         =   "Display.frx":258B
      colorFocus      =   "Display.frx":25B5
      colorDisabled   =   "Display.frx":25DF
      colorPressed    =   "Display.frx":2609
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
      Office2003ColorNorth=   14408667
      Office2003ColorSouth=   14211288
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   14408667
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   14408667
      Office2003ColorSouthPressed=   14211288
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh End 
      Height          =   1020
      Left            =   12720
      TabIndex        =   25
      Top             =   9480
      Width           =   1755
      _Version        =   589828
      _ExtentX        =   3096
      _ExtentY        =   1799
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
      Shape           =   1
      Surface         =   15
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Display.frx":2633
      textLT          =   "Display.frx":269B
      textCT          =   "Display.frx":26B3
      textRT          =   "Display.frx":26CB
      textLM          =   "Display.frx":26E3
      textRM          =   "Display.frx":26FB
      textLB          =   "Display.frx":2713
      textCB          =   "Display.frx":272B
      textRB          =   "Display.frx":2743
      colorBack       =   "Display.frx":275B
      colorIntern     =   "Display.frx":2785
      colorMO         =   "Display.frx":27AF
      colorFocus      =   "Display.frx":27D9
      colorDisabled   =   "Display.frx":2803
      colorPressed    =   "Display.frx":282D
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
   Begin BTNENHLib4.BtnEnh SSCommand4 
      Height          =   1020
      Left            =   10320
      TabIndex        =   26
      Top             =   9480
      Width           =   975
      _Version        =   589828
      _ExtentX        =   1720
      _ExtentY        =   1799
      _StockProps     =   66
      Caption         =   $"Display.frx":2857
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
      Shape           =   1
      Surface         =   15
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Display.frx":2867
      textLT          =   "Display.frx":28DF
      textCT          =   "Display.frx":28F7
      textRT          =   "Display.frx":290F
      textLM          =   "Display.frx":2927
      textRM          =   "Display.frx":293F
      textLB          =   "Display.frx":2957
      textCB          =   "Display.frx":296F
      textRB          =   "Display.frx":2987
      colorBack       =   "Display.frx":299F
      colorIntern     =   "Display.frx":29C9
      colorMO         =   "Display.frx":29F3
      colorFocus      =   "Display.frx":2A1D
      colorDisabled   =   "Display.frx":2A47
      colorPressed    =   "Display.frx":2A71
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
   Begin BTNENHLib4.BtnEnh SSCommand1 
      Height          =   1020
      Left            =   11400
      TabIndex        =   27
      Top             =   9480
      Width           =   975
      _Version        =   589828
      _ExtentX        =   1720
      _ExtentY        =   1799
      _StockProps     =   66
      Caption         =   "Print"
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
      Shape           =   1
      Surface         =   15
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Display.frx":2A9B
      textLT          =   "Display.frx":2B05
      textCT          =   "Display.frx":2B1D
      textRT          =   "Display.frx":2B35
      textLM          =   "Display.frx":2B4D
      textRM          =   "Display.frx":2B65
      textLB          =   "Display.frx":2B7D
      textCB          =   "Display.frx":2B95
      textRB          =   "Display.frx":2BAD
      colorBack       =   "Display.frx":2BC5
      colorIntern     =   "Display.frx":2BEF
      colorMO         =   "Display.frx":2C19
      colorFocus      =   "Display.frx":2C43
      colorDisabled   =   "Display.frx":2C6D
      colorPressed    =   "Display.frx":2C97
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
      Office2003ColorNorth=   14408667
      Office2003ColorSouth=   14211288
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   14408667
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   14408667
      Office2003ColorSouthPressed=   14211288
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VSFlex8DAOCtl.VSFlexGrid SummarySheet 
      Height          =   8175
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   10935
      _cx             =   1972194136
      _cy             =   1972189268
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   14079702
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   8
      GridLinesFixed  =   8
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
      ScrollBars      =   2
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TTL To Account For = Less Discount"
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      Top             =   8800
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Red Date Indicates Attached Memo"
      Height          =   255
      Left            =   11400
      TabIndex        =   22
      Top             =   8880
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   8160
      Top             =   9360
      Width           =   6495
   End
   Begin VB.Label SearchY 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label LocName 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Summary Sheet"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   8850
      Left            =   3240
      Top             =   240
      Width           =   11415
   End
   Begin VB.Menu Quit 
      Caption         =   "Exit"
   End
   Begin VB.Menu History 
      Caption         =   "History Reports"
      Begin VB.Menu SaleSum 
         Caption         =   "Sales Summary"
      End
      Begin VB.Menu DeptTnd 
         Caption         =   "Department/Tender"
      End
      Begin VB.Menu DSD 
         Caption         =   "Department Sales Data"
         Begin VB.Menu DEPSales 
            Caption         =   "Department Sales"
         End
         Begin VB.Menu SBSP 
            Caption         =   "Sales By Individual"
         End
      End
      Begin VB.Menu CDa 
         Caption         =   "Category Sales Data"
         Begin VB.Menu CATSales 
            Caption         =   "Category Sales"
         End
         Begin VB.Menu CSBP 
            Caption         =   "Sales By Individual"
         End
      End
      Begin VB.Menu ORDDest 
         Caption         =   "Order Destination"
      End
      Begin VB.Menu WKSales 
         Caption         =   "Weekly Sales"
      End
      Begin VB.Menu MPeriods 
         Caption         =   "Meal Periods"
      End
   End
   Begin VB.Menu LBRHST 
      Caption         =   "Labor History"
   End
   Begin VB.Menu DaySum 
      Caption         =   "Daily Audit Report"
   End
   Begin VB.Menu MISHistory 
      Caption         =   "Misc History"
      Begin VB.Menu Pdouts 
         Caption         =   "Paid Outs"
         Begin VB.Menu SBN 
            Caption         =   "Summary By Name"
         End
         Begin VB.Menu SBD 
            Caption         =   "Summary By Description"
         End
         Begin VB.Menu LBD 
            Caption         =   "List By Description"
         End
      End
      Begin VB.Menu Tdrs 
         Caption         =   "Tenders"
      End
      Begin VB.Menu Dps 
         Caption         =   "Deposits"
      End
      Begin VB.Menu Disc 
         Caption         =   "Discounts"
      End
      Begin VB.Menu VRPT 
         Caption         =   "Voids"
      End
      Begin VB.Menu CashBalance 
         Caption         =   "Cash Balance Sheet"
      End
      Begin VB.Menu CCounts 
         Caption         =   "Customer Counts"
      End
      Begin VB.Menu WRP 
         Caption         =   "Waste"
      End
      Begin VB.Menu Refund 
         Caption         =   "Item Refunded"
      End
      Begin VB.Menu NoSaleReport 
         Caption         =   "No Sales"
      End
      Begin VB.Menu IBT 
         Caption         =   "Items By Time"
      End
      Begin VB.Menu XCPT 
         Caption         =   "Export Compeat"
      End
   End
   Begin VB.Menu PMix 
      Caption         =   "Product Mix"
   End
   Begin VB.Menu DSRBalance 
      Caption         =   "Balance Sheet"
   End
   Begin VB.Menu SRC 
      Caption         =   "Sales Recap"
   End
   Begin VB.Menu FindRS 
      Caption         =   "Find Top 10 Sales For"
      Begin VB.Menu RS 
         Caption         =   "Monday"
         Index           =   0
         Begin VB.Menu TBM 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBM 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
      Begin VB.Menu RS 
         Caption         =   "Tuesday"
         Index           =   1
         Begin VB.Menu TBT 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBT 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
      Begin VB.Menu RS 
         Caption         =   "Wednesday"
         Index           =   2
         Begin VB.Menu TBW 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBW 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
      Begin VB.Menu RS 
         Caption         =   "Thursday"
         Index           =   3
         Begin VB.Menu TBTH 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBTH 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
      Begin VB.Menu RS 
         Caption         =   "Friday"
         Index           =   4
         Begin VB.Menu TBFR 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBFR 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
      Begin VB.Menu RS 
         Caption         =   "Saturday"
         Index           =   5
         Begin VB.Menu TBSAT 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBSAT 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
      Begin VB.Menu RS 
         Caption         =   "Sunday"
         Index           =   6
         Begin VB.Menu TBSUN 
            Caption         =   "Top 10"
            Index           =   0
         End
         Begin VB.Menu TBSUN 
            Caption         =   "Bottom 10"
            Index           =   1
         End
      End
   End
   Begin VB.Menu WeekS 
      Caption         =   "Weekly Graphs"
      Begin VB.Menu WSG 
         Caption         =   "NetSales"
      End
      Begin VB.Menu WDS 
         Caption         =   "Discounts"
      End
      Begin VB.Menu WLP 
         Caption         =   "Labor %"
      End
      Begin VB.Menu WLH 
         Caption         =   "Labor Hours"
      End
      Begin VB.Menu VS 
         Caption         =   "Voided Sales"
      End
      Begin VB.Menu DMS 
         Caption         =   "Department Sales"
      End
      Begin VB.Menu MPS 
         Caption         =   "Meal Period Sales"
      End
      Begin VB.Menu OSS 
         Caption         =   "Over/Short"
      End
   End
   Begin VB.Menu EditDay 
      Caption         =   "Edit"
      Begin VB.Menu RCD 
         Caption         =   "ReClose Day"
      End
      Begin VB.Menu EDDay 
         Caption         =   "Edit Day"
         Begin VB.Menu EditHistory 
            Caption         =   "Edit Sales"
         End
         Begin VB.Menu EditDeposit 
            Caption         =   "Edit Deposits"
         End
      End
      Begin VB.Menu CgDate 
         Caption         =   "Change Sales Date"
      End
      Begin VB.Menu DataMerge 
         Caption         =   "Merge Data"
      End
      Begin VB.Menu MSBUP 
         Caption         =   "Make Sales Backup"
      End
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditRecord  As Boolean
Dim EditItem       As String
Public Scroll      As Integer
Dim ClrC           As Boolean
Dim SearchDate     As String
Private SearchSpan As Long
Dim AdvSearch      As Long
Private MDate      As String
Private EditDate   As String
Private wYear      As Long
Private EditType   As String
Private PaymentType  As String
Private SqlBetweenStatment As String
Private AccessBetweenStatment As String
Private BusDayEndsAt As String
'// temp not sure if we need this
Private Function Non_CashTenders(Db As ADODB.Connection, InDate As String, EDate As String) As Currency
Dim RS As ADODB.Recordset
'ConnectADODB SvrPath.TerminalPath & ProductHistory, Db


Sql = "Select [Media], Sum([Amount]) as TAmount,Sum([Tip]) as TTip from [CheckPayments]"
Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & CDate(InDate & " " & GetStartTime) & "'"
Sql = Sql & " and '" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "'"
Sql = Sql & " Group By [Media]"
ConnectADORS (Sql), Db, RS
Non_CashTenders = 0
'// sum the total amount of tenders
    Do While Not RS.EOF
     If RS![Media] & "" <> "Cash" Then
      If Val(RS![TAmount] & "") <> 0 Then
          Non_CashTenders = Non_CashTenders + Val(RS![TAmount] & "") + RS.fields("TTip").Value
      End If
     End If
     RS.MoveNext
    Loop
RS.Close

End Function

Private Sub EditData(DisplayDate As String)
On Error GoTo ErrorHandler
If Trim(DisplayDate) = "" Then
 See_Msg "Invalid Date", 5, ""
 Exit Sub
End If
EditDate = DisplayDate
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Cdb) = False Then Exit Sub
Sql = "Select * from [Sales] where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(DisplayDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Cdb, RS
ShowEdit.Clear flexClearEverywhere, flexClearText
CenterObject Me, ShowEdit
With ShowEdit
  .TextMatrix(0, 0) = "Edit Record For"
  .TextMatrix(1, 0) = "Gross Sales"
  .TextMatrix(2, 0) = "Net Sales"
  .TextMatrix(3, 0) = "Sales Tax"
  .TextMatrix(4, 0) = "Special Tax"
  .TextMatrix(5, 0) = SurChargeTxt
  .TextMatrix(6, 0) = "Voids"
  .TextMatrix(7, 0) = "Discounts"
  .TextMatrix(8, 0) = "NonCash"
  .TextMatrix(9, 0) = "Paid Outs"
  .TextMatrix(10, 0) = "Deposits"
  .TextMatrix(11, 0) = "Non Taxed Sales"
  .TextMatrix(13, 0) = "TTl Accounted For"
  .TextMatrix(12, 0) = "TTl To Account For"
  .TextMatrix(14, 0) = "Over / Short"
  .TextMatrix(15, 0) = "Labor Dollars"
  .TextMatrix(16, 0) = "Labor Hours"
  .TextMatrix(17, 0) = "Labor Percentage"
End With
Dim DayInfo() As Currency
ReDim DayInfo(16, 8)
ShowEdit.TextMatrix(0, 1) = Format(DisplayDate, "mm/dd/yyyy")
  Do While Not RS.EOF
    If Check_Null(RS![Date]) > "" Then
        DayInfo(0, 1) = CStr(Val(Check_Null(RS![Gross Sales])))
        DayInfo(1, 1) = CStr(Val(Check_Null(RS![Net Sales])))
        DayInfo(2, 1) = CStr(Val(Check_Null(RS![Sales Tax])))
        DayInfo(3, 1) = CStr(Val(Check_Null(RS![Special Tax])))
        DayInfo(4, 1) = CStr(Val(Check_Null(RS![Liquor Tax])))
        DayInfo(5, 1) = CStr(Val(Check_Null(RS![Voids])))
        DayInfo(6, 1) = CStr(Val(Check_Null(RS![Discounts])))
        DayInfo(7, 1) = CStr(Val(Check_Null(RS![Tenders])))
        DayInfo(8, 1) = CStr(Val(Check_Null(RS![Paid Outs])))
        DayInfo(9, 1) = CStr(Val(Check_Null(RS![Deposits])))
        DayInfo(10, 1) = CStr(Val(Check_Null(RS![None Taxed Sales])))
        DayInfo(11, 1) = DayInfo(1, 1) + DayInfo(2, 1) + DayInfo(3, 1) + DayInfo(4, 1) + (Val(RS![Exempt Sales] & "") - Val(RS![Charged Gratuity] & ""))
        DayInfo(12, 1) = DayInfo(7, 1) + DayInfo(8, 1) + DayInfo(9, 1)
        DayInfo(13, 1) = CStr(Val(Check_Null(RS![Over Short])))
        DayInfo(13, 1) = DayInfo(12, 1) - DayInfo(11, 1)
        DayInfo(14, 1) = CStr(Val(Check_Null(RS![Labor Dollars])))
        DayInfo(15, 1) = CStr(Val(Check_Null(RS![Labor Hours])))
        DayInfo(16, 1) = CStr(Val(Check_Null(RS![Labor Percentage])))
    End If
  RS.MoveNext
  Loop
DbClose Cdb, RS
Dim FI As Integer, DN As Integer
For DN = 1 To 1
For FI = 0 To 16
        ShowEdit.TextMatrix(FI + 1, DN) = Format(DayInfo(FI, DN), "0.00")
Next FI
Next DN
With ShowEdit
.cols = 2
.ColWidth(0) = 1600 * WidScale1024
.ColWidth(1) = 1000 * WidScale1024
.rows = 50
.width = .ColWidth(0) + .ColWidth(1) + 400
.Height = (.RowHeight(1) * (.rows - 30)) + 600
.Visible = True
.Refresh
ShowEdit.TextMatrix(19, 1) = "Save"
ShowEdit.TextMatrix(20, 1) = "Exit"
ShowEdit.TextMatrix(21, 1) = "Rebuild Mix"
End With
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Edit Data ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub EditDayData(DisplayDate As String)
On Error GoTo ErrorHandler
If Trim(DisplayDate) = "" Then Exit Sub
EditDate = DisplayDate
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Cdb) = False Then Exit Sub
Sql = "Select * from [Deposits] where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(DisplayDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Cdb, RS
If RS.EOF Then
   CenterObject Me, NewDate
   Label3.Caption = "Change Deposit"
   StoreDate.Text = Format(SummarySheet.TextMatrix(10, SummarySheet.Col), "0.00")
   KeyCal.Caption = "Show Calculator"
   NewDate.ZOrder 0
   NewDate.Visible = True
   DbClose Cdb, RS
   Exit Sub
Else
  ShowEdit.Clear flexClearEverywhere, flexClearText
  CenterObject Me, ShowEdit
  With ShowEdit
        .cols = 2
        .rows = 50
        .ColWidth(0) = 1600 * WidScale1024
        .ColWidth(1) = 1000 * WidScale1024
        .width = .ColWidth(0) + .ColWidth(1) + 400
        .Height = (.RowHeight(1) * (.rows - 30)) + 600
        .Visible = True
        .Refresh
        ShowEdit.TextMatrix(0, 0) = "Date"
        ShowEdit.TextMatrix(0, 1) = "Amount"
        ShowEdit.TextMatrix(19, 1) = "Save"
        ShowEdit.TextMatrix(20, 1) = "Exit"
 End With
 ShowEdit.Visible = True
 ShowEdit.ZOrder 0
 Dim FI As Integer, DN As Integer
 FI = -1
  Do While Not RS.EOF
    If Check_Null(RS![Date]) > "" Then
        FI = FI + 1
        ShowEdit.TextMatrix(FI + 1, 0) = Format(RS![Date], "MM/DD/YYYY")
        ShowEdit.TextMatrix(FI + 1, 1) = Format(Val(Check_Null(RS![Amount] & "")), "0.00")
    End If
  RS.MoveNext
  Loop
End If
DbClose Cdb, RS
Sleep 0
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Edit Data ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SaveChanges(EditTable As String, SearchDate As String)
On Error GoTo ErrorHandler
Dim Sql As String
Dim F   As Integer
Dim DTotal As Currency
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Cdb) = False Then Exit Sub
      Select Case EditTable
         Case Is = "Deposits"
            Sql = "Delete from [Deposits] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'"
            Cdb.Execute (Sql)
            Sql = "Select * from [Deposits] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'"
            ConnectADORS (Sql), Cdb, RS
            With ShowEdit
            DTotal = 0
             For F = 1 To .rows - 1
              If .TextMatrix(F, 1) = "" Then Exit For
              RS.AddNew
              RS![Date] = Format(SearchDate, "MM/DD/YYYY")
              RS![Amount] = Format(.TextMatrix(F, 1), "0.00")
              DTotal = DTotal + Val(.TextMatrix(F, 1))
              RS.Update
             Next F
            End With
            Sql = "Update [Sales] Set [Deposits] = " & Format(Val(DTotal), "0.00") & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(Format(CDate(SearchDate), "MM/DD/YYYY")), "yyyy-mm-dd hh:nn:ss") & "'"
            Cdb.Execute Sql
            DbClose Cdb, RS
         Case Else
            Sql = "Select * from [Sales] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'"
            ConnectADORS Sql, Cdb, RS
            With ShowEdit
            If RS.EOF Then RS.AddNew
            RS![Date] = SearchDate
            RS![Gross Sales] = .TextMatrix(1, 1)
            RS![Net Sales] = .TextMatrix(2, 1)
            RS![Sales Tax] = .TextMatrix(3, 1)
            RS![Special Tax] = .TextMatrix(4, 1)
            RS![Liquor Tax] = .TextMatrix(5, 1)
            RS![Voids] = .TextMatrix(6, 1)
            RS![Discounts] = .TextMatrix(7, 1)
            RS![Tenders] = .TextMatrix(8, 1)
            RS![Paid Outs] = .TextMatrix(9, 1)
            RS![Deposits] = .TextMatrix(10, 1)
            RS![None Taxed Sales] = .TextMatrix(11, 1)
            RS![Accounted For] = .TextMatrix(12, 1)
            RS![Adjusted Sales] = .TextMatrix(13, 1)
            RS![Over Short] = .TextMatrix(14, 1)
            RS![Labor Dollars] = .TextMatrix(15, 1)
            RS![Labor Hours] = .TextMatrix(16, 1)
            RS![Labor Percentage] = .TextMatrix(17, 1)
            RS.Update
            End With
        End Select
        DbClose Cdb, RS
Form_Load
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Save Changes ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub RePrintSales(PrintDate As String)
On Error GoTo ErrorHandler
Dim DeptTTl As Currency
Dim Cdb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim DRS As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Cdb) = False Then Exit Sub
 Sql = "Select * from [Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(PrintDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Cdb, CRS
    If Not CRS.EOF Then
        If Check_Option(96) = 1 Then
            SetPage ReportView.View, "Reprinted Sales Summary"
            PrintCLine "For " & Format(PrintDate, "MM/DD/YYYY"), -1, 0
            DrawLine
             Sql = "Select * from [Department Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Department Sales", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            DeptTTl = 0
            Do While Not DRS.EOF
            If DRS![Amount] > 0 Then
             PrintLine DRS![Name], 4500: PrintLine Format(DRS![Amount], CurrencyPreFix), 6500: PrintLine Format(DRS![Quantity], "###"), -1
             DeptTTl = DeptTTl + DRS![Amount]
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            PrintLine "Department Total", 4500: PrintLine Format(DeptTTl, "0.00"), -1
            DrawLine
            End If
            
            Sql = "Select * from [Category Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
             PrintLine "Category Sales", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
             DrawLine
             DeptTTl = 0
             Do While Not DRS.EOF
              If DRS![Amount] > 0 Then
               PrintLine DRS![Name], 4500: PrintLine Format(DRS![Amount], CurrencyPreFix), 6500: PrintLine Format(DRS![Quantity], "###"), -1
               DeptTTl = DeptTTl + DRS![Amount]
              End If
              DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            PrintLine "Category Total", 4500: PrintLine Format(DeptTTl, "0.00"), -1
            DrawLine
            End If
            Sql = "Select * from [Meal Periods] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(PrintDate)), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Meal Periods", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            Do While Not DRS.EOF
             If Val(DRS![Amount] & "") > 0 Then
             PrintLine DRS![Name], 4500: PrintLine Format(DRS![Amount], CurrencyPreFix), 6500: PrintLine Format(DRS![Quantity], "###"), -1
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            End If
            Sql = "Select * from [Hourly Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Hourly Sales", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            Do While Not DRS.EOF
             If Val(DRS![Amount] & "") > 0 Then
               PrintLine DRS![Hour] & "", 4500: PrintLine Format(DRS![Amount], CurrencyPreFix), 6500: PrintLine Format(DRS![Quantity], "###"), -1
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            End If
            PrintLine "Gross Sales", 4500: PrintLine Format(Val(Check_Null(CRS![Gross Sales])), CurrencyPreFix), -1
            PrintLine "Net Sales", 4500: PrintLine Format(Val(Check_Null(CRS![Net Sales])), CurrencyPreFix), -1
            PrintLine "Non Taxed Sales", 4500: PrintLine Format(Val(Check_Null(CRS![None Taxed Sales])), CurrencyPreFix), -1
            PrintLine "Sales Tax", 4500: PrintLine Format(Val(Check_Null(CRS![Sales Tax])), CurrencyPreFix), -1
            PrintLine "Special Tax", 4500: PrintLine Format(Val(Check_Null(CRS![Special Tax])), CurrencyPreFix), -1
            PrintLine SurChargeTxt, 4500: PrintLine Format(Val(Check_Null(CRS![Liquor Tax])), CurrencyPreFix), -1
            PrintLine "Non Sale Revenue", 4500: PrintLine Format(Val(CRS![Exempt Sales] & "") - Val(CRS![Charged Gratuity] & ""), "0.00"), -1
            PrintLine "Refunds", 4500: PrintLine Format(Val(CRS![Refunds] & ""), "0.00"), -1
            DrawLine
            Sql = "Select * from [Discounts] Where TRY_CAST([Date] AS datetime) = '" & CDate(Trim(PrintDate)) & "'"
            ConnectADORS Sql, Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Discounts", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            Do While Not DRS.EOF
             If Val(DRS![Amount] & "") > 0 Then
              PrintLine DRS![Name], 4500: PrintLine Format(Val(Check_Null(DRS![Amount])), CurrencyPreFix), 6500: PrintLine Format(Val(Check_Null(DRS![Quantity])), "###"), -1
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            End If
            PrintLine "Total Discount", 4500: PrintLine Format(Val(Check_Null(CRS![Discounts])), CurrencyPreFix), -1
            PrintLine "Total Voids", 4500: PrintLine Format(Val(Check_Null(CRS![Voids])), CurrencyPreFix), -1
            Sql = "Select * from [Paid Outs] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(PrintDate)), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Paid Outs", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            Do While Not DRS.EOF
             PrintLine DRS![Name], 4500: PrintLine Format(DRS![Amount], CurrencyPreFix), 6500: PrintLine Format(DRS![Quantity], "###"), -1
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            End If
            PrintLine "Total Paid Outs", 4500: PrintLine Format(Val(Check_Null(CRS![Paid Outs])), CurrencyPreFix), -1
            If Val(CRS![Credit Card Tips]) > 0 Then
              PrintLine "Credit Card Tip", 4500: PrintLine Format(Val(Check_Null(CRS![Credit Card Tips])), CurrencyPreFix), -1
            End If
            DrawLine
            Sql = "Select * from [Tenders] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Tender Types", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            Do While Not DRS.EOF
             PrintLine DRS![Name], 4500: PrintLine Format(Val(Check_Null(DRS![Amount])), CurrencyPreFix), 6500: PrintLine Format(Val(Check_Null(DRS![Quantity])), "###"), -1
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            End If
            PrintLine "Non Cash Tenders", 4500: PrintLine Format(Val(Check_Null(CRS![Tenders])), CurrencyPreFix), -1
            DrawLine
            Sql = "Select * from [Deposits] Where TRY_CAST([Date] AS datetime) = '" & CDate(Trim(PrintDate)) & "'"
            ConnectADORS Sql, Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Deposits", 4500: PrintLine "Amount", -1
            DrawLine
            Do While Not DRS.EOF
             PrintLine "Deposit", 4500: PrintLine Format(DRS![Amount], CurrencyPreFix), -1
             DRS.MoveNext
            Loop
            SetNothing DRS
            DrawLine
            End If
            PrintLine "Cash Deposits", 4500: PrintLine Format(Val(Check_Null(CRS![Deposits])), CurrencyPreFix), -1
            PrintLine "Deposits Collected", 4500: PrintLine Format(Val(Check_Null(CRS![Deposits Collected])), CurrencyPreFix), -1
            PrintLine "House Charges", 4500: PrintLine Format(Val(Check_Null(CRS![House Charges])), CurrencyPreFix), -1
            PrintLine "Tip Charge Back", 4500: PrintLine Format(Val(Check_Null(CRS![Tip Charge Back])), CurrencyPreFix), -1
            DrawLine
            PrintLine "Total Accounted For", 4500: PrintLine Format(Val(Check_Null(CRS![Accounted For])), CurrencyPreFix), -1
            
            PrintLine "Total To Account For", 4500: PrintLine Format(Val(Check_Null(CRS![Adjusted Sales])) + (Val(CRS![Exempt Sales]) - Val(CRS![Charged Gratuity])), CurrencyPreFix), -1
            'Val (CRS![Exempt Sales]) - Val(CRS![Charged Gratuity])
            
            DrawLine
            PrintLine "Over / Short", 4500: PrintLine Format(Val(Check_Null(CRS![Accounted For])) - Val(Check_Null(CRS![Adjusted Sales]) + (Val(CRS![Exempt Sales]) - Val(CRS![Charged Gratuity]))), CurrencyPreFix), -1  'Val(CRS![Exempt Sales]) -
            DrawLine
            PrintLine "Labor Summary", -1
            DrawLine
            PrintLine "Labor Dollars", 4500: PrintLine Format(Val(Check_Null(CRS![Labor Dollars])), CurrencyPreFix), -1
            PrintLine "Labor Hours", 4500: PrintLine Format(Val(Check_Null(CRS![Labor Hours])), "0.00"), -1
            PrintLine "Labor Percentage", 4500: PrintLine Format(Val(Check_Null(CRS![Labor Percentage])), "%0.00"), -1
            DrawLine
            PrintLine "Summary Data", -1
            DrawLine
            '***********************
            If CRS![Table Average] > 0 And System_config = FullService Then
             PrintLine "Average Table Turn", 4500: PrintLine Format(Val(Check_Null(CRS![Table Average])), "0.00"), -1
            End If
            Sql = "Select * from [Order Destination] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            PrintLine "Order Type", 4500: PrintLine "Amount", 6500: PrintLine "Quantity", -1
            DrawLine
            Do While Not DRS.EOF
             If Val(DRS![Amount] & "") > 0 Then
              PrintLine DRS![Order Destination], 4500: PrintLine Format(Val(DRS![Amount] & ""), CurrencyPreFix), 6500: PrintLine Format(Val(Check_Null(DRS![Quantity])), "###"), -1
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            End If
         '***************************
            DrawLine
            FinishJob 1
        Else
            Dim PrinterPort As String, Heading As String
            Printer_Drivers "Receipt"
            CenterText = EC(8)
            LeftText = EC(9)
            
            Dim PF As Integer
            PF = FreeFile
            PrinterPort = "Reprint Sales"
OneMoreTime:
            Open PrinterPort For Output Shared As #PF
            Print #PF, EC(1); EC(4); CenterText; "Reprinted Sales"; EC(2); EC(7)
            Print #PF, EC(1); EC(4); CenterText; "For " & PrintDate; EC(2); EC(7)
            Print #PF, EC(4); LeftText; EC(7)
            Print #PF, EC(7)
            '****************
            Sql = "Select * from [Department Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Department Sales"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             If DRS![Amount] > 0 Then
               Print #PF, DRS![Name], Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If
            '******************************111111111
            Sql = "Select * from [Category Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Category Sales"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             If DRS![Amount] > 0 Then
               Print #PF, DRS![Name], Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If
            Sql = "Select * from [Meal Periods] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(PrintDate)), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Meal Periods"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             Print #PF, DRS![Name], Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If
            Sql = "Select * from [Hourly Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS Sql, Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Hourly Sales"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             If Val(DRS![Amount]) > 0 Then
               Print #PF, DRS![Hour], Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             End If
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If

            Print #PF, "Item"; Tab(30); "Amount"
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Gross Sales"; Tab(30); Format(Val(Check_Null(CRS![Gross Sales])), CurrencyPreFix); EC(7)
            Print #PF, "Net Sales"; Tab(30); Format(Val(Check_Null(CRS![Net Sales])), CurrencyPreFix); EC(7)
            Print #PF, "Non Taxed Sales"; Tab(30); Format(Val(Check_Null(CRS![None Taxed Sales])), CurrencyPreFix); EC(7)
            Print #PF, "Sales Tax"; Tab(30); Format(Val(Check_Null(CRS![Sales Tax])), CurrencyPreFix); EC(7)
            Print #PF, "Special Tax"; Tab(30); Format(Val(Check_Null(CRS![Special Tax])), CurrencyPreFix); EC(7)
            Print #PF, SurChargeTxt; Tab(30); Format(Val(Check_Null(CRS![Liquor Tax])), CurrencyPreFix); EC(7)
            Print #PF, "Non Sale Revenue", Tab(30); Format(Val(CRS![Exempt Sales] & "") - Val(CRS![Charged Gratuity] & ""), "0.00"); EC(7)
            Print #PF, "Refunds"; Tab(30); Format(Val(Check_Null(CRS![Refunds])), CurrencyPreFix); EC(7)
            
            Print #PF, String(40, 45); EC(7)
            Sql = "Select * from [Discounts] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Discounts"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             Print #PF, DRS![Name], Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If
            
            Print #PF, "Total Discount"; Tab(30); Format(Val(Check_Null(CRS![Discounts])), CurrencyPreFix); EC(7)
            Print #PF, "Total Voids"; Tab(30); Format(Val(Check_Null(CRS![Voids])), CurrencyPreFix); EC(7)
            Sql = "Select * from [Paid Outs] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS

            If Not DRS.EOF Then
            Print #PF, "Paid Outs"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             Print #PF, DRS![Name], Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If

            Print #PF, "Total Paid Outs"; Tab(30); Format(Val(Check_Null(CRS![Paid Outs])), CurrencyPreFix); EC(7)
            If Val(Check_Null(CRS![Credit Card Tips])) > 0 Then
              Print #PF, "Credit Card Tip"; Tab(30); Format(Val(Check_Null(CRS![Credit Card Tips])), CurrencyPreFix); EC(7)
            End If
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Non Cash Tenders"; Tab(30); Format(Val(Check_Null(CRS![Tenders])), CurrencyPreFix); EC(7)
            Sql = "Select * from [Deposits] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(Trim(PrintDate))), "yyyy-mm-dd hh:nn:ss") & "'"
            ConnectADORS (Sql), Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Deposits"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             Print #PF, "Deposit", Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             DRS.MoveNext
            Loop
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            End If
            Print #PF, "Cash Deposits"; Tab(30); Format(Val(Check_Null(CRS![Deposits])), CurrencyPreFix); EC(7)
            Print #PF, "Deposits Collected"; Tab(30); Format(Val(Check_Null(CRS![Deposits Collected])), CurrencyPreFix); EC(7)
            Print #PF, "House Charges"; Tab(30); Format(Val(Check_Null(CRS![House Charges])), CurrencyPreFix); EC(7)
            Print #PF, "Tip Charge Back"; Tab(30); Format(Val(Check_Null(CRS![Tip Charge Back])), CurrencyPreFix); EC(7)
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Total Accounted For"; Tab(30); Format(Val(Check_Null(CRS![Accounted For])), CurrencyPreFix); EC(7)
            Print #PF, "Total To Account For"; Tab(30); Format(Val(Check_Null(CRS![Adjusted Sales]) + (Val(CRS![Exempt Sales]) - Val(CRS![Charged Gratuity]))), CurrencyPreFix); EC(7)
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Over / Short"; Tab(30); Format(Val(Check_Null(CRS![Accounted For])) - Val(Check_Null(CRS![Adjusted Sales] + (Val(CRS![Exempt Sales]) - Val(CRS![Charged Gratuity])))), CurrencyPreFix); EC(7)
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Labor Summary"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Labor Dollars"; Tab(30); Format(Val(Check_Null(CRS![Labor Dollars])), CurrencyPreFix); EC(7)
            Print #PF, "Labor Hours"; Tab(30); Format(Val(Check_Null(CRS![Labor Hours])), "0.00"); EC(7)
            Print #PF, "Labor Percentage"; Tab(30); Format(Val(Check_Null(CRS![Labor Percentage])), "%0.00"); EC(7)
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Summary Data"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Sql = "Select * from [Order Destination] Where TRY_CAST([Date] AS datetime) = '" & CDate(Trim(PrintDate)) & "'"
            ConnectADORS Sql, Cdb, DRS
            If Not DRS.EOF Then
            Print #PF, "Order Type"; Tab(30); "Amount"; EC(7)
            Print #PF, String(40, 45); EC(7)
            Do While Not DRS.EOF
             Print #PF, DRS![Order Destination] & "", Tab(30); Format(DRS![Amount], CurrencyPreFix); EC(7)
             DRS.MoveNext
            Loop
            End If
            SetNothing DRS
            Print #PF, String(40, 45); EC(7)
            If CRS![Table Average] > 0 And System_config = FullService Then
              Print #PF, "Average Table Turn"; Tab(30); Format(Val(Check_Null(CRS![Table Average])), "0.00"); EC(7)
            End If
            Print #PF, EC(6) & EC(6) & EC(5)
            Close #PF
            SendJobToPrinter PrinterPort, "Receipt", False
        End If
    End If
DbClose Cdb, CRS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reprint sales ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub bdc_Click()
End Sub

Private Sub BtnEnh1_Click()
End Sub

Private Sub B10_Click()
End Sub
Private Sub CashBalance_Click()
LoadDates "Select Distinct [Date] from [Sales] Where not isnull([date]) order by [Date] asc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
On Error GoTo ErrorHandler
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim Data            As String
Dim H$
Dim TSales          As Currency
Dim Mtotal          As Currency
Dim TCash           As Currency
Dim F$
Dim OverSales       As Currency
Dim Revenus         As Currency
Dim NonCashTenders  As Currency
Dim Collected       As Currency
Dim POUT            As Currency
Dim RFND            As Currency
Dim TPO             As Currency
Dim PIN             As Currency
SetPage ReportView.View, "Cash Balance Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
  Sql = "Select [Media], sum([Amount])  as TSold  from [CheckPayments] "
  Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and [Media] = 'Cash'"
  Sql = Sql & " Group By [Media]"
 ConnectADORS ReturnDateString(Sql), Db, RS
 TCash = 0: Data = ""
 If Not RS.EOF Then
    Do While Not RS.EOF
     TCash = TCash + Val(RS.fields("TSold").Value & "")
     RS.MoveNext
    Loop
  'SetNothing RS
 End If
Data = Data & "Cash Collected" & "|"
Data = Data & Format(TCash, "0.00") & ";"
DbClose Db, RS
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
  Sql = "Select [Type], sum([Amount])  as TSold  from [Paid Outs] "
  Sql = Sql & " Where  TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(CDate(SearchDates.Item(1).EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Group By [Type]"
 ConnectADORS Sql, Db, RS
 POUT = 0
 If Not RS.EOF Then
    Do While Not RS.EOF
     POUT = POUT + Val(RS.fields("TSold").Value & "")
     RS.MoveNext
    Loop
  SetNothing RS
  Mtotal = Mtotal + POUT
 End If
 Sql = "Select Sum([Exempt Sales]) as ESale, Sum([Charged Gratuity]) as TGrat"
 Sql = Sql & " from [Sales] Where "
 Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SearchDates.Item(1).EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 ConnectADORS Sql, Db, RS
 PIN = 0
 If Not RS.EOF Then
  PIN = Val(RS![ESale] & "") - Val(RS![TGrat] & "")
  If PIN < 0 Then PIN = 0
  PIN = (PIN * -1)
  SetNothing RS
  Mtotal = Mtotal + PIN
 End If
 Sql = "Select [Name], sum([Amount])  as TSold  from [Refunds] "
 Sql = Sql & " Where  TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " and '" & Format(CDate(CDate(SearchDates.Item(1).EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " and instr([Name],'Credit Card') = 0 "
 Sql = Sql & " Group By [Name]"
 ConnectADORS Sql, Db, RS
 If Not RS.EOF Then
   RFND = 0 ': Data = ""
    Do While Not RS.EOF
     RFND = RFND + Val(RS.fields("TSold").Value & "")
     RS.MoveNext
    Loop
   SetNothing RS
   Mtotal = Mtotal + RFND
 End If
DbClose Db, RS
ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
  Sql = "Select [Server Name], sum([Charged Tips])  as CTip, Sum([Gratuity]) as GChg,sum([Tips Paid Out]) as TPO  from [Tip Income] "
  Sql = Sql & " Where  TRY_CAST([Date] AS date) Between '" & CDate(SearchDates.Item(1).StartDate) & "'"
  Sql = Sql & " and '" & CDate(SearchDates.Item(1).EndDate) & "'"
  Sql = Sql & " Group By [Server Name]"
 ConnectADORS Sql, Db, RS
 TPO = 0
 If Not RS.EOF Then
    Do While Not RS.EOF
     TPO = TPO + (RS.fields("CTip").Value + RS.fields("GChg").Value) - RS.fields("TPO").Value
     RS.MoveNext
    Loop
  SetNothing RS
  Mtotal = Mtotal + TPO
 End If
 Data = Data & "Cash Adjustments" & "|"
 Data = Data & Format(Mtotal, "0.00") & ";"
 Data = Data & " ** Paid Outs" & "|"
 Data = Data & Format(POUT, "0.00") & ";"
 Data = Data & " ** Cash Refunds" & "|"
 Data = Data & Format(RFND, "0.00") & ";"
 Data = Data & " ** Tips Paid" & "|"
 Data = Data & Format(TPO, "0.00") & ";"
 Data = Data & " ** Cash Paid In" & "|"
 Data = Data & Format(PIN, "0.00") & ";"
 H = "Name|Amount;"
 F = "+<2500|+>2000;"
 ReportView.View.TableBorder = tbTopBottom
 ReportView.View.addTable F, H, Data, vbBlack, vbWhite
 Data = "Cash Deposit" & "|"
 Data = Data & Format(TCash - Mtotal, "0.00") & ";"
 DbClose Db, RS
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
 Sql = "Select  Sum([Exempt Sales]) as ESale, Sum([Charged Gratuity]) as TGrat"
 Sql = Sql & " ,Sum([Tenders]) as TTender, Sum([Paid Outs]) as TPaidOut, Sum([Deposits]) as TDeposits"
 Sql = Sql & " ,Sum([Net Sales]) as TNet ,Sum([Sales Tax]) as Stax, Sum([Special Tax]) as SpTax"
 Sql = Sql & " ,Sum([Liquor Tax]) as LTax, Sum([Deposits Collected])as TCollected, Sum([House Charges]) as HCharge"
 Sql = Sql & " ,Sum([Tip Charge Back]) as TCB"
 Sql = Sql & " from [Sales] Where "
 Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SearchDates.Item(1).EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 ConnectADORS (Sql), Db, RS
 If Not RS.EOF Then
      OverSales = Val(RS![ESale] & "") - Val(RS![TGrat] & "")
      If OverSales < 0 Then OverSales = 0
      Revenus = Val(Check_Null(RS![TTender])) + Val(Check_Null(RS![TPaidout])) + Val(Check_Null(RS![TDeposits])) - ((Val(RS![TCollected] & "") + Val(RS![HCharge] & "")))
      Collected = Val(Check_Null(RS![TNet])) + Val(Check_Null(RS![STAX])) + Val(Check_Null(RS![SPTax])) + Val(Check_Null(RS![LTax])) + OverSales + Val(RS![TCB] & "")
 End If
 DbClose Db, RS
 Data = Data & "POS Cash" & "|"
 Data = Data & Format(Revenus - Collected, "0.00") & ";"
 H = "|;"
 F = "+<2500|+>2000;"
 ReportView.View.TableBorder = tbBottom
 ReportView.View.addTable F, H, Data, vbBlack, vbWhite
 FinishJob 1
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Mix Reports")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub CATSales_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        Mix_History "", "Category", SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub CCounts_Click()
LoadDates "Select Distinct [Date] from [Sales] Where not isnull([date]) order by [Date] asc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
'On Error GoTo errorhandler
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim F       As String
Dim H       As String
Dim Data    As String
Dim TSales  As Currency
Dim TCust   As Long
 OpenDb SvrPath.TerminalPath & ClosingHistory, Db
   Sql = "Select [Name], Sum([Amount]) as TSold, Sum([Quantity]) as TCust from [Meal Periods] Where [Amount] > 0 and [Quantity] > 0 and TRY_CAST([date] AS datetime) between '" & Format(CDate(SearchDates.Item(1).StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
   Sql = Sql & " Group by [Name]"
 OpenTb (Sql), RS, Db
 If RS Is Nothing Then
    DbClose Db, RS
    Exit Sub
 End If
 F = "+2500|+1000|+>1500|+>1500;"
 H = "Meal Name|Qty|Amount|Average;"
 Data = "": TCust = 0: TSales = 0
 Do While Not RS.EOF
   Data = Data & RS.fields("Name").Value & "|"
   Data = Data & Format(Val(RS![TCust] & ""), "##") & "|"
   Data = Data & Format(Val(RS![TSold] & ""), "0.00") & "|"
   Data = Data & Format(Val(RS![TSold] & "") / Val(RS![TCust] & ""), "0.00") & ";"
   TCust = TCust + RS![TCust]
   TSales = TSales + RS![TSold]
  RS.MoveNext
 Loop
 ReportView.View.Orientation = orPortrait
 SetPage ReportView.View, "Meal Period/Customer " & InDate
 
 ReportView.View.StartTable
 ReportView.View.TableBorder = tbBoxColumns
 ReportView.View.addTable F, H, Data, vbGreen
 ReportView.View.EndTable
 
 Data = "Total" & "|"
 Data = Data & Format(TCust, "##") & "|"
 Data = Data & Format(TSales, "0.00") & "|"
 Data = Data & Format(TSales / TCust, "0.00") & ";"
 DbClose Db, RS
 ReportView.View.StartTable
 ReportView.View.TableBorder = tbBoxColumns
 ReportView.View.addTable F, "", Data, vbGreen
 ReportView.View.EndTable
 ReportView.View.EndDoc
 ReportView.Show 1
End Sub

Private Sub CgDate_Click()
If Can_Continue(CgDate.Caption) = False Then Exit Sub
Dim OldDate    As String
Dim NewDate    As String
If See_Msg("Select Date To Change", 4, "Enter Existing Date") = 6 Then
    With Review
       CenterObject Me, Review
       .ZOrder 0
       .Visible = True
       .Value = Format(Date, "mm/dd/yyyy")
    End With
    Do While Review.Visible = True
     DoEvents
    Loop
    If IsDate(SearchDate) Then
        OldDate = SearchDate
        If See_Msg("Select New Date to Move Date To", 4, "Enter New Date") = 6 Then
        With Review
           CenterObject Me, Review
           .ZOrder 0
           .Visible = True
           .Value = Date
        End With
        Do While Review.Visible = True
         DoEvents
        Loop
        If IsDate(SearchDate) Then
        NewDate = SearchDate
         If CloseDateFound(NewDate) = True Then
           See_Msg "Date Found Cannot Move To an Exsiting Date", 5, "Try a different Date"
           Exit Sub
         End If
         DateChange Format(OldDate, "MM/DD/YYYY"), Format(NewDate, "MM/DD/YYYY")
        End If
     End If
   End If
End If
Form_Load
End Sub


Private Sub CompareByTotal_Click()
SearchSpan = 0
If CompareByTotal.Text = "Week" Then
Compare.TextMatrix(0, 1) = SummarySheet.TextMatrix(0, 1)
SearchHow.Text = "Week"
Dim LR As Integer
   With Compare
    For LR = 1 To 21
         If SummarySheet.TextMatrix(LR, 8) = "" Then Exit For
         .TextMatrix(LR, 1) = SummarySheet.TextMatrix(LR, 8)
    Next LR
   End With
End If
End Sub

Private Sub CSBP_Click()
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
OpenDb SvrPath.TerminalPath & ClosingHistory, SL
Sql = "Select Distinct [Date] from [Sales] order by [date] asc "
ConnectADORS Sql, SL, SRd
Do While Not SRd.EOF
   DatePicker.StartingDate.AddItem SRd![Date] & ""
   DatePicker.EndingDate.AddItem SRd![Date] & ""
   SRd.MoveNext
Loop
DbClose SL, SRd
If DatePicker.StartingDate.ListCount > -1 Then
  DatePicker.Show 1
  If SearchDates Is Nothing Then

  Else
    LoadIndividualData "Category"
    PaymentType = "Category"
  End If
Else
  Call See_Msg("Not Dates Found", 5, "")
End If

End Sub

Private Sub DataMerge_Click()
Touch.FileName = ""
Touch.ShowOpen
If Touch.FileName = "" Then Exit Sub
If See_Msg("Merge Files", 4, "") <> 6 Then Exit Sub
Sleep 0
Dim SelFileName   As String
Dim splt() As String
splt = Split(Touch.FileName, "\")
 SelFileName = splt(UBound(splt))
FileCopy Touch.FileName, SvrPath.TerminalPath & "MF" & SelFileName
If BlendData(SvrPath.TerminalPath, "MF" & SelFileName, SvrPath.TerminalPath, SelFileName, False, "") Then
  See_Msg "Merge Finished....", 5, ""
Else
  See_Msg "Merge Failed....", 5, ""
  FileCopy SvrPath.TerminalPath & "MF" & SelFileName, Touch.FileName
End If

End Sub

Private Sub DaySum_Click()
With Review
  CenterObject Me, Review
  .ZOrder 0
  .Visible = True
  .Value = Format(Date, "mm/dd/yyyy")
End With
Do While Review.Visible = True
  DoEvents
  Sleep 100
Loop
If Not IsDate(SearchDate) Then Exit Sub
DailySummary SearchDate, False, 0

End Sub
 Public Sub SetBetweenString(StartDate As String)
        Dim Db As ADODB.Connection
        Dim RS As ADODB.Recordset
        Dim BDay As Integer
        Dim STime As String
        Dim ETime As String
        BusDay = 0
        If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
        Sql = "Select * from [AuditSales] Where TRY_CAST([Date] AS date) = '" & StartDate & "'"
        Sql = Sql & "Order by [start Time] asc"
        Call ConnectADORS(Sql, Db, RS)
        If Not RS.EOF Then
            BusDay = RS.fields("Close Out Day").Value
        End If
        RS.Close
        Sql = "Select * from [AuditSales] Where TRY_CAST([Date] AS date) Between '" & StartDate & "' and '" & DateAdd("d", 1, CDate(StartDate)) & "'"
        Sql = Sql & " and [close out day] = " & BusDay
        Sql = Sql & " Order by [Date], [start Time] asc"
        Call ConnectADORS(Sql, Db, RS)
        If Not RS.EOF Then
            STime = RS.fields("Start Time").Value
        End If
        Do While Not RS.EOF
            ETime = RS.fields("End Time").Value
            RS.MoveNext
        Loop
        Call DbClose(Db, RS)
        BusDayEndsAt = ETime
        If Not IsDate(CDate(Input_Option(84))) Then Input_Option(84) = "04:00 AM"  '// end of business day  -- default is 4:00 am
        If CDate(DateAdd("d", 1, CDate(StartDate)) & " " & CDate(Input_Option(84))) < CDate(DateAdd("d", 1, CDate(StartDate)) & " " & ETime) Then
          BusDayEndsAt = Input_Option(84)
        End If
        SqlBetweenStatment = "Between '" & CDate(StartDate & " " & STime) & "' and '" & CDate(DateAdd("d", 1, CDate(StartDate))) & " " & BusDayEndsAt & "'"
        AccessBetweenStatment = "Between '" & CDate(StartDate & " " & STime) & "' and '" & CDate(DateAdd("d", 1, CDate(StartDate))) & " " & BusDayEndsAt & "'"
    End Sub
Private Sub DEPSales_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        Mix_History "", "Department", SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
        Set SearchDates = Nothing
    End If
End If
End Sub

Private Sub DeptTnd_Click()
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    DepartmentTender SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End Sub

Private Sub Disc_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        DiscountHistory
        Set SearchDates = Nothing
    End If
End If
End Sub
Private Sub DMS_Click()
        Sql = "Select Distinct [Name], sum([Amount]) as TSold, Sum([quantity]) as QSold"
        Sql = Sql & " From [Department Sales] Where TRY_CAST([date] AS datetime) between '" & Format(CDate(SearchDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SearchDate) + 6), "yyyy-mm-dd hh:nn:ss") & "'"
        Sql = Sql & " Group By [Name]"
        Weeklytotals Sql
End Sub
Private Sub Dps_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        If SearchDates.Item(1).StartDate <> SearchDates.Item(1).EndDate Then
         Sql = "Select Distinct Sum([Amount]) as [Price] from [Deposits]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
        Else
         Sql = "Select [Amount] as [Price] from [Deposits]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
        End If
        HistoryReport "Deposit Report", Sql
        Set SearchDates = Nothing
    End If
End If
End Sub

Private Sub drp_Click()
End Sub

Private Sub DRS_Click()
End Sub

Private Sub DSRBalance_Click()
LoadDates "Select Distinct [Date] from [Sales] order by [Date] asc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
DSRBalanceSheet SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, False, 1
End Sub
Private Sub editdata_Click()
End Sub
Private Sub DSUMREPORT_Click()
End Sub
Private Sub EditDeposit_Click()
If Can_Continue(EditDeposit.Caption) = False Then Exit Sub
With Review
   CenterObject Me, Review
   .ZOrder 0
   .Visible = True
   .Value = Date
End With
Do While Review.Visible = True
 DoEvents
Loop
EditItem = "Deposits"
EditType = "Deposit"
If IsDate(SearchDate) Then
  EditDayData SearchDate
Else
  See_Msg "Invalid Date", 5, ""
End If
End Sub
Private Sub EditHistory_Click()
If Can_Continue(EditHistory.Caption) = False Then Exit Sub
With Review
   CenterObject Me, Review
   .ZOrder 0
   .Visible = True
   .Value = Date
End With
Do While Review.Visible = True
 DoEvents
Loop
EditItem = "Sales"
If IsDate(SearchDate) Then
  EditData SearchDate
Else
  See_Msg "Invalid Date", 5, ""
End If
End Sub
Private Sub End_Click()
Unload Me
End Sub
Public Function GetDualPriceAmount(SDate As String, EDate As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetDualPriceAmount = 0
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
Sql = "SELECT [Media], SUM([Amount]) AS Amount FROM [CheckPayments] "
Sql = Sql & "WHERE [Media] <> 'Cash' AND "
Sql = Sql & "TRY_CAST([Date] AS datetime) BETWEEN '" & Format(CDate(SDate & " " & GetStartTime), "yyyy-mm-dd hh:nn:ss") & "' "
Sql = Sql & "AND '" & Format(CDate(SetEndingDate(EDate) & " " & SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "' "
Sql = Sql & "GROUP BY [Media]"

ConnectADORS Sql, Db, RS
Do While Not RS.EOF
 GetDualPriceAmount = GetDualPriceAmount + Val(RS![Amount])
 RS.MoveNext
Loop
DbClose Db, RS
End Function

Public Function GetDiscountDepartment(DeptName As String, SDate As String, EDate As String, Db As ADODB.Connection) As Currency
'Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetDiscountDepartment = 0
If GetConfigInfo("Balance Sheet", "Net Department Totals", "Yes", "FUNCTIONS.INI") = "DISABLED" Then Exit Function
'ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
Sql = "Select [Department Name], sum([Item Price]) as TCoupons from [Detail] "
Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & CDate(SDate & " " & GetStartTime) & "'"
Sql = Sql & " and '" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "'"
Sql = Sql & " and [Item Price] < 0"
Sql = Sql & " and [Department Name] = '" & DeptName & "'"
Sql = Sql & " Group by [Department Name] "
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
 GetDiscountDepartment = Format(Val(RS![TCoupons] & ""), "0.00")
End If
RS.Close
'DbClose Db, RS
End Function
Private Function GetBankDeposits(DepData() As String, SDate As String, EDate As String) As Boolean
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetBankDeposits = False
Dim R  As Integer
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select * from [Deposits] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
R = 0
If Not RS.EOF Then
 GetBankDeposits = True
 R = RecordsCount(RS)
 ReDim DepData(R)
 R = 0
 Do While Not RS.EOF
   DepData(R) = Format(Val(RS![Amount] & ""), "0.00")
  R = R + 1
  RS.MoveNext
 Loop
End If
DbClose Db, RS
End Function
Public Function GetSafeAdjustments(DepData() As String, SDate As String, EDate As String) As Boolean
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
GetSafeAdjustments = False
'If Input_Option(45) = "Yes" Then Exit Function
Dim R  As Integer
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
 Sql = "Select * from [Non Sales Items] Where [Name] = 'Safe Amount' and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
 R = 0
If Not RS.EOF Then
 GetSafeAdjustments = True
 R = RecordsCount(RS)
 ReDim DepData(R)
 R = 0
 Do While Not RS.EOF
   DepData(R) = Format(Val(RS![Amount] & ""), "0.00")
  R = R + 1
  RS.MoveNext
 Loop
End If
DbClose Db, RS
End Function
Private Function NonSalesRevenue(InDate As String, EDate As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim TotalNSR As Currency
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select [Name], Sum([Amount]) as [Price] from [Department Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(CDate(InDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and [Amount] > 0"
Sql = Sql & " Group By [Name] "
ConnectADORS (Sql), Db, RS
TotalNSR = 0
Do While Not RS.EOF
  Select Case RS.fields("Name").Value
     Case Is = "Gift Certificate", "Gratuity", "Service Charge", "Paid In", "Gift EPay New Issue", "Gift EPay ReLoad"
          TotalNSR = TotalNSR + RS.fields("Price").Value
     Case Is = "Account Payment", "Account Deposit", "Non Sales Discount"
          TotalNSR = TotalNSR + RS.fields("Price").Value
  End Select
  RS.MoveNext
Loop
DbClose Db, RS
NonSalesRevenue = TotalNSR
End Function
Private Function GetGrossSales(InDate As String, EDate As String) As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim TotalNSR As Currency
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select [Name], Sum([Amount]) as [Price] from [Department Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and [Amount] > 0"
Sql = Sql & " Group by [Name]"
ConnectADORS (Sql), Db, RS
TotalNSR = 0
Do While Not RS.EOF
  TotalNSR = TotalNSR + RS.fields("Price").Value
  RS.MoveNext
Loop
DbClose Db, RS
GetGrossSales = TotalNSR
End Function
Private Sub SetSalesData(InDate As String, EDate As String, NonSale As Currency, GSales As Currency)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
'If NonSale = 0 Then Exit Sub
If GSales = 0 Then Exit Sub
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select * from [Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(InDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
Dim RC As Boolean
RC = False
If Not RS.EOF Then
   If NonSale > 0 Then
    If RS.fields("Exempt Sales").Value <> NonSale Then
      RS.fields("Exempt Sales").Value = NonSale
      RC = True
      RS.Update
    End If
   End If
   If RS.fields("Gross Sales").Value <> GSales + RS.fields("Voids").Value Then
      RS.fields("Gross Sales").Value = GSales + RS.fields("Voids").Value
      RC = True
      RS.Update
   End If
   If RC Then
    RS.fields("Net Sales").Value = ((RS.fields("Gross Sales").Value - RS.fields("Voids").Value) - RS.fields("Discounts").Value) - (RS.fields("Exempt Sales").Value + Abs(RS.fields("Refunds").Value))
    RS.Update
   End If
End If
DbClose Db, RS
End Sub
Public Sub DSRBalanceSheet(InDate As String, EDate As String, EmailReport As Boolean, PBS As Integer)
On Error GoTo ErrorHandler
Dim RwPos As Integer
Dim Col1Row As Integer
Dim Col2Row As Integer
Dim Col1Total As Currency
Dim Col2Total As Currency
Dim Collected As Currency
Dim Revenus   As Currency
Dim NetSales  As Currency
Dim SubTotal  As Currency
Dim GrossS    As Currency
Dim NetS      As Currency
Dim Voids     As Currency
Dim Grat      As Currency
Dim DiscDept  As Currency
Dim DeptDisc  As Currency
Dim SalesTax  As Currency
Dim BankDep()   As String
Dim R           As Integer
Dim TipFee    As Currency
Dim TotalTipHold As Currency
Dim TipWithheld  As Currency
Dim GratHeld     As Currency
Dim DepartmentDicount As Currency
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim PDB As ADODB.Connection
Dim Prs As ADODB.Recordset
Dim NonCashTenders As Currency
Dim DiscountTotal  As Currency
Dim NonTaxedSales  As Currency
Dim VoidsOnly      As Currency
Dim Discrepency    As Currency
Dim NetofDiscount  As Currency
Dim CreditOverCash As Currency
Dim ProcessingFee  As Currency
Dim PaidOuts       As Currency
Dim BankDrop       As Currency
Dim GiftCardSales  As Currency
With ShowEdit
    .Clear
    .rows = 500
    .cols = 5
    .FixedRows = 1
    .FixedCols = 0
    .BorderStyle = flexBorderFlat
    .GridLines = flexGridNone
    .GridLinesFixed = flexGridNone
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColWidth(0) = 2500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 2500
    .ColWidth(4) = 1600
    .width = 10200
    .Height = 7200
    .FontName = arial
    .FontSize = 8
End With

If Not IsDate(InDate) Then
  See_Msg "Invalid Starting Date", 5, ""
  Exit Sub
End If
If Not IsDate(EDate) Then
  See_Msg "Invalid Ending Date", 5, ""
  Exit Sub
End If
If GetConfigInfo("Balance Sheet", "Fix Net Sales", "No", "Functions.INI") = "ENABLED" Then
 If CDate(InDate) = CDate(EDate) Then
  Dim NonSalesRev    As Currency
  Dim NetDeptSales   As Currency
  NetDeptSales = 0
  NetDeptSales = 0
  NetofDiscount = 0
  NonSalesRev = NonSalesRevenue(InDate, InDate)
  NetDeptSales = GetGrossSales(InDate, InDate)
  SetSalesData InDate, InDate, NonSalesRev, NetDeptSales
 End If
End If

CenterObject Me, ShowEdit
ShowEdit.Visible = True
With ShowEdit
   .TextMatrix(0, 0) = "Total Revenue"
   .TextMatrix(0, 1) = Format(InDate, "MM/DD/YY")
   .TextMatrix(0, 3) = "Total Accounted"
   .TextMatrix(0, 2) = " to " & Format(EDate, "MM/DD/YY")
End With
Col1Total = 0
Dim DptTotal As Currency
Dim MealCount As MealCounts
MealCount = ReturnMealCount(InDate, EDate)

ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
'///////////////////////////
Sql = "Select [Department Name], sum([Item Price]) as Price from [Detail] "
Sql = Sql & " Where [Date] Between '" & CDate(InDate & " " & GetStartTime) & "'"
Sql = Sql & " and '" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "'"
Sql = Sql & " and [Department Name] Not In ('Gift Certificate','Gratuity','Service Charge','Paid In','Gift EPay New Issue','Gift EPay ReLoad','Account Payment','Account Deposit','Non Sales Discount','Discount','Coupon','Online Tip','Frequent Diner','Donation','Online Fee','Online Delivery','Item Discount')"
Sql = Sql & " Group By [Department Name] order by [department Name] asc "
ConnectADORS Sql, Db, RS
Col1Row = 0
DptTotal = 0
With ShowEdit
Col1Row = 1
.FontBold = True
.CellFontBold = True
.TextMatrix(Col1Row, 0) = "Departments Sales"
.TextMatrix(Col1Row, 1) = "         Totals"
.CellFontBold = False
DiscDept = 0: DepartmentDicount = 0: NetofDiscount = 0
Do While Not RS.EOF
     If Val(RS![Price] & "") <> 0 Then
       Col1Row = Col1Row + 1
       DeptDisc = GetDiscountDepartment(RS![Department Name] & "", InDate, EDate, Db)
       If Abs(DeptDisc) > RS.fields("Price").Value Then
         DeptDisc = 0
       End If
       NetofDiscount = NetofDiscount + Abs(DeptDisc)
       .TextMatrix(Col1Row, 0) = "  " & RS![Department Name] & ""
       .TextMatrix(Col1Row, 1) = Format(Val(RS![Price] & "") + Abs(DeptDisc), "0.00")
       If Val(RS![Price] & "") > 0 Then
         DptTotal = DptTotal + (Val(RS![Price] & ""))
       End If
       If DeptDisc <> 0 Then
        Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = " ** " & RS![Department Name] & " Discount"
        .TextMatrix(Col1Row, 1) = Format(Val(DeptDisc), "0.00")
        DepartmentDicount = DepartmentDicount + Format(Val(DeptDisc), "0.00")
        Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = " ** " & RS![Department Name] & " Net"
        .TextMatrix(Col1Row, 1) = Format(Val(RS![Price] & ""), "0.00")
       End If
     End If
   RS.MoveNext
Loop
Col1Row = Col1Row + 1
If NetofDiscount = 0 Then
   .TextMatrix(Col1Row, 0) = "Total  *(includes discounts)*"
Else
   .TextMatrix(Col1Row, 0) = "Total  *(Net of discounts)*"
End If
.TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
 Col1Total = DptTotal
End With
DbClose Db, RS
'NetSales = DptTotal
    ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
    Sql = "Select [Menu Item], sum([Item Price]) as TExempt from [Detail] "
    Sql = Sql & " Where [Department Name] IN ('Paid In','Donation')"
    Sql = Sql & " and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(InDate & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(EDate) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " Group by [Menu Item] "
    ConnectADORS ReturnDateString(Sql), Db, RS
    If Not RS.EOF Then
        DptTotal = 0
        With ShowEdit
         Col1Row = Col1Row + 1
         Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = "Non Sales Revenue"
        Do While Not RS.EOF
              Col1Row = Col1Row + 1
              .TextMatrix(Col1Row, 0) = "  " & RS![Menu Item] & ""
              .TextMatrix(Col1Row, 1) = Format(Val(RS![TExempt] & ""), "0.00")
              DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
           RS.MoveNext
        Loop
              Col1Row = Col1Row + 1
              .TextMatrix(Col1Row, 0) = "Total"
              .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
              Col1Total = Col1Total + DptTotal
        End With
     End If
     VoidsOnly = 0
     If Check_Option(244) = 1 Then
        RS.Close
        Sql = "Select [Item Name], sum([Amount]) as TVoids From [Voids]"
        Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(InDate & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
        Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(EDate) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
        Sql = Sql & " and [Reason] Not In ('Cashier Correct','Manager Correct','Cancel Order')"
        Sql = Sql & " Group by [Item Name]"
       ConnectADORS ReturnDateString(Sql), Db, RS
         Do While Not RS.EOF
           VoidsOnly = VoidsOnly + RS![TVoids]
           RS.MoveNext
         Loop
     End If
DbClose Db, RS
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select [Name], Sum([Amount]) as [Price] from [Department Sales] "
Sql = Sql & " Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and [Name] In ('Mercury Gift','Gift Certificate','Gift EPay New Issue','Gift EPay ReLoad','Account Payment','Account Deposit','Non Sales Discount','Donation')"
Sql = Sql & " and TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " Group By [Name] "
GiftCardSales = 0
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
    DptTotal = 0
    With ShowEdit
     Col1Row = Col1Row + 1
     Col1Row = Col1Row + 1
    .TextMatrix(Col1Row, 0) = "Gift Certificate Sales"
    Do While Not RS.EOF
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "  " & RS![Name] & ""
          .TextMatrix(Col1Row, 1) = Format(Val(RS![Price] & ""), "0.00")
          DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
       RS.MoveNext
    Loop
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "Total"
          .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
          Col1Total = Col1Total + DptTotal
    End With
    GiftCardSales = DptTotal
 ' End If
End If
SetNothing RS
Sql = "Select [Name], Sum([Amount]) as [Price] from [Department Sales] "
Sql = Sql & " Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and [Name] = 'Service Charge'"
Sql = Sql & " Group By [Name] "
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
  DptTotal = 0
  With ShowEdit
     Col1Row = Col1Row + 1
     Col1Row = Col1Row + 1
    .TextMatrix(Col1Row, 0) = "Service Charges/Fees"
    Do While Not RS.EOF
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "  " & RS![Name] & ""
          '.TextMatrix(Col1Row, 1) = Format(Val(RS![Price] & ""), "0.00")
          DptTotal = DptTotal + Format(Val(RS![Price] & ""), "0.00")   ' Val(.TextMatrix(Col1Row, 1) & "")
       RS.MoveNext
    Loop
SetNothing RS
ConnectADODB SvrPath.TerminalPath & ProductHistory, PDB
   Sql = "Select [Menu Item], sum([Item Price]) as TExempt from [Detail] "
   Sql = Sql & " Where [Department Name] = 'Service Charge'"
   Sql = Sql & " and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(InDate & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
   Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(EDate) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
   Sql = Sql & " Group by [Menu Item] "
ConnectADORS ReturnDateString(Sql), PDB, Prs
   If Not Prs.EOF Then
       Do While Not Prs.EOF
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = " ** " & Prs![Menu Item] & ""
          .TextMatrix(Col1Row, 1) = Format(Val(Prs![TExempt] & ""), "0.00")
          'DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
          Prs.MoveNext
       Loop
  End If
  DbClose PDB, Prs
'*************************************************************************
  Col1Row = Col1Row + 1
  .TextMatrix(Col1Row, 0) = "Total"
  .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
  Col1Total = Col1Total + DptTotal
 End With
End If

SetNothing RS
'ConnectADODB SvrPath.TerminalPath & ClosingHistory, DB
Sql = "Select sum([Gross Sales]) as GSales, Sum([Net Sales]) as NSales, Sum([Sales Tax]) as STax, Sum([special Tax]) as SpTax,sum([Liquor Tax]) as LTax, sum([Voids]) as VDS,sum([Charged Gratuity]) as Grat, sum([None Taxed Sales]) as NonTaxedSales from [Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
    GrossS = Val(RS![GSales] & "")
    NetS = Val(RS![NSales] & "")
    Voids = Val(RS![Vds] & "")
    Grat = Val(RS![Grat] & "")
    NonTaxedSales = Val(RS![NonTaxedSales] & "")
    DptTotal = Val(RS![STAX] & "") + Val(RS![SPTax] & "") + Val(RS![LTax] & "")
    'SetNothing RS
    With ShowEdit
     Col1Row = Col1Row + 1
     Col1Row = Col1Row + 1
     .TextMatrix(Col1Row, 0) = "Sales Tax Collected"
        Col1Row = Col1Row + 1
       .TextMatrix(Col1Row, 0) = "  " & GetTextFromINI("Sales Tax 1")
       .TextMatrix(Col1Row, 1) = Format(Val(RS![STAX] & ""), "0.00")
       If Val(RS![SPTax] & "") <> 0 Then
         Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = "  " & GetTextFromINI("Sales Tax 2")
        .TextMatrix(Col1Row, 1) = Format(Val(RS![SPTax] & ""), "0.00")
       End If
       If Val(RS![LTax] & "") <> 0 Then
        Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = "  " & GetTextFromINI("Sur Charge")
        .TextMatrix(Col1Row, 1) = Format(Val(RS![LTax] & ""), "0.00")
       End If
       'DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
       'End If
       Col1Row = Col1Row + 1
       .TextMatrix(Col1Row, 0) = "Total"
       .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
       Col1Total = Col1Total + DptTotal
    End With
End If
SalesTax = DptTotal
DbClose Db, RS
TipWithheld = 0: TotalTipHold = 0
  ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    Sql = "Select [Job Code], Sum([Tips Withheld]) as TipsWithheld from [Tipout Withheld] Where TRY_CAST([Date] AS date) between '" & InDate & "'"
    Sql = Sql & " and '" & CDate(EDate) & "'"
    Sql = Sql & " Group By [Job Code]"
    ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
       If Val(RS![TipsWithheld] & "") <> 0 Then
        DptTotal = 0
        With ShowEdit
         Col1Row = Col1Row + 1
         Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = "Tip Share Collected"
        Do While Not RS.EOF
              Col1Row = Col1Row + 1
              .TextMatrix(Col1Row, 0) = "  " & RS![Job Code] & ""
              .TextMatrix(Col1Row, 1) = Format(Val(RS![TipsWithheld] & ""), "0.00")
              DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
           RS.MoveNext
        Loop
        Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "** Total is reference only **"
          .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
          TipWithheld = DptTotal
          DptTotal = 0
        End With
       End If
     End If
  DbClose Db, RS
GratHeld = 0
If Input_Option(58) = "Yes" Then
    ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
    Sql = "Select [Name], Sum([Amount]) as [Price] from [Department Sales] "
    Sql = Sql & " Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " and [Name] = 'Gratuity'"
    Sql = Sql & " Group By [Name] "
    ConnectADORS (Sql), Db, RS
    If Not RS.EOF Then
    ' If RS![Name] & "" = "Gift Certificate" Or RS![Name] & "" = "Mercury Gift" Then
       DptTotal = 0
        With ShowEdit
         Col1Row = Col1Row + 1
         Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = "Gratuity Collected"
        Do While Not RS.EOF
              Col1Row = Col1Row + 1
              .TextMatrix(Col1Row, 0) = "  " & RS![Name] & ""
              .TextMatrix(Col1Row, 1) = Format(Val(RS![Price] & ""), "0.00")
              DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
           RS.MoveNext
        Loop
              Col1Row = Col1Row + 1
              .TextMatrix(Col1Row, 0) = "Total"
              .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
              Col1Total = Col1Total + DptTotal
        End With
        GratHeld = DptTotal
     ' End If
    End If
    DbClose Db, RS
End If

If Check_Option(237) = 1 Then 'tips are paid on payrolld
  ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    Sql = "Select [Register Group], Sum([Tips Withheld]) as TTips from [Tips Withheld] Where TRY_CAST([Date] AS date) between '" & InDate & "'"
    Sql = Sql & " and '" & CDate(EDate) & "'"
    Sql = Sql & " Group By [Register Group]"
    ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
       If Val(RS![TTips] & "") <> 0 Then
        DptTotal = 0
        With ShowEdit
         Col1Row = Col1Row + 1
         Col1Row = Col1Row + 1
        .TextMatrix(Col1Row, 0) = "Tips Withheld"
        Do While Not RS.EOF
              Col1Row = Col1Row + 1
              .TextMatrix(Col1Row, 0) = "  " & RS![Register Group] & ""
              .TextMatrix(Col1Row, 1) = Format(Val(RS![TTips] & ""), "0.00")
              DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
           RS.MoveNext
        Loop
        Col1Row = Col1Row + 1
        If GetConfigInfo("Misc", "Tip Out Is Reference", "No", "Functions.INI") = "ENABLED" Then
         .TextMatrix(Col1Row, 0) = "Total (Reference Only)"
         .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
        Else
         .TextMatrix(Col1Row, 0) = "Total"
         .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
         Col1Total = Col1Total + DptTotal
         TotalTipHold = DptTotal
         TotalTipHold = (TotalTipHold - TipWithheld)
         'TipWithheld = 0
        End If
        DptTotal = 0
        End With
       End If
     End If
  DbClose Db, RS
End If
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
'***************************************************************************
Sql = "Select Sum([Tip Charge Back]) as TipChargeBack from [Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
  If Val(RS![TipChargeBack] & "") > 0 Then
    DptTotal = 0
    With ShowEdit
     Col1Row = Col1Row + 1
     Col1Row = Col1Row + 1
    .TextMatrix(Col1Row, 0) = "C/C Fee Income"
    Do While Not RS.EOF
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "  Tip Fee Income"
          .TextMatrix(Col1Row, 1) = Format(Val(RS![TipChargeBack] & ""), "0.00")
          DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
       RS.MoveNext
    Loop
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "Total"
          .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
          Col1Total = Col1Total + DptTotal
          TipFee = DptTotal
    End With
   End If
End If
SetNothing RS
'***************************************************************************
Sql = "Select Sum([Refunds]) as TotalRefunds from [Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
Dim RFT As Currency
'// Refund total
If Not RS.EOF Then
  RFT = Val(RS![TotalRefunds] & "") * -1
  '// set refunds to a negitive amount
End If
SetNothing RS
Sql = "Select * from [Refunds] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If RFT <> 0 Then  '// will not display if no refunds are present
    DptTotal = 0: RFT = 0
    With ShowEdit
     Col1Row = Col1Row + 1
     Col1Row = Col1Row + 1
    .TextMatrix(Col1Row, 0) = "Refunds"
    Do While Not RS.EOF
       Col1Row = Col1Row + 1
       .TextMatrix(Col1Row, 0) = "  " & RS![Name] & ""  'Customer Refunds"
       .TextMatrix(Col1Row, 1) = Format(Abs(Val(RS![Amount] & "")) * -1, "0.00")
       RFT = RFT + Val(.TextMatrix(Col1Row, 1) & "")
       RS.MoveNext
    Loop
    Col1Row = Col1Row + 1
    .TextMatrix(Col1Row, 0) = "Total"
    .TextMatrix(Col1Row, 1) = Format(RFT, CurrencyPreFix)
    Col1Total = Col1Total + RFT
   ' DptTotal = DptTotal + RFT
    End With
End If
SetNothing RS
Sql = "Select Sum([Deposits Collected]) as PartyDeposits from [Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
  If RS![PartyDeposits] > 0 Then
    DptTotal = 0
    With ShowEdit
     Col1Row = Col1Row + 1
     Col1Row = Col1Row + 1
    .TextMatrix(Col1Row, 0) = "Deposits Paid In"
    Do While Not RS.EOF
          Col1Row = Col1Row + 1
          .TextMatrix(Col1Row, 0) = "  Deposits Collected"
          .TextMatrix(Col1Row, 1) = Format(Abs(Val(RS![PartyDeposits] & "")), "0.00")
          DptTotal = DptTotal + Val(.TextMatrix(Col1Row, 1) & "")
       RS.MoveNext
    Loop
    Col1Row = Col1Row + 1
     .TextMatrix(Col1Row, 0) = "Total"
     .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
    Col1Total = Col1Total + DptTotal
    End With
  End If
End If
SetNothing RS
DbClose Db, RS
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db

'********************col 2

Sql = "Select [Media], Sum([Amount]) as TAmount,Sum([Tip]) as TTip from [CheckPayments]"
Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & CDate(InDate & " " & GetStartTime) & "'"
Sql = Sql & " and '" & CDate(SetEndingDate(EDate) & " " & SetEndingTime) & "'"
Sql = Sql & " Group By [Media]"
ConnectADORS (Sql), Db, RS
'// sum the total amount of tenders
DptTotal = 0
Col2Total = 0
SubTotal = 0
With ShowEdit
Col2Row = 1
    .TextMatrix(Col2Row, 3) = "Credit Cards"
    .TextMatrix(Col2Row, 4) = "        Totals"
    Do While Not RS.EOF
     If RS![Media] & "" <> "Cash" Then
      If Val(RS![TAmount] & "") <> 0 Then
        Select Case RS![Media] & ""
         Case Is = "Visa", Trim$("MasterCard"), "Amex", "Discover", "Diners", Trim$("Master Card")
          Col2Row = Col2Row + 1
          .TextMatrix(Col2Row, 3) = "  " & RS![Media]
          .TextMatrix(Col2Row, 4) = Format(Val(RS![TAmount] & "") + RS.fields("TTip").Value, "0.00")
          DptTotal = DptTotal + Val(RS![TAmount] & "") + RS.fields("TTip").Value
          SubTotal = SubTotal + Val(RS![TAmount] & "") + RS.fields("TTip").Value
        End Select
      End If
     End If
     RS.MoveNext
    Loop
    Col2Row = Col2Row + 1
    .TextMatrix(Col2Row, 3) = " Sub Total"
    .TextMatrix(Col2Row, 4) = Format(SubTotal, CurrencyPreFix)
    Col2Total = DptTotal
    '*********** add sub total
    Col2Row = Col2Row + 1
    .TextMatrix(Col2Row, 3) = "Non Cash Tenders"
    .TextMatrix(Col2Row, 4) = "        "
    SubTotal = 0
    If Not RS.BOF Then RS.MoveFirst
    Do While Not RS.EOF
      If RS![Media] & "" <> "Cash" Then
       If Val(RS![TAmount] & "") <> 0 Then
         Select Case Trim$(RS![Media] & "")
         Case Is = "Visa", Trim$("MasterCard"), "Amex", "Discover", "Diners", Trim$("Master Card")
          
         Case Else
           Col2Row = Col2Row + 1
           .TextMatrix(Col2Row, 3) = "  " & RS![Media] & ""
           .TextMatrix(Col2Row, 4) = Format(Val(RS![TAmount] & "") + RS.fields("TTip").Value, "0.00")
           DptTotal = DptTotal + Val(RS![TAmount] & "") + RS.fields("TTip").Value
           SubTotal = SubTotal + Val(RS![TAmount] & "") + RS.fields("TTip").Value
         End Select
       End If
      End If
      RS.MoveNext
    Loop
          Col2Row = Col2Row + 1
          .TextMatrix(Col2Row, 3) = " Sub Total"
          .TextMatrix(Col2Row, 4) = Format(SubTotal, CurrencyPreFix)
          Col2Total = DptTotal
          Col2Row = Col2Row + 1
          .TextMatrix(Col2Row, 3) = "Total"
          .TextMatrix(Col2Row, 4) = Format(DptTotal, CurrencyPreFix)
          Col2Total = DptTotal
          NonCashTenders = Col2Total
''*************************
'End If
End With
DbClose Db, RS
'SetNothing RS
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db

Col2Row = Col2Row + 1
DptTotal = 0
 Sql = "Select [Name], Sum([Amount]) as TAmount from [Paid Outs] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " Group by [Name]"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
 With ShowEdit
 Do While Not RS.EOF
     DptTotal = DptTotal + Val(RS![TAmount] & "")
   RS.MoveNext
 Loop
 Col2Row = Col2Row + 1
  .TextMatrix(Col2Row, 3) = "Paid Outs"
  .TextMatrix(Col2Row, 4) = Format(DptTotal, CurrencyPreFix)
 Col2Total = Col2Total + DptTotal
End With
PaidOuts = DptTotal
End If
SetNothing RS
Sql = "Select [Name], Sum([Amount]) as TAmount from [Discounts] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " Group by [Name]"
ConnectADORS (Sql), Db, RS
DptTotal = 0
If Not RS.EOF Then
 With ShowEdit
 Col2Row = Col2Row + 1
  Do While Not RS.EOF
   DptTotal = DptTotal + Val(RS![TAmount] & "")
   RS.MoveNext
  Loop
  Col2Row = Col2Row + 1
   .TextMatrix(Col2Row, 3) = "Discounts"
   .TextMatrix(Col2Row, 4) = Format(DptTotal, CurrencyPreFix)
   If DepartmentDicount <> 0 Then
    Col2Row = Col2Row + 1
    .TextMatrix(Col2Row, 3) = "Less Dept Disc"
    .TextMatrix(Col2Row, 4) = Format(DptTotal - Abs(DepartmentDicount), CurrencyPreFix)
   End If
    Col2Total = Col2Total + (DptTotal - Abs(DepartmentDicount))
   DiscDept = (DptTotal - Abs(DepartmentDicount))
 End With
End If
DiscountTotal = DptTotal
SetNothing RS
 Sql = "Select  Sum([Deposits]) as TAmount from [Sales] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
 DptTotal = 0
 Col2Row = Col2Row + 1
 With ShowEdit
 Do While Not RS.EOF
   DptTotal = DptTotal + Val(RS![TAmount] & "")
   RS.MoveNext
 Loop
 Col2Row = Col2Row + 1
  .TextMatrix(Col2Row, 3) = "Bank Deposits"
  .TextMatrix(Col2Row, 4) = Format(DptTotal, CurrencyPreFix)
 Col2Total = Col2Total + DptTotal
  BankDrop = 0
 If DateDiff("d", InDate, EDate) < 8 Then
   If GetBankDeposits(BankDep(), InDate, EDate) Then
     For R = 0 To UBound(BankDep, 1) - 1
      Col2Row = Col2Row + 1
      .TextMatrix(Col2Row, 3) = "** Drops ( " & Format(BankDep(R), CurrencyPreFix) & " )"
      BankDrop = BankDrop + BankDep(R)
     Next R
   End If
 End If
 If DateDiff("d", InDate, EDate) < 8 Then
   If GetSafeAdjustments(BankDep(), InDate, EDate) Then
    For R = 0 To UBound(BankDep, 1) - 1
     Col2Row = Col2Row + 1
     .TextMatrix(Col2Row, 3) = "* Safe Amt ( " & Format(BankDep(R), CurrencyPreFix) & " )"
    Next R
   End If
 End If
End With
End If
SetNothing RS
Sql = "Select Sum([Exempt Sales]) as ESale, Sum([Charged Gratuity]) as TGrat"
Sql = Sql & " ,Sum([Tenders]) as TTender, Sum([Paid Outs]) as TPaidOut, Sum([Deposits]) as TDeposits"
Sql = Sql & " ,Sum([Net Sales]) as TNet ,Sum([Sales Tax]) as Stax, Sum([Special Tax]) as SpTax"
Sql = Sql & " ,Sum([Liquor Tax]) as LTax, Sum([Deposits Collected])as TCollected, Sum([House Charges]) as HCharge"
Sql = Sql & " ,Sum([Tip Charge Back]) as TCB"
Sql = Sql & " from [Sales] Where "
Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(InDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
'// TCollected = Deposits Collected
'// HCharge = House payments connected
If Not RS.EOF Then
DptTotal = 0
Col2Row = Col2Row + 1
Dim ReturnTotals As BalanceSheetValues
ReturnTotals = GetCashTenders(InDate, EDate)
ProcessingFee = ReturnProcessingFees(InDate, EDate)

With ShowEdit
      Col2Row = Col2Row + 1
      If GratHeld > 0 Then  'No adjustment for non sale reven
          OverSales = Val(RS![ESale] & "") + TotalTipHold
      Else
          OverSales = Val(RS![ESale] & "") - (Val(RS![TGrat] & "") + TotalTipHold)
      End If '// OverSales is the amount of non sales revenue collected
      If TotalTipHold = 0 And TipWithheld <> 0 Then
         If GratHeld > 0 Then
          OverSales = Val(RS![ESale] & "") - Abs(TipWithheld)
         Else
          OverSales = Val(RS![ESale] & "") - (Val(RS![TGrat] & "") + Abs(TipWithheld))
         End If
      End If
      If OverSales < 0 Then OverSales = 0
      Revenus = NonCashTenders + Val(RS![TPaidout] & "") + Val(RS![TDeposits] & "") - ((Val(RS![TCollected] & "") + Val(RS![HCharge] & "")))  'TotalTipHold
      'Revenus = NonCashTenders + Val(RS![TPaidout] & "") + Val(RS![TDeposits] & "") - (Val(RS![HCharge] & ""))  'TotalTipHold
      '// TCollected is deposits
      '// HCharge is the total payments pay to house accounts
      
      '// Credit Cards and gift cards  + paid outs + Deposits collected from phone orders less party deposit and house payments
      Collected = Val(RS![TNet] & "") + Val(RS![STAX] & "") + Val(RS![SPTax] & "") + Val(RS![LTax] & "") + OverSales + Val(RS![TCB] & "")
      'Collected = (NetS + GiftCardSales) + (SalesTax + TipFee) + ProcessingFee
      '// Net Sale + all taxes + non sales revenue
      '// TCB tip fees collected from server staff
      '// overSales is service charges non sales revenue etc.
      
      Discrepency = Format(Revenus - Collected, "0.00")
      DptTotal = DptTotal + (Discrepency) * -1
      Col2Total = Col2Total + DptTotal
      Dim Disdiff As Currency
      Disdiff = Col1Total - Col2Total
      'If Col2Total <> Col1Total Then
         '.TextMatrix(Col2Row, 4) = Format((Revenus - Collected) + Disdiff, "0.00")
       '  Trans_Log "Balance Sheet Information Data " & Format(Disdiff, "0.00")
      '   Col2Total = Col2Total + Disdiff
      '   .TextMatrix(Col2Row, 4) = Format(((Revenus + Abs(DiscDept) + Val(RS![TCollected] & ""))) - Col1Total, "0.00")
      'Else
         .TextMatrix(Col2Row, 4) = Format(Discrepency, "0.00") '    Format(((Revenus + Abs(DiscDept) + Val(RS![TCollected] & ""))) - Col1Total, "0.00")
      'End If
      If Val(.TextMatrix(Col2Row, 4) & "") < 0 Then
        .TextMatrix(Col2Row, 3) = "Short"
      ElseIf Val(.TextMatrix(Col2Row, 4) & "") > 0 Then
        .TextMatrix(Col2Row, 3) = "Over"
      Else
        .TextMatrix(Col2Row, 3) = "Perfect"
      End If
      
      If Check_Option(237) = 1 Then
       If BankDrop > 0 Then
        Col2Row = Col2Row + 1
        .TextMatrix(Col2Row, 3) = "Counted "
        .TextMatrix(Col2Row, 4) = Format((BankDrop), "0.00")   ' removed PaidOuts from this line
        
        If PaidOuts > 0 Then
        Col2Row = Col2Row + 1
        .TextMatrix(Col2Row, 3) = "Paid Outs "
        .TextMatrix(Col2Row, 4) = Format((PaidOuts), "0.00")   ' removed PaidOuts from this line
        End If
        If TotalTipHold > 0 Then
        Col2Row = Col2Row + 1
        .TextMatrix(Col2Row, 3) = "Tips Withheld "
        .TextMatrix(Col2Row, 4) = Format((TotalTipHold), "0.00")   ' removed PaidOuts from this line
        End If
       End If
       Col2Row = Col2Row + 1
        .TextMatrix(Col2Row, 3) = "Cash Tenders "
        .TextMatrix(Col2Row, 4) = Format((ReturnTotals.DayCashTenders), "0.00")   ' removed PaidOuts from this line
       Col2Row = Col2Row + 1
        .TextMatrix(Col2Row, 3) = "Difference "
        .TextMatrix(Col2Row, 4) = Format(ReturnTotals.DayCashTenders - ((BankDrop + PaidOuts) + TotalTipHold), "0.00") ' removed PaidOuts from this line
      End If
     'If GetConfigInfo("Balance Sheet", "Show Cash Over Tips", "No", "FUNCTIONS.INI") = "ENABLED" Then
      If ReturnTotals.DayCashTenders - (ReturnTotals.DayNoCashTips + Grat) < 0 Then
         CreditOverCash = TotalCashTenders - (ReturnTotals.DayNoCashTips + Grat)
         .TextMatrix(Col2Row, 3) = "Credit Tips Over Cash (" & Format(Abs(CreditOverCash), "0.00") & ")"
         .TextMatrix(Col2Row, 4) = Format(.TextMatrix(Col2Row, 4) - Abs(CreditOverCash), "0.00")
      End If
     'End If
End With
End If
SetNothing RS
Dim WhatRow As Integer
If Col1Row > Col2Row Then
  WhatRow = Col1Row
Else
  WhatRow = Col2Row
End If
Dim TTTL As Currency
Dim TCT  As Long
Dim CCT  As Long
Sql = "Select Sum([Amount]) as TSale, Sum([Quantity]) as TCt"
Sql = Sql & " from [Order Destination] Where "
Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(InDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
 TTTL = Val(RS.fields("TSale").Value & "")
 TCT = Val(RS.fields("TCT").Value & "")
Else
 TTTL = 0
 TCT = 0
End If
SetNothing RS
Sql = "Select Sum([Quantity]) as TCt"
Sql = Sql & " from [Meal Periods] Where "
Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(InDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
 CCT = Val(RS.fields("TCT").Value & "")
Else
 CCT = 0
End If
SetNothing RS
With ShowEdit
    WhatRow = WhatRow + 1
    WhatRow = WhatRow + 1
    .TextMatrix(WhatRow, 0) = "** ACCOUNT FOR"
    .TextMatrix(WhatRow, 1) = Format(Col1Total, CurrencyPreFix)
    .TextMatrix(WhatRow, 3) = "** ACCOUNTED FOR"
    .TextMatrix(WhatRow, 4) = Format(Col2Total, CurrencyPreFix)
End With
Col1Row = WhatRow + 2
Col2Row = WhatRow + 2
'Dim TotalCashTenders As Currency
'Dim TotalNonCashTips As Currency
'TotalCashTenders = GetCashTenders(InDate, EDate, TotalNonCashTips)
'ProcessingFee = ReturnProcessingFees(InDate, EDate)

With ShowEdit
    WhatRow = WhatRow + 2
    .TextMatrix(WhatRow, 0) = "** Gross Sales **"
    .TextMatrix(WhatRow, 1) = "** Net + Discounts **"
    .TextMatrix(WhatRow, 2) = "** Net Sales **"
    .TextMatrix(WhatRow, 3) = "** Voids **"
    .TextMatrix(WhatRow, 4) = "** Not Taxed"
    WhatRow = WhatRow + 1
    .TextMatrix(WhatRow, 0) = Format(GrossS, CurrencyPreFix)
    .TextMatrix(WhatRow, 1) = Format(NetS + DiscountTotal, CurrencyPreFix)
    .TextMatrix(WhatRow, 2) = Format(NetS, CurrencyPreFix)
    If VoidsOnly > 0 Then
      .TextMatrix(WhatRow, 3) = " E/C TL " & Format((Voids - VoidsOnly), "0.00") & " Void TL " & Format(VoidsOnly, CurrencyPreFix)
    Else
      .TextMatrix(WhatRow, 3) = Format(Voids, CurrencyPreFix)
    End If
    If NonTaxedSales > 0 Then
     .TextMatrix(WhatRow, 4) = Format(NonTaxedSales, "0.00")
    Else
     .TextMatrix(WhatRow, 4) = "0.00"
    End If
    WhatRow = WhatRow + 1
    If ProcessingFee = 0 Then
    .TextMatrix(WhatRow, 0) = "** Grand Total     "
    Else
    .TextMatrix(WhatRow, 0) = "** Grand Total (fees included)"
    End If

    .TextMatrix(WhatRow, 1) = "** Cash Tenders    "
    .TextMatrix(WhatRow, 2) = "** Non Cash Tips   "
    .TextMatrix(WhatRow, 3) = "** Gratuity **"
    'If Check_Option(237) = 1 Then
    ' .TextMatrix(WhatRow, 4) = "** Cash Tenders  "
    'Else
     .TextMatrix(WhatRow, 4) = "** Cash Less Tips  "
    'End If
    WhatRow = WhatRow + 1
   
    .TextMatrix(WhatRow, 0) = Format((NetS + GiftCardSales) + (SalesTax + TipFee) + ProcessingFee, CurrencyPreFix)
    .TextMatrix(WhatRow, 1) = Format(ReturnTotals.DayCashTenders, "0.00")
    .TextMatrix(WhatRow, 2) = Format(ReturnTotals.DayNoCashTips, "0.00")
    .TextMatrix(WhatRow, 3) = Format(Grat, CurrencyPreFix)
    .TextMatrix(WhatRow, 4) = Format(ReturnTotals.DayCashTenders - (ReturnTotals.DayNoCashTips + Grat), "0.00")
    If TTTL > 0 And TCT > 0 Then
    WhatRow = WhatRow + 1
    .TextMatrix(WhatRow, 0) = "** Tranaction Avg  " & Format(TTTL / TCT, CurrencyPreFix)
    .TextMatrix(WhatRow, 1) = "Tran # " & Format(TCT, "###")
    .TextMatrix(WhatRow, 2) = ""
    .TextMatrix(WhatRow, 3) = "Trans $ " & Format(TTTL, "0.00")
     If CCT <> TCT Then
     .TextMatrix(WhatRow, 4) = "Cust Ct " & Format(CCT, "###")
     Else
     .TextMatrix(WhatRow, 4) = ""
     End If
     If DualPricing > 0 Then '// dual price when the software raises the prices
      WhatRow = WhatRow + 1
      .TextMatrix(WhatRow, 0) = "** Dual Price " & (DualPricing * 100) & "% " & Format((GetDualPriceAmount(InDate, EDate) * DualPricing), "$0.00")
      WhatRow = WhatRow + 1
      .TextMatrix(WhatRow, 0) = "** Dual Price Disc  " & Format((ReturnTotals.DayCashTenders * DualPricing), "$0.00")
     End If
     If CashDiscountData.CreditAmount > 0 And DualPricing = 0 Then ' dual price where you raise your prices
      WhatRow = WhatRow + 1
      .TextMatrix(WhatRow, 0) = "** Dual Price " & CashDiscountData.CreditAmount & "% " & Format((GetDualPriceAmount(InDate, EDate) * (CashDiscountData.CreditAmount / 100)), "$0.00")
     End If
     If ProcessingFee > 0 Then
      WhatRow = WhatRow + 1
     .TextMatrix(WhatRow, 0) = "** Processing Fee " & Format(ProcessingFee, "$0.00")
     End If
     If MealCount.MealCount > 0 And MealCount.MealSale > 0 Then '// tied to menu items that count as a meal
      WhatRow = WhatRow + 1
     .TextMatrix(WhatRow, 0) = "** Meal Count " & Format(MealCount.MealCount, "0.00")
      WhatRow = WhatRow + 1
     .TextMatrix(WhatRow, 0) = "** Meal Avg.  " & Format(MealCount.MealSale / MealCount.MealCount, "0.00")
     End If
    End If
End With
Col1Row = WhatRow + 1
Col2Row = WhatRow + 1
Sql = "Select [Name], Sum([Amount]) as TAmount, Sum([Percentage]) as TPerc from [Labor] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " Group by [Name]"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
DptTotal = 0
With ShowEdit
Col1Row = Col1Row + 1
.TextMatrix(Col1Row, 0) = "Labor"
.TextMatrix(Col1Row, 2) = "Percentage %"

Do While Not RS.EOF
   If Val(RS![TAmount] & "") > 0 Then
      Col1Row = Col1Row + 1
      .TextMatrix(Col1Row, 0) = "  " & RS![Name] & ""
      .TextMatrix(Col1Row, 1) = Format(Val(RS![TAmount] & ""), "0.00")
      If NetS > 0 And Val(RS![TAmount] & "") > 0 Then
        .TextMatrix(Col1Row, 2) = Format(Val(RS![TAmount] & "") / NetS, "0.00%")
      End If
      DptTotal = DptTotal + Val(RS![TAmount] & "")
   End If
   RS.MoveNext
Loop
      Col1Row = Col1Row + 1
      .TextMatrix(Col1Row, 0) = "Total"
      .TextMatrix(Col1Row, 1) = Format(DptTotal, CurrencyPreFix)
      If NetS > 0 And DptTotal > 0 Then
      .TextMatrix(Col1Row, 2) = Format((DptTotal / NetS), "0.00%")
      End If
End With
End If


Sql = "Select [Name], Sum([Amount]) as TAmount from [Paid Outs] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " Group by [Name]"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
DptTotal = 0
With ShowEdit
Col2Row = Col2Row + 1
.TextMatrix(Col2Row, 3) = "Paid Outs"
Do While Not RS.EOF
      Col2Row = Col2Row + 1
      .TextMatrix(Col2Row, 3) = "  " & RS![Name] & ""
      .TextMatrix(Col2Row, 4) = Format(Val(RS![TAmount] & ""), "0.00")
      DptTotal = DptTotal + Val(RS![TAmount] & "")
   RS.MoveNext
Loop
      Col2Row = Col2Row + 1
      .TextMatrix(Col2Row, 3) = "Total"
      .TextMatrix(Col2Row, 4) = Format(DptTotal, CurrencyPreFix)
End With
End If
SetNothing RS
Sql = "Select [Name], Sum([Amount]) as TAmount from [Discounts] Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and '" & Format(CDate(CDate(EDate)), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " Group by [Name]"
ConnectADORS (Sql), Db, RS
If Not RS.EOF Then
DptTotal = 0
Col2Row = Col2Row + 1
With ShowEdit
Col2Row = Col2Row + 1
.TextMatrix(Col2Row, 3) = "Discounts"
Do While Not RS.EOF
      Col2Row = Col2Row + 1
      .TextMatrix(Col2Row, 3) = "  " & RS![Name] & ""
      .TextMatrix(Col2Row, 4) = Format(Val(RS![TAmount] & ""), "0.00")
      DptTotal = DptTotal + Val(RS![TAmount] & "")
   RS.MoveNext
Loop
      Col2Row = Col2Row + 1
      .TextMatrix(Col2Row, 3) = "Total"
      .TextMatrix(Col2Row, 4) = Format(DptTotal, CurrencyPreFix)
      'Col2Total = Col2Total + DptTotal
End With
End If
DbClose Db, RS
If Col1Row > Col2Row Then
  WhatRow = Col1Row
Else
  WhatRow = Col2Row
End If
ShowEdit.rows = WhatRow + 1
SetPage ReportView.View, "Daily Balance Report For " & ShowEdit.TextMatrix(0, 1) & " " & ShowEdit.TextMatrix(0, 2)
ReportView.View.RenderControl = ShowEdit.hwnd
ReportView.View.EndDoc
Sleep 0
If EmailReport Then
  If PBS = 1 Then
    ReportView.View.PrintDoc
  End If
  ReportView.EmailEOD "Daily Balance Report For " & ShowEdit.TextMatrix(0, 1) & " " & ShowEdit.TextMatrix(0, 2)
Else
 ReportView.Show 1
 ShowEdit.Visible = False
 ShowEdit.Left = -5000
End If
Call MemoryCheck(ReportView)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "DSR Balance Sheet ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Public Sub DailySummary(InDate As String, EmailReport As Boolean, PBS As Integer)
On Error GoTo ErrorHandler
Dim SalesTotal      As Currency
Dim SalesQty        As Integer
Dim LYSales         As Currency
Dim LYQty           As Integer
Dim Revenus         As Currency
Dim NetSales        As Currency
Dim SubTotal        As Currency
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim GH              As ADODB.Recordset
Dim Gt              As Currency
Dim Disc            As Currency
Dim Vds             As Currency
Dim Rfd             As Currency
Dim POUT            As Currency
Dim NCash           As Currency
Dim Cash            As Currency
Dim OS              As Currency
Dim Bank            As Currency
Dim AccountedFor    As Currency
Dim LastPos         As Long
Dim Lns             As Long
Dim L               As Long
Dim LastLine        As Long
Dim LC              As Integer
Dim CRLF            As String   'lfcr
Dim OverSales       As Currency
Dim OverShort       As Currency
Dim LastYear        As String
Dim HPrc            As Single
Dim DataFile        As String
Dim PosForODest     As String
Dim Data            As String
Dim F               As String
Dim H               As String
Dim Skim            As Currency
Dim Deposit         As Currency
Dim Hs              As Integer
Dim Terms           As Integer
Dim TipAdj          As Currency
Dim BusDay          As Integer
 If Not IsDate(InDate) Then
    See_Msg "Invalid Starting Date", 5, ""
    Exit Sub
 End If
 LastYear = CDate(InDate) - 365
 SetBetweenString InDate
 If Not IsDate(BusDayEndsAt) Then
  BusDayEndsAt = SetEndingTime
 End If
 ReportView.View.Orientation = orLandscape
 SetPage ReportView.View, "Daily Audit Report For " & InDate
 'ReportView.View.Text = GetWeather(InDate) & vbCrLf
 LastPos = GetCurY
 Dim FA As Integer
 If System_config = ExpressService Then
       Data = ""
       ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
          Sql = "Select * from [Audits] Where TRY_CAST([Start Date] AS datetime) + TRY_CAST([start time] AS datetime) " & SqlBetweenStatment
       ConnectADORS Sql, Db, RS
        Do While Not RS.EOF
          Data = Data & Format(RS![Start Time] & "", "HH:MM ampm") & "|"
          Data = Data & "Open" & "|"
          Data = Data & RS![Name] & "" & "|"
          Data = Data & RS![Register Name] & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & "|"
          Data = Data & "" & ";"
        RS.MoveNext
       Loop
      DbClose Db, RS
    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
        Sql = "Select * from [AuditSales] Where TRY_CAST([Date] AS date) = '" & InDate & "'"
        Sql = Sql & "Order by [start Time] asc"
        Call ConnectADORS(Sql, Db, RS)
        BusDay = 0
        If Not RS.EOF Then
            BusDay = RS.fields("Close Out Day").Value
        End If
        RS.Close
        Sql = "Select * from [AuditSales] Where TRY_CAST([Date] AS date) Between '" & InDate & "' and '" & SetEndingDate(InDate) & "'"
        Sql = Sql & " and [close out day] = " & BusDay
        Sql = Sql & " Order by TRY_CAST([Date] AS datetime), TRY_CAST([start Time] AS datetime) asc"

    ConnectADORS Sql, Db, RS
    LastPos = GetCurY
    Gt = 0: Disc = 0: Vds = 0: Rfd = 0: POUT = 0: NCash = 0: OS = 0: Bank = 0: Cash = 0: Skim = 0: Deposit = 0
    Do While Not RS.EOF
        
        Disc = Disc + RS![Discounts]
        Vds = Vds + RS![Voids]
        Rfd = Rfd + RS![Refunds]
        POUT = POUT + RS![Paid Outs]
        NCash = NCash + RS![Non Cash Tenders]
        Cash = Cash + Val(RS![Cash In Drawer] & "")
        Bank = Bank + Val(RS![Bank Deposit] & "") + Val(RS![Skim] & "")
        Skim = Skim + Val(RS![Skim] & "")
        Deposit = Deposit + Val(RS![Bank Deposit] & "")
        Gt = Gt + RS![Total To Account For] '- RS![Starting Drawer])
        'Gt = Gt + (Val(RS![Cash In Drawer] & "") + RS![Paid Outs]) + RS![Non Cash Tenders] '- RS![Starting Drawer])
        Data = Data & Format(RS![Start Time] & "", "HH:MM ampm") & "|"
        Data = Data & Format(RS![End Time] & "", "HH:MM ampm") & "|"
        Data = Data & RS![Operator] & "" & "|"
        Data = Data & RS![Register Name] & "" & "|"
        'Data = Data & Format((Val(RS![Cash In Drawer] & "") + RS![Paid Outs]) + RS![Non Cash Tenders], "0.00") & "|"    '- RS![Starting Drawer]
        '// Modified for David Peel Daily Audits for Dairy Queen.
        
        Data = Data & Format(Val(RS![Total To Account For] & ""), "0.00") & "|"    '- RS![Starting Drawer]
        Data = Data & Format(RS![Discounts], "0.00") & "|"
        Data = Data & Format(RS![Voids], "0.00") & "|"
        Data = Data & Format(RS![Refunds], "0.00") & "|"
        Data = Data & Format(RS![Paid Outs], "0.00") & "|"
        Data = Data & Format(RS![Non Cash Tenders], "0.00") & "|"
        Data = Data & Format(Val(RS![Cash In Drawer] & ""), "0.00") & "|"
        Data = Data & Format(Val(RS![Skim] & ""), "0.00") & "|"
        Data = Data & Format(Val(RS![Cash In Drawer] & "") - Val(RS![Skim] & ""), "0.00") & "|"
        'Data = Data & Format(Val(RS![Bank Deposit] & ""), "0.00") & "|"
        AccountedFor = ((RS![Bank Deposit])) + Val(RS![Skim] & "") - RS![Cash In Drawer]
        
        OS = OS + AccountedFor
        Data = Data & Format(RS![Bank Deposit] + Val(RS![Skim] & ""), "0.00") & "|"
        Data = Data & Format(AccountedFor, "0.00") & ";"
       RS.MoveNext
    Loop
        Data = Data & "" & "|"
        Data = Data & "" & "|"
        Data = Data & "" & "|"
        Data = Data & "" & "|"
        Data = Data & Format(Gt, "0.00") & "|"
        Data = Data & Format(Disc, "0.00") & "|"
        Data = Data & Format(Vds, "0.00") & "|"
        Data = Data & Format(Rfd, "0.00") & "|"
        Data = Data & Format(POUT, "0.00") & "|"
        Data = Data & Format(NCash, "0.00") & "|"
        Data = Data & Format(Cash, "0.00") & "|"
        Data = Data & Format(Skim, "0.00") & "|"
        'Data = Data & Format(Deposit, "0.00") & "|"
        Data = Data & Format(Cash, "0.00") & "|"
        Data = Data & Format(Skim + Deposit, "0.00") & "|"
        If System_config = ExpressService Then
          'Data = Data & Format(Bank, "0.00") & "|"
          Data = Data & Format(OS, "0.00") & ";"
        Else
          Data = Data & Format(OS, "0.00") & ";"
        End If
  DbClose Db, RS
  F = "+<1000|+<1000|+<1500|+<1500|+>1000|+>800|+>800|+>800|+>800|+>1000|+>1000|+>1000|+>1000|+>1000|+>800;"
  H = "Start|Ending|Cashier|Register|Acct For|Disc|Void|Refund|Paid Out|Non Cash|Cash Collect|Skim|Expected|Total CID|+/-;"
  ReportView.View.MarginLeft = ".20in"
  ReportView.View.CurrentY = 2700
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
End If
LastLine = 100
'************************ hourly sales  *****************************************
'ConnectADODB SvrPath.TerminalPath & "Time history.mdb", "Time history.mdb", Db
'  Sql = "Select * From [Employee Records] Where "
'  Sql = Sql & " TRY_CAST([Shift Date] AS datetime) = '" & InDate & "'"
'ConnectADORS Sql, Db, RS
'If RS.EOF Then
  DataFile = TimeRecordsDb
'Else
'  DataFile = "Time history.Mdb"
'End If
'DbClose Db, RS
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
  Sql = "Select * from [Hourly Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(Format(InDate, "MM/DD/YYYY")), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by TRY_CAST([date] AS datetime) , TRY_CAST([Hour] AS datetime) asc"
ConnectADORS (Sql), Db, RS
LastPos = GetCurY
     SalesTotal = 0: SalesQty = 0:  LYSales = 0: LYQty = 0: Data = ""
     Do While Not RS.EOF
         
         If Hour(Format(RS![Hour] & "", "HH:MM ampm")) > Hour(BusDayEndsAt) Then
           Data = Data & Format(RS![Hour] & "", "HH:MM ampm") & "|"
           Data = Data & Format(Val(RS![Amount] & ""), "0.00") & "|"
           HPrc = HourlyLaborData(InDate, Hour(RS![Hour]), Val(RS![Amount] & ""), DataFile)
           Data = Data & Format(HPrc, "0.0") & "|"
           Data = Data & Format(Val(RS![Quantity] & ""), "###") & "|"
           SalesTotal = SalesTotal + Val(RS![Amount] & "")
           SalesQty = SalesQty + Val(RS![Quantity] & "")
           Sql = "Select * from [Hourly Sales] Where Year(TRY_CAST([Date] AS datetime)) = " & Year(LastYear)
           Sql = Sql & " and datepart('ww', cdate([date]),2,2) = " & DatePart("ww", InDate, vbMonday, vbFirstFourDays)
           Sql = Sql & " and datepart('w',cdate([date]),2,2) = " & DatePart("w", InDate, vbMonday, vbFirstFourDays)
           Sql = Sql & " and TRY_CAST([Hour] AS time) = TRY_CAST('" & Format(RS![Hour] & "", "HH:nn") & "' AS time)"
           ConnectADORS Sql, Db, GH
           If Not GH.EOF Then
            Data = Data & Format(Val(GH![Amount] & ""), "0.00") & "|"
            Data = Data & Format(Val(GH![Quantity] & ""), "###") & "|"
            Data = Data & Format(Val(RS![Amount] & "") - Val(GH![Amount] & ""), "0.00") & ";"
            LYSales = LYSales + Val(GH![Amount] & "")
            LYQty = LYQty + Val(GH![Quantity] & "")
           Else
            Data = Data & "" & "|"
            Data = Data & "" & "|"
            Data = Data & "" & ";"
           End If
           SetNothing GH
         End If
       RS.MoveNext
     Loop
     If Not RS.BOF Then RS.MoveFirst
     Do While Not RS.EOF
         'If Hour(Format(RS![Hour] & "", "HH:MM ampm")) = Hour(Format(RS![Hour] & "", "HH:MM ampm")) Then Stop
         If Hour(Format(RS![Hour] & "", "HH:MM ampm")) <= Hour(BusDayEndsAt) Then
           Data = Data & Format(RS![Hour] & "", "HH:MM ampm") & "|"
           Data = Data & Format(Val(RS![Amount] & ""), "0.00") & "|"
           HPrc = HourlyLaborData(InDate, Hour(RS![Hour]), Val(RS![Amount] & ""), DataFile)
           Data = Data & Format(HPrc, "0.0") & "|"
           Data = Data & Format(Val(RS![Quantity] & ""), "###") & "|"
           SalesTotal = SalesTotal + Val(RS![Amount] & "")
           SalesQty = SalesQty + Val(RS![Quantity] & "")
           'Sql = "Select * from [Hourly Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(LastYear), "yyyy-mm-dd hh:nn:ss") & "'"
           Sql = "Select * from [Hourly Sales] Where Year(TRY_CAST([Date] AS datetime)) = " & Year(LastYear)
           Sql = Sql & " and datepart('ww', cdate([date]),2,2) = " & DatePart("ww", InDate, vbMonday, vbFirstFourDays)
           Sql = Sql & " and datepart('w',cdate([date]),2,2) = " & DatePart("w", InDate, vbMonday, vbFirstFourDays)
           Sql = Sql & " and TRY_CAST([Hour] AS time) = TRY_CAST('" & Format(RS![Hour] & "", "HH:nn") & "' AS time)"
           ConnectADORS Sql, Db, GH
           'Data = GH![Date]
           If Not GH.EOF Then
            Data = Data & Format(Val(GH![Amount] & ""), "0.00") & "|"
            Data = Data & Format(Val(GH![Quantity] & ""), "###") & "|"
            Data = Data & Format(Val(RS![Amount] & "") - Val(GH![Amount] & ""), "0.00") & ";"
            LYSales = LYSales + Val(GH![Amount] & "")
            LYQty = LYQty + Val(GH![Quantity] & "")
           Else
            Data = Data & "" & "|"
            Data = Data & "" & "|"
            Data = Data & "" & ";"
           End If
           SetNothing GH
          End If
       RS.MoveNext
     Loop
     SetNothing RS
     Data = Data & "Total" & "|"
     Data = Data & Format(SalesTotal, "0.00") & "|"
     Data = Data & "" & "|"
     Data = Data & Format(SalesQty, "####") & "|"
     Data = Data & Format(LYSales, "0.00") & "|"
     Data = Data & Format(LYQty, "####") & "|"
     Data = Data & Format(SalesTotal - LYSales, "0.00") & ";"
     F = "+<1000|+>1000|+>1000|^750|+>1000|^750|+>1000;"
     H = "Hour|Dollars|Labor%|Tran|LY $|LY Qty|+/-;"
     ReportView.View.MarginLeft = ".20in"
     ReportView.View.MarginTop = GetCurY
     ReportView.View.StartTable
     ReportView.View.TableBorder = tbBoxColumns
     ReportView.View.addTable F, H, Data, vbYellow
     ReportView.View.EndTable
     LastLine = Val(GetCurY)
     If LastLine < 7000 Then LastLine = 7061
''' End hourly-----------------------------------------------------------------
  Sql = "Select * from [Discounts] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Name] asc"
ConnectADORS Sql, Db, RS
  Data = ""
     SalesTotal = 0: SalesQty = 0: Lns = 0: LC = 0
     Do While Not RS.EOF
        LC = LC + 1
        SalesTotal = SalesTotal + Val(RS![Amount] & "")
        SalesQty = SalesQty + Val(RS![Quantity] & "")
        RS.MoveNext
     Loop
     If LC < 15 Then
       If Not RS.BOF Then RS.MoveFirst
       Do While Not RS.EOF
        Data = Data & RS![Name] & "" & "|"
        Data = Data & Format(Val(RS![Amount] & ""), "0.00") & "|"
        Data = Data & Format(Val(RS![Quantity] & ""), "###") & ";"
        RS.MoveNext
       Loop
      End If
     Data = Data & "Total" & "|"
     Data = Data & Format(SalesTotal, "0.00") & "|"
     Data = Data & Format(SalesQty, "####") & ";"
  SetNothing RS
  F = "+<2000|+>1000|^750;"
  H = "Discount|Amount|Qty;"
  ReportView.View.MarginLeft = "4.74in"
  ReportView.View.CurrentY = LastPos
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
''***************************************************************
  Sql = "Select * from [Tenders] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and [Name] <> 'Cash Tenders'"
  Sql = Sql & " Order by [Name] asc"
ConnectADORS Sql, Db, RS
     SalesTotal = 0: SalesQty = 0: Data = "": TipAdj = 0
     Do While Not RS.EOF
           Data = Data & RS![Name] & "" & "|"
           Data = Data & Format(Val(RS![Amount] & ""), "0.00") & ";"
           SalesTotal = SalesTotal + Val(RS![Amount] & "")
           TipAdj = TipAdj + Val(RS![Tip Total] & "")
       RS.MoveNext
     Loop
     
     Data = Data & "Total" & "|"
     Data = Data & Format(SalesTotal, "0.00") & ";"
     If Check_Option(237) = 0 Then TipAdj = 0
SetNothing RS
  F = "+<2000|+>1000;"
  H = "Tender Name|Amount;"
  ReportView.View.MarginLeft = "7.37in"
  ReportView.View.CurrentY = LastPos
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
  If LastLine < Val(GetCurY) Then LastLine = Val(GetCurY)
''***********************************************************
'end first Line
'*************************************************************
  Sql = "Select * from [Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS Sql, Db, RS
SalesTotal = 0: SalesQty = 0: Data = ""
SetX "0.1"
   If Not RS.EOF Then
           Data = Data & "Account For" & "|"
           If Input_Option(58) = "Yes" Then
            Data = Data & Format((Val(RS![Adjusted Sales] & "") + Val(RS![Discounts] & "")) + (Val(RS![Exempt Sales] & "")) + TipAdj, "0.00") & ";"
           Else
            Data = Data & Format((Val(RS![Adjusted Sales] & "") + Val(RS![Discounts] & "")) + (Val(RS![Exempt Sales] & "") - Val(RS![Charged Gratuity] & "") + TipAdj), "0.00") & ";"
           End If
           Data = Data & "Account For(Less Disc)" & "|"
           Data = Data & Format(Val(RS![Adjusted Sales] & "") + (Val(RS![Exempt Sales] & "") - Val(RS![Charged Gratuity] & "")), "0.00") & ";"
           Data = Data & "Sales Tax" & "|"
           Data = Data & Format(Val(RS![Sales Tax] & ""), "0.00") & ";"
           If Val(RS![Special Tax] & "") > 0 Then
           Data = Data & "Special Tax" & "|"
           Data = Data & Format(Val(RS![Special Tax] & ""), "0.00") & ";"
           End If
           Data = Data & "Refunds" & "|"
           Data = Data & Format(Val(RS![Refunds] & ""), "0.00") & ";"
           Data = Data & "Discounts" & "|"
           Data = Data & Format(Val(RS![Discounts] & ""), "0.00") & ";"
           Data = Data & "Non Sales Revenue" & "|"
           If Input_Option(58) = "Yes" Then
            Data = Data & Format(Val(RS![Exempt Sales] & "") + TipAdj, "0.00") & ";"
           Else
            Data = Data & Format((Val(RS![Exempt Sales] & "") + TipAdj) - Val(RS![Charged Gratuity] & ""), "0.00") & ";"
           End If
           Data = Data & "Net Sales" & "|"
           Data = Data & Format(Val(RS![Net Sales] & ""), "0.00") & ";"
           If Val(RS![None Taxed Sales] & "") > 0 Then
            Data = Data & "Taxed Sales" & "|"
            Data = Data & Format(Val(RS![Net Sales] & "") - Val(RS![None Taxed Sales] & ""), "0.00") & ";"
            Data = Data & "Non Taxed Sales" & "|"
            Data = Data & Format(Val(RS![None Taxed Sales] & ""), "0.00") & ";"
           End If
           Data = Data & "Voids" & "|"
           Data = Data & Format(Val(RS![Voids] & ""), "0.00") & ";"
           Data = Data & "Bank Deposit" & "|"
           Data = Data & Format(Val(RS![Deposits] & ""), "0.00") & ";"
           Data = Data & "Non Cash Tender" & "|"
           Data = Data & Format(Val(RS![Tenders] & ""), "0.00") & ";"
           Data = Data & "Deposits Collected" & "|"
           Data = Data & Format(Val(RS![Deposits Collected] & ""), "0.00") & ";"
           Data = Data & "Paid Outs" & "|"
           Data = Data & Format(Val(RS![Paid Outs] & ""), "0.00") & ";"
    End If
SetNothing RS
    Sql = "Select Sum([Exempt Sales]) as ESale, Sum([Charged Gratuity]) as TGrat"
    Sql = Sql & " ,Sum([Tenders]) as TTender, Sum([Paid Outs]) as TPaidOut, Sum([Deposits]) as TDeposits"
    Sql = Sql & " ,Sum([Net Sales]) as TNet ,Sum([Sales Tax]) as Stax, Sum([Special Tax]) as SpTax"
    Sql = Sql & " ,Sum([Liquor Tax]) as LTax, Sum([Deposits Collected])as TCollected, Sum([House Charges]) as HCharge"
    Sql = Sql & " ,Sum([Tip Charge Back]) as TCB"
    Sql = Sql & " from [Sales] Where "
    Sql = Sql & " TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(InDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
      If Input_Option(58) = "Yes" Then
        OverSales = Val(RS![ESale] & "") + TipAdj
      Else
        OverSales = (Val(RS![ESale] & "") + TipAdj) - Val(RS![TGrat] & "")
      End If
      'If Check_Option(237) = 1 Then
       ' OverSales = OverSales + TipAdj
      'End If
      If OverSales < 0 Then OverSales = 0
      Revenus = Val(RS![TTender] & "") + Val(Check_Null(RS![TPaidout])) + Val(Check_Null(RS![TDeposits])) - (Val(RS![TCollected] & "") + Val(RS![HCharge] & ""))
      Collected = Val(Check_Null(RS![TNet])) + Val(Check_Null(RS![STAX])) + Val(Check_Null(RS![SPTax])) + Val(Check_Null(RS![LTax])) + OverSales + Val(RS![TCB] & "")
      OverShort = Format(Revenus - Collected, "0.00")
      If OverShort < 0 Then
        Data = Data & "Short" & "|"
      ElseIf OverShort > 0 Then
        Data = Data & "Over" & "|"
      Else
        Data = Data & "Perfect" & "|"
      End If
      Data = Data & Format(OverShort, "0.00") & ";"
   End If
SetNothing RS
  'LastLine = LastLine + 100
  F = "+<2000|+>1000;"
  H = "Sales Recap|Amount;"
  ReportView.View.MarginLeft = ".20in"
  ReportView.View.CurrentY = LastLine
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
'******************************************************************
  Sql = "Select * from [Order Destination] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Order Destination] asc"
ConnectADORS Sql, Db, RS
     SalesTotal = 0: SalesQty = 0: Data = ""
     Do While Not RS.EOF
           Data = Data & RS![Order Destination] & "" & "|"
           Data = Data & Format(Val(RS![Amount] & ""), "0.00") & "|"
           Data = Data & Format(Val(RS![Quantity] & ""), "###") & "|"
           SalesTotal = SalesTotal + Val(RS![Amount] & "")
           SalesQty = SalesQty + Val(RS![Quantity] & "")
           If Val(RS![Quantity] & "") > 0 And Val(RS![Amount] & "") > 0 Then
             Data = Data & Format(Val(RS![Amount] & "") / Val(RS![Quantity] & ""), "0.00") & ";"
           Else
             Data = Data & "" & ";"
           End If
       RS.MoveNext
   Loop
SetNothing RS
     Data = Data & "Total" & "|"
     Data = Data & Format(SalesTotal, "0.00") & "|"
     Data = Data & Format(SalesQty, "###") & "|"
     If SalesTotal > 0 And SalesQty > 0 Then
       Data = Data & Format(SalesTotal / SalesQty, "0.00") & ";"
     Else
       Data = Data & "" & ";"
     End If
  F = "+<2000|+>1000|^750|^750;"
  H = "Destination|Amount|Qty|Avg;"
  ReportView.View.MarginLeft = "2.31in"
  ReportView.View.CurrentY = LastLine
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
  PosForODest = GetCurY
'''**********************************************************************
  Sql = "Select * from [Labor] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Name] asc"
ConnectADORS Sql, Db, RS
   SalesTotal = 0: SalesQty = 0: Data = ""
   Do While Not RS.EOF
       Data = Data & RS![Name] & "" & "|"
       Data = Data & Format(Val(RS![Amount] & ""), "0.00") & "|"
       Data = Data & Format(Val(RS![Percentage] & ""), "0.00") & ";"
   RS.MoveNext
  Loop
  SetNothing RS
     Sql = "Select * from [Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
     ConnectADORS Sql, Db, RS
     If Not RS.EOF Then
        Data = Data & "Hrs Used " & Format(RS![Labor Hours], "0.00") & "|"
        Data = Data & Format(RS![Labor Dollars], "0.00") & "|"
        Data = Data & Format(RS![Labor Percentage], "0.00") & ";"
     End If
  SetNothing RS
  F = "+<1750|+>1000|^750;"
  H = "Department Labor|Amount|Lab %;"
  ReportView.View.MarginLeft = "5.46in"
  ReportView.View.CurrentY = LastLine
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
'***************************************
  Sql = "Select * from [Paid Outs] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Name] asc"
ConnectADORS Sql, Db, RS
     SalesTotal = 0: SalesQty = 0: Data = ""
     Do While Not RS.EOF
           Data = Data & RS![Name] & "" & "|"
           Data = Data & Format(Val(RS![Amount] & ""), "0.00") & ";"
           SalesTotal = SalesTotal + Val(RS![Amount] & "")
       RS.MoveNext
     Loop
     Data = Data & "Total" & "|"
     Data = Data & Format(SalesTotal, "0.00") & ";"
SetNothing RS
  F = "+<1500|+>1000;"
  H = "Paid Outs|Amount;"
  ReportView.View.MarginLeft = "7.92in"
  ReportView.View.CurrentY = LastLine
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
  LastLine = Val(GetCurY)
  If LastLine < Val(GetCurY) Then LastLine = Val(GetCurY)
  
'***************************************
  Sql = "Select * from [Refunds] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Name] asc"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
         SalesTotal = 0: SalesQty = 0: Data = ""
         Do While Not RS.EOF
               Data = Data & RS![Name] & "" & "|"
               Data = Data & Format(Val(RS![Amount] & ""), "0.00") & ";"
               SalesTotal = SalesTotal + Val(RS![Amount] & "")
           RS.MoveNext
         Loop
  Data = Data & "Total" & "|"
  Data = Data & Format(SalesTotal, "0.00") & ";"
  F = "+<1500|+>1000;"
  H = "Refund|Amount;"
  ReportView.View.MarginLeft = "7.92in"
  ReportView.View.CurrentY = LastLine
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
End If
SetNothing RS
'''''''''''''''''''''''''''''''''''''''''''''''''
'------------------------------------------------
  Sql = "Select * from [Department Sales] Where TRY_CAST([Date] AS datetime) = '" & Format(CDate(InDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Name] asc"
ConnectADORS Sql, Db, RS
PrintLine "", -1
SetXandY "2.85in", PosForODest
SetX "0"
Data = ""
SalesTotal = 0: SalesQty = 0
     Do While Not RS.EOF
           Data = Data & RS![Name] & "|"
           Data = Data & Format(Val(RS![Amount] & ""), "0.00") & "|"
           Data = Data & Format(Val(RS![Quantity] & ""), "###") & ";"
           SalesTotal = SalesTotal + Val(RS![Amount] & "")
           SalesQty = SalesQty + Val(RS![Quantity] & "")
       RS.MoveNext
   Loop
DbClose Db, RS
ReportView.View.MarginLeft = "2.31in"
ReportView.View.MarginTop = LastLine
Data = Data & "Total" & "|"
Data = Data & Format(SalesTotal, "0.00") & "|"
Data = Data & Format(SalesQty, "###") & ";"
F = "+<1500|+>1000|^750;"
H = "Department|Amount|Qty;"
ReportView.View.StartTable
ReportView.View.TableBorder = tbBoxColumns
ReportView.View.addTable F, H, Data, vbYellow
On Error Resume Next
ReportView.View.EndTable
If Err.Number <> 0 Then
   See_Msg "Cannot Display Page...", 5, "Try Lower Font Size!!"
End If
ReportView.View.EndDoc
If EmailReport Then
 ReportView.EmailEOD "Daily Sale Summary"
Else
 ReportView.Show 1
End If
Call MemoryCheck(ReportView)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "DSR Balance Sheet ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Function HourlyLaborData(SDate As String, CountHrs As Integer, HrlySales As Currency, FileName As String) As Single
If IsDate(SDate) = False Then Exit Function
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim PRTb        As ADODB.Recordset
Dim EndTime     As String
Dim HrsUse      As Single
Dim HrlyPerc    As Single
Dim HRLYDollars     As String
Dim MinsWorked  As Single
Dim BH          As Integer
OpenDb SvrPath.TerminalPath & FileName, Db

  Sql = "Select * From [Employee Records] Where "
  If CountHrs = 0 Then CountHrs = 12
  Sql = Sql & CountHrs - 1 & " Between datepart(Hour,TRY_CAST([Time In] AS datetime)) and datepart(Hour,TRY_CAST([Time Out] AS datetime))"
  
  Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) = '" & SDate & "'"
  Sql = Sql & " and [Paid By] = 1"
  Sql = Sql & " and isdate([time in]) = 1 and isdate([Time Out]) = 1"
  ConnectADORS Sql, Db, PRTb
  
   HrlyPerc = 0: HRLYDollars = 0
    Do While Not PRTb.EOF
       EndTime = PRTb![Time Out] & ""
       If Hour(EndTime) >= CountHrs - 1 Then
        HrsUse = DateDiff("n", CDate(PRTb![Shift date]) & " " & PRTb![Time In], CDate(PRTb![Punch out Date]) & " " & CDate(EndTime))
        
        If HrsUse > 0 Then HrsUse = HrsUse / 60
        MinsWorked = DateDiff("n", CDate(SDate) & " " & CDate(CStr(CountHrs - 1) & ":00"), CDate(PRTb![Punch out Date] & " " & CDate(EndTime)))
        If MinsWorked > 0 And MinsWorked < 60 Then
         HrsUse = MinsWorked / 60
        End If
        If HrsUse > 1 Then HrsUse = 1
        HRLYDollars = HRLYDollars + Val(PRTb![Rate of Pay] & "") * HrsUse
       End If
      PRTb.MoveNext
    Loop
DbClose Db, PRTb
If HrlySales > 0 And HRLYDollars > 0 Then
  HourlyLaborData = Format((HRLYDollars / HrlySales) * 100, "0.00")
Else
  HourlyLaborData = 0
End If
Exit Function
ErrorHandler:
If Err.Number = 3464 Then
  See_Msg "Please Correct Time Records", 5, "Cannot Calculate Labor with Incomplete records"
  Exit Function
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Hourly Data Time Records ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Function
Private Sub Form_Load()
On Error GoTo ErrorHandler
Display.ZOrder 0
ResizeScreen1024 Me
'Me.Picture = LoadResPicture(101, 0)
'back ground loaded in resizescreen
Dim Cdb As ADODB.Connection
Dim Mdb As ADODB.Connection
Dim Location As String
Dim OverSales As Currency
Dim LP As Variant
Dim LL As Integer
Dim FFDOW As Integer
Dim SearchDate As String
MonthView1.StartOfWeek = WeekStartsOn
MonthView1.Value = Date - (MonthView1.StartOfWeek + 1)

Dim RS As ADODB.Recordset
SummarySheet.Clear
With SummarySheet
   .ColWidth(0) = 1700 * WidScale1024
   For LL = 1 To 7
   .ColWidth(LL) = 1100 * WidScale1024
   Next LL
   .rows = 30
End With
Compare.Clear
With Compare
   .ColWidth(0) = 1700 * WidScale1024
   For LL = 1 To 4
   .ColWidth(LL) = 1100 * WidScale1024
   Next LL
End With
With Compare
  .TextMatrix(0, 0) = ""
  .TextMatrix(0, 1) = "Current"
  .TextMatrix(0, 2) = "Past"
  .TextMatrix(0, 3) = "+/- $"
  .TextMatrix(0, 4) = "+/- %"
  .TextMatrix(1, 0) = "Gross Sales"
  .TextMatrix(2, 0) = "Net Sales"
  .TextMatrix(3, 0) = "Sales Tax"
  .TextMatrix(4, 0) = "Special Tax"
  .TextMatrix(5, 0) = SurChargeTxt
  .TextMatrix(6, 0) = "Voids"
  .TextMatrix(7, 0) = "Discounts"
  .TextMatrix(8, 0) = "Tender Types"
  .TextMatrix(9, 0) = "Paid Outs"
  .TextMatrix(10, 0) = "Bank Deposits"
  .TextMatrix(11, 0) = "Non Taxed Sales"
  .TextMatrix(13, 0) = "Ttl Accounted for"
  .TextMatrix(12, 0) = "Ttl To Acct for"
  .TextMatrix(14, 0) = "Cash Over/Short"
  .TextMatrix(15, 0) = "Labor Dollars"
  .TextMatrix(16, 0) = "Labor Hours"
  .TextMatrix(17, 0) = "Labor Percentage"
  .TextMatrix(18, 0) = "Sales Percentage"
  .TextMatrix(19, 0) = "Go Sales"
  .TextMatrix(20, 0) = "Call In Sales"
  .TextMatrix(21, 0) = "Drive Thru Sales"
  .TextMatrix(22, 0) = "Gift Cert/Paid In"
  .TextMatrix(23, 0) = "Tip Fee Income"
  .TextMatrix(24, 0) = "Refunds"
  .TextMatrix(25, 0) = "Deposit Paid In"
End With
Dim WTDsales As Currency
If wYear = 0 Then wYear = Year(Date)
OpenDb SvrPath.TerminalPath & ClosingHistory, Cdb
Sql = "Delete from [Sales] Where not isdate([date])"
Cdb.Execute Sql
Sql = "Select * from [Sales]"
Sql = Sql & " Order by TRY_CAST([date] AS datetime) desc"
ConnectADORS (Sql), Cdb, RS
  SearchDate = ""
  HistoryDate.Text = Format(Date, "MM/DD/YY")
  If Not RS.EOF Then
   SearchDate = RS.fields("Date").Value
   Do While Not RS.EOF
      HistoryDate.AddItem Format(RS!Date, "MM/DD/YY")
      HistoryDate.Text = Format(RS!Date, "MM/DD/YY")
     RS.MoveNext
   Loop
  End If
SetNothing RS
wYear = Year(HistoryDate.Text)
If Not IsDate(SearchDate) Then
 SearchDate = Format(Date, "mm/dd/yy")
End If
For FFDOW = 0 To HistoryDate.ListCount - 1
   If Format(HistoryDate.List(FFDOW), "dddd") = Input_Option(6) Then
     MonthView1_DateClick HistoryDate.List(FFDOW)
     Exit For
   End If
Next FFDOW
'ShiftWeeks ""
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Display Load ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub DiscountHistory()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim PrinterPort As String
Dim ReportType As String
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
  Sql = "Select [Menu Item],[Department Name], Sum([Item Price]) as [Price], Sum([Quantity]) as [Qty] from [Detail]"
  Sql = Sql & " Where [Item Price] < 0 and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate + " " + GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(SetEndingDate(SearchDates.Item(1).EndDate) + " " + SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " group by [Department Name],[Menu Item] order by [Department Name],[menu Item] asc"
ConnectADORS ReturnDateString(Sql), Db, RS
If RS.EOF Then
   See_Msg "No Discounts Found", 5, ""
   DbClose Db, RS
   Exit Sub
End If

If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    PF = FreeFile
    PrinterPort = "Discount History"
tryagain:
    Close
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    Print #PF, CenterText; "Discount History"; EC(7)
    Print #PF, CenterText; "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate; EC(7)
    Print #PF, LeftText & EC(7)
Else
    StartPage
    PrintLine "", -1
    PrintCLine "Discount History", -1, 0
    PrintCLine "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate, -1, 0
    PrintLine "", -1
End If
Dim TDisc As Currency, QDisc As Integer
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Name                     Disc       #"; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Discount Name", 2500: PrintLine "Type", 4500: PrintLine "Discount Amt", 6500
        PrintLine "Quantity", -1
        DrawLine
    End If
    Do While Not RS.EOF
            TDisc = TDisc + Val(RS![Price] & "")
            QDisc = QDisc + Val(RS![qty] & "")
            If Check_Option(96) = 0 Then
                Print #PF, Trim(RS![Menu Item]); String(20 - Len(Trim(RS![Menu Item])), 32); ":";
                Print #PF, Tab(26); Align(4, Abs(Val(RS![Price]))); Tab(36); Format(Val(RS![qty]), "##"); EC(7)
            Else
                PrintLine RS![Menu Item], 2500
                PrintLine RS![Department Name], 4500
                PrintLine Format(Abs(Val(RS![Price] & "")), CurrencyPreFix), 6500: PrintLine Format(Val(RS![qty] & ""), "##"), -1
            End If
       RS.MoveNext
   Loop
SetNothing RS
Sql = "Select [Department Name], Sum([Item Price]) as [Price], Sum([Quantity]) as [Qty] from [Detail]"
  Sql = Sql & " Where [Item Price] < 0 and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate + " " + GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(SetEndingDate(SearchDates.Item(1).EndDate) + " " + SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " group by [Department Name] order by [Department Name] asc"
ConnectADORS ReturnDateString(Sql), Db, RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
Else
    DrawLine
End If
TDisc = 0: QDisc = 0
Do While Not RS.EOF
TDisc = TDisc + Abs(Val(RS![Price]))
QDisc = QDisc + Val(RS![qty])
If Check_Option(96) = 0 Then
    Print #PF, RS![Department Name]; Tab(26); Align(7, Abs(Val(RS![Price]))); Tab(36); Format(RS![qty], "##"); EC(7)
Else
    PrintLine RS![Department Name], 4500
    PrintLine Format(Abs(RS![Price]), CurrencyPreFix), 6500
    PrintLine Format(Val(RS![qty]), "###"), -1
End If
   RS.MoveNext
Loop
SetNothing RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Total"; Tab(26); Align(7, Val(TDisc)); Tab(36); Format(QDisc, "##"); EC(7)
Else
    DrawLine
    PrintLine "Total", 4500
    PrintLine Format(TDisc, CurrencyPreFix), 6500
    PrintLine Format(QDisc, "###"), -1
End If

DbClose Db, RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Discount History ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub LaborHistory()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim PrinterPort As String
Dim ReportType As String
OpenDb SvrPath.TerminalPath & ClosingHistory, Db

 If SearchDates.Item(1).StartDate <> SearchDates.Item(1).EndDate Then
  Sql = "Select [Name], Sum([Amount]) as [Price], Sum([Percentage]) as [Qty] from [Labor]"
  Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " group by [Name] order by  [Name] asc"
 Else
  Sql = "Select [Name], [Amount] as [Price], [Percentage] as [Qty] from [Labor]"
  Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " order by  [Name] asc"
 End If
ConnectADORS (Sql), Db, RS
If RS.EOF Then
   See_Msg "No Labor Found", 5, ""
   DbClose Db, RS
   Exit Sub
End If

If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    PF = FreeFile
    PrinterPort = "Discount History"
tryagain:
    Close
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    
    Print #PF, CenterText; "Labor History"; EC(7)
    Print #PF, CenterText; "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate; EC(7)
    Print #PF, LeftText & EC(7)
Else
    StartPage
    PrintLine "", -1
    PrintCLine "Labor History", -1, 0
    PrintCLine "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate, -1, 0
    PrintLine "", -1
End If
Dim TDisc As Currency, QDisc As Integer
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Name                   Amt       %"; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Department", 2500: PrintLine "", 4500: PrintLine "Amount", 6500
        PrintLine "Percentage", -1
        DrawLine
    End If
    Do While Not RS.EOF
            TDisc = TDisc + Val(RS![Price] & "")
            QDisc = QDisc + Val(RS![qty] & "")
           If RS![Price] > 0 Then
            If Check_Option(96) = 0 Then
                Print #PF, Trim(RS![Name]); String(20 - Len(Trim(RS![Name])), 32); ":";
                Print #PF, Tab(26); Align(4, Abs(Val(RS![Price]))); Tab(36); Format(Val(RS![qty]), "0.0"); EC(7)
            Else
                PrintLine RS![Name], 2500
                PrintLine "", 4500
                PrintLine Format(Abs(Val(RS![Price])), CurrencyPreFix), 6500: PrintLine Format(Val(RS![qty]), "0.0"), -1
            End If
           End If
       RS.MoveNext
   Loop
'DbClose db, Rs
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Total"; Tab(26); Align(7, Val(TDisc)); EC(7)
Else
    DrawLine
    PrintLine "Total", 4500
    PrintLine Format(TDisc, CurrencyPreFix), -1
    'PrintLine Format(QDisc, "##.#"), -1
End If

DbClose Db, RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Labor History ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub TenderHistory()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim PrinterPort As String
Dim ReportType As String
OpenDb SvrPath.TerminalPath & ClosingHistory, Db
  Sql = "Select [Name], Sum([Amount]) as [Price],Sum([Tip Total]) as [TTotal] from [Tenders]"
  Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and [Amount] <> 0"
  Sql = Sql & " group by [Name] order by [Name] asc"
ConnectADORS Sql, Db, RS
If RS.EOF Then
   See_Msg "No Tenders Found", 5, ""
   DbClose Db, RS
   Exit Sub
End If
If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    PF = FreeFile
    PrinterPort = "Discount History"
tryagain:
    Close
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    
    Print #PF, CenterText; "Non Cash Tender History"; EC(7)
    Print #PF, CenterText; "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate; EC(7)
    Print #PF, LeftText & EC(7)
Else
    StartPage
    PrintLine "", -1
    PrintCLine "Non Cash Tender History", -1, 0
    PrintCLine "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate, -1, 0
    PrintLine "", -1
End If
Dim TDisc As Currency, QDisc As Integer
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Name                     Amount    Tip"; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Tender Name", 2500: PrintLine "Tip", 4500: PrintLine "Amount", -1
        'PrintLine "Quantity", -1
        DrawLine
    End If
    Do While Not RS.EOF
            TDisc = TDisc + Val(RS![Price] & "")
            If Check_Option(96) = 0 Then
                Print #PF, Trim(RS![Name]); String(20 - Len(Trim(RS![Name])), 32); ":";
                Print #PF, Tab(26); Align(4, Val(RS![Price])); Tab(36); Format(Val(RS![TTotal]), "0.00"); EC(7)
            Else
                PrintLine RS![Name], 2500
                PrintLine Format(Val(RS![TTotal] & ""), "0.00"), 4500
                PrintLine Format(Val(RS![Price]), "0.00"), -1   ': PrintLine Format(Val(Rs![qty]), "##"), -1
            End If
       RS.MoveNext
   Loop
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Total"; Tab(26); Align(7, Val(TDisc)); EC(7) 'Tab(36); Format(QDisc, "##"); Ec(7)
Else
    DrawLine
    PrintLine "Total", 4500
    PrintLine Format(TDisc, CurrencyPreFix), -1
    'PrintLine Format(QDisc, "###"), -1
End If

DbClose Db, RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Tender History ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select


End Sub
Private Sub HistoryReport(ReportHeading As String, Sql As String)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim PrinterPort As String
Dim ReportType As String
OpenDb SvrPath.TerminalPath & ClosingHistory, Db
ConnectADORS Sql, Db, RS
If RS.EOF Then
   See_Msg "No " & ReportHeading & " Found", 5, ""
   DbClose Db, RS
   Exit Sub
End If
If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    PF = FreeFile
    PrinterPort = "Discount History"
tryagain:
    Close
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    
    Print #PF, CenterText; ReportHeading; EC(7)
    Print #PF, CenterText; "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate; EC(7)
    Print #PF, LeftText & EC(7)
Else
    StartPage
    PrintLine "", -1
    PrintCLine ReportHeading, -1, 0
    PrintCLine "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate, -1, 0
    PrintLine "", -1
End If
Dim TDisc As Currency, QDisc As Integer
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Name                     Amount       "; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Name", 2500: PrintLine "", 4500: PrintLine "Amount", -1
        'PrintLine "Quantity", -1
        DrawLine
    End If
    Do While Not RS.EOF
            TDisc = TDisc + Val(RS![Price] & "")
            If Check_Option(96) = 0 Then
                Print #PF, Trim("Deposit"); String(20 - Len(Trim("Deposit")), 32); ":";
                Print #PF, Tab(26); Align(4, Abs(Val(RS![Price]))); EC(7) '; Tab(36); Format(Val(Rs![qty]), "##"); Ec(7)
            Else
                PrintLine "Deposit", 2500
                PrintLine "", 4500
                PrintLine Format(Abs(Val(RS![Price] & "")), CurrencyPreFix), -1  ': PrintLine Format(Val(Rs![qty]), "##"), -1
            End If
       RS.MoveNext
   Loop
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Total"; Tab(26); Align(7, Val(TDisc)); EC(7) 'Tab(36); Format(QDisc, "##"); Ec(7)
Else
    DrawLine
    PrintLine "Total", 4500
    PrintLine Format(TDisc, CurrencyPreFix), -1
    'PrintLine Format(QDisc, "###"), -1
End If

DbClose Db, RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Report History ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub PaidOutHistory(Sql As String)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim PrinterPort As String
Dim ReportType As String
OpenDb SvrPath.TerminalPath & ClosingHistory, Db
ConnectADORS Sql, Db, RS
If RS.EOF Then
   See_Msg "No Paid Outs Found", 5, ""
   DbClose Db, RS
   Exit Sub
End If
If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    PF = FreeFile
    PrinterPort = "Discount History"
tryagain:
    Close
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    
    Print #PF, CenterText; "Paid Out History"; EC(7)
    Print #PF, CenterText; "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate; EC(7)
    Print #PF, LeftText & EC(7)
Else
    StartPage
    PrintLine "", -1
    PrintCLine "Paid Out History", -1, 0
    PrintCLine "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate, -1, 0
    PrintLine "", -1
End If
Dim TDisc As Currency, QDisc As Integer
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Name                     Amount       "; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Name", 2500: PrintLine "", 4500: PrintLine "Amount", -1
        'PrintLine "Quantity", -1
        DrawLine
    End If
    Do While Not RS.EOF
            TDisc = TDisc + Val(RS![Price] & "")
            'QDisc = QDisc + Val(Rs![qty] & "")
            If Check_Option(96) = 0 Then
                Print #PF, Trim(Left(RS![Name], 20)); String(20 - Len(Trim(Left(RS![Name], 20))), 32); ":";
                Print #PF, Tab(26); Align(4, Abs(Val(RS![Price]))); EC(7) '; Tab(36); Format(Val(Rs![qty]), "##"); Ec(7)
            Else
                PrintLine RS![Name], 2500
                PrintLine "", 4500
                PrintLine Format(Abs(Val(RS![Price])), CurrencyPreFix), -1   ': PrintLine Format(Val(Rs![qty]), "##"), -1
            End If
       RS.MoveNext
   Loop
'DbClose db, Rs
'TDisc = TDisc + Abs(Val(Rs![Price]))
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Total"; Tab(26); Align(7, Val(TDisc)); EC(7) 'Tab(36); Format(QDisc, "##"); Ec(7)
Else
    DrawLine
    PrintLine "Total", 4500
    PrintLine Format(TDisc, CurrencyPreFix), -1
    'PrintLine Format(QDisc, "###"), -1
End If

DbClose Db, RS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Paid Out History ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub FRS_Click()

End Sub

Private Sub GAME_Click()

End Sub

Private Sub HistoryDate_Click()
On Error Resume Next
'SummarySheet.Clear
If IsDate(HistoryDate.Text) = False Then Exit Sub
wYear = Year(HistoryDate.Text)
SearchY.Caption = "Searching Year " & wYear
Scroll = DatePart("WW", CDate(HistoryDate.Text), WeekStartsOn, vbFirstFullWeek)
ShiftWeeks HistoryDate.Text, ""
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Display History Date ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub IBT_Click()
LoadDates "Select Distinct [Date] from [Sales] order by [Date] asc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
ItemSaleTime SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate

End Sub
Private Sub ItemSaleTime(StartDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Records     As ADODB.Recordset
Dim Tb          As ADODB.Recordset
Dim RTB         As ADODB.Recordset
Dim T_Rec       As Long
Dim R           As Long
Dim CSvrName    As String
Dim Data        As String
Dim LCount      As Long
 ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
   Sql = "Select * from [Detail]"
   Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & CDate(StartDate & " " & GetStartTime()) & "'"
   Sql = Sql & " and '" & CDate(SetEndingDate(EndDate) & " " & SetEndingTime()) & "'"
   Sql = Sql & " Order by [Server Name],TRY_CAST([Date] AS datetime),TRY_CAST([Time] AS datetime) asc"
 ConnectADORS Sql, Db, Tb
 If Tb.EOF Then
    See_Msg "No Sales found", 5, ""
    DbClose Db, Tb
    Exit Sub
 End If
Update.Show
Update.UPdateWhere.Caption = "Compiling sales data...."
T_Rec = RecordsCount(Tb)
R = 0
 CSvrName = "": Data = ""
 Do While Not Tb.EOF
        R = R + 1
        Update.Progress.ProgressBarValue = (R / T_Rec) * 100
        Update.Progress.Refresh
        'Sleep 0
        If Tb.fields("Server Name").Value <> CSvrName Then
         Data = Data & Tb.fields("Server Name").Value & "|"
         Data = Data & ">>" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & ";"
         CSvrName = Tb.fields("Server Name").Value
        End If
        Data = Data & Format(Tb![Date], "MM/DD/YY") & "|"
        Data = Data & Format(Tb![Time], "HH:MM ampm") & "|"
        Data = Data & Tb![Menu Item] & "|"
        Data = Data & Format(Val(Tb![Item Price] & ""), "0.00") & "|"
        Data = Data & Tb![Revenue Center] & "|"
        Data = Data & Tb![Check Number] & ";"
       Tb.MoveNext
Loop
Unload Update
DbClose Db, Tb
Dim H$
H = "Date|Time|Menu Item|Price|Revenue Center|Check Number;"
Dim F$
F = "+<1600|+<1000|+>2500|+>750|+>1500|+>1000;"
ReportView.View.Orientation = orPortrait
SetPage ReportView.View, "Transactions " & "From " & Format(StartDate, "MM/DD/YY") & " To " & Format(EndDate, "MM/DD/YY")
ReportView.View.TableBorder = tbTopBottom
ReportView.View.addTable F, H, Data, vbBlack, vbWhite
FinishJob 1
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Mix Reports")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub KeyCal_Click()
If KeyCal.Caption = "Show Calculator" Then
  Dim DrwCt As AllCounts
  Set DrwCt = New AllCounts
  Dim MYC As GeneralLib.Utilities
  Set MYC = New GeneralLib.Utilities
  StoreDate.Text = Format(MYC.GetCalculator(DrwCt), "0.00")
  Set DrwCt = Nothing
  Set MYC = Nothing
Else
  StoreDate.Text = ShowKeyBoard("Enter Date")
End If
End Sub

Private Sub LB_Click()
Time_Reports.SSCommand4.DoClick
End Sub

Private Sub LBD_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
         Sql = "Select [Type] as Name,[Amount] as [Price] from [Paid Outs]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " order by [Type] asc"
        PaidOutHistory Sql
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub LBRHST_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        LaborHistory
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub mday_Click(Index As Integer)
End Sub

Private Sub LPR_Click()
Time_Reports.SSCommand2.DoClick

End Sub

Private Sub MCH_Click()

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
SummarySheet.Clear flexClearEverywhere, flexClearText

HistoryDate.Text = Format(DateClicked, "MM/DD/YYYY")
wYear = Year(HistoryDate.Text)
SearchY.Caption = "Searching Year " & wYear
Scroll = DatePart("WW", CDate(HistoryDate.Text), WeekStartsOn, vbFirstFullWeek)
ShiftWeeks HistoryDate.Text, ""
End Sub

Private Sub MPeriods_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
       Financal_Reports.MealPeriodReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub MRS_Click()
End Sub
Private Sub GetRecords(InDate As String, SqlStr As String)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Cnt As Integer
Dim Data()
Dim Labels()
Dim Colors()
ReDim Data(10)
ReDim Labels(10)
OpenDb SvrPath.TerminalPath & ClosingHistory, Db
ConnectADORS SqlStr, Db, RS
Charts.DisplayGrid.Clear flexClearEverywhere, flexClearText
If Not RS.EOF Then
        Charts.DisplayGrid.cols = 2
        Charts.DisplayGrid.TextMatrix(0, 0) = InDate
        Charts.DisplayGrid.TextMatrix(0, 1) = "Sales"
        Charts.DisplayGrid.rows = RS.RecordCount + 1
        ReDim FieldName(RS.fields.count - 1)
        For Cnt = 0 To RS.RecordCount - 1
          Labels(Cnt) = Cnt + 1
          Data(Cnt) = Format(RS![Net Sales] & "", "0.00")
          Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS![Date] & ""
          Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(RS![Net Sales] & "", "0.00")
          If Cnt = 10 Then Exit For
          RS.MoveNext
        Next
        Charts.DisplayGrid.ColWidth(0) = 1500 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1500 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        '
 ' colors = Array(&HB8BC9C, &HECF0B9, &H999966, &H333366, &HC3C3E6, &H594330, &HA0BDC4)
    Colors = Array(&HB8BC9C, &HA0BDC4, &H999966, &H333366, &HC3C3E6)
    Charts.DisplayGrid.Visible = True
    Charts.SetBarData Data, Labels, Colors, "Sales Data"
    Charts.Show 1
        
End If
DbClose Db, RS
End Sub
Private Sub Weeklytotals(PassSql As String)
If Not IsDate(SearchDate) Then Exit Sub
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Data()
Dim Labels()
Dim Colors()
    OpenDb SvrPath.TerminalPath & ClosingHistory, Cdb
        ConnectADORS PassSql, Cdb, RS
        If RS.RecordCount = 0 Then SetNothing RS: Set RS = Nothing: Cdb.Close: Set Cdb = Nothing: Exit Sub
        RS.MoveLast: RS.MoveFirst
        Charts.DisplayGrid.rows = RS.RecordCount + 1
        Charts.DisplayGrid.TextMatrix(0, 0) = "Name"
        Charts.DisplayGrid.TextMatrix(0, 1) = "Amount"
        Charts.DisplayGrid.TextMatrix(0, 2) = "Qty"
        ReDim Data(RS.RecordCount - 1)
        ReDim Labels(RS.RecordCount - 1)
        For Cnt = 0 To RS.RecordCount - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS(0)
            Labels(Cnt) = RS(0)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(RS(1), "0.00")
            Data(Cnt) = RS(1)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = Format(RS(2), "##")
            RS.MoveNext
        Next
        Charts.DisplayGrid.ColFormat(2) = "(#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1600 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 900 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1000 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        DbClose Cdb, RS
        Charts.SetData Data, Labels, Colors, "By Dollars" & " for " & LocName.Caption
        Charts.Show 1
        Exit Sub
End Sub

Private Sub MPR_Click()
End Sub
Private Sub MPS_Click()
If Not IsDate(SearchDate) Then Exit Sub
Sql = "Select [Name], sum([Amount]) as TSold, Sum([quantity]) as QSold"
Sql = Sql & " From [Meal periods] Where TRY_CAST([date] AS datetime) between '" & Format(CDate(SearchDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SearchDate) + 6), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " Group By [Name]"
Weeklytotals Sql
End Sub
Private Sub MSBUP_Click()
Dim Destination As String
Dim BeginDate   As String
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
     Sql = "Select [order Date] from [Closed Checks] order by [order Date] asc"
    ConnectADORS Sql, Db, Tb
     If Not Tb.EOF Then
       BeginDate = Tb.fields("Order Date").Value
     End If
     DbClose Db, Tb
     If Trim$(BeginDate) = "" Then BeginDate = Format(Date, "mm/dd/yyyy")
     If Not IsDate(BeginDate) Then BeginDate = Format(Date, "mm/dd/yyyy")
     Destination = Format(BeginDate, "MM-DD-YYYY") & " Week " & SvrPath.TerminalName & " Sales BackUp.MDB"
     FileCopy SvrPath.TerminalPath & "Chk-Stat.mdb", SvrPath.TerminalPath & Destination
     WeeklyBackUp BeginDate, "", Destination, 0
     See_Msg "Sales History Update Complete", 5, SvrPath.TerminalName
See_Msg "Backup Complete...", 5, ""
End Sub
Private Sub NoSaleReport_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
If SearchDates Is Nothing Then

Else
    Get_NoSaleReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End If
End Sub
Private Sub OSS_Click()
  ShowGraph 14, "Over/Short" & " " & LocName.Caption
End Sub
Private Sub PMix_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
If SearchDates Is Nothing Then

Else
    Mix_History "", "Mix", SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End If
End Sub
Private Sub Quit_Click()
Unload Me
End Sub
Private Sub RCD_Click()
Trans_Log "Accessing Reclose Utility"
Shell SvrPath.LocalTerminalPath & "MSaleReClose.EXE", vbNormalFocus
End Sub
Private Sub Refund_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
     RefundHistoryItemReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    End If
End If
End Sub
Private Sub Review_DateClick(ByVal DateClicked As Date)
SearchDate = Format(DateClicked, "MM/DD/YY")
With Review
  .Left = -4000
  .Visible = False
End With
End Sub
Private Sub RPSTE_Click()
End Sub
Private Sub SaleSum_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
If SearchDates Is Nothing Then

Else
    Financal_Reports.SummaryReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End If
End Sub
Private Sub SBD_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        If SearchDates.Item(1).StartDate <> SearchDates.Item(1).EndDate Then
         Sql = "Select [Type] as Name, Sum([Amount]) as [Price] from [Paid Outs]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " group by [Type] order by [Type] asc"
        Else
         Sql = "Select [Type] as Name, [Amount] as [Price] from [Paid Outs]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " order by [Type] asc"
        End If
        PaidOutHistory Sql
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub SBN_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        If SearchDates.Item(1).StartDate <> SearchDates.Item(1).EndDate Then
         Sql = "Select [Name],Sum([Amount]) as [Price] from [Paid Outs]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " group by [Name] order by [Name] asc"
        Else
         Sql = "Select [Name], [Amount] as [Price] from [Paid Outs]"
         Sql = Sql & " Where not isnull([date]) and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
         Sql = Sql & " order by [Name] asc"
        End If
        PaidOutHistory Sql
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub SBSP_Click()
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
OpenDb SvrPath.TerminalPath & ClosingHistory, SL
Sql = "Select Distinct [Date] from [Sales] order by [date] asc "
ConnectADORS Sql, SL, SRd
Do While Not SRd.EOF
   DatePicker.StartingDate.AddItem SRd![Date] & ""
   DatePicker.EndingDate.AddItem SRd![Date] & ""
   SRd.MoveNext
Loop
DbClose SL, SRd
If DatePicker.StartingDate.ListCount > -1 Then
  DatePicker.Show 1
  If SearchDates Is Nothing Then

  Else
    LoadIndividualData "Department"
    PaymentType = "Department"
  End If
Else
   See_Msg "Not Dates Found", 5, ""
End If
End Sub
Private Sub LoadIndividualData(WhichList As String)
On Error GoTo ErrorHandler
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
If WhichList = "Sales Mix" Or WhichList = "Department" Or WhichList = "Category" Then
  ConnectADODB SvrPath.TerminalPath & ProductHistory, SL
    Sql = "Select Distinct [Server Name] from [Detail] "
    Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  ConnectADORS ReturnDateString(Sql), SL, SRd
    SList.Clear
    If Not SRd.EOF Then
        
        CenterObject Me, ServerList
        ServerList.Visible = True
        Sleep 0
        SRd.MoveFirst
    Else
        DbClose SL, SRd
        Call See_Msg("No Names Found", 5, "")
        Exit Sub
    End If
       Do While Not SRd.EOF
          SList.AddItem SRd![Server Name]
          SRd.MoveNext
       Loop
    DbClose SL, SRd
    Exit Sub
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Show Edit ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SearchHow_Click()
SearchSpan = 0
End Sub
Private Sub ShowEdit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ClrC = True
End Sub
Private Sub ShowEdit_Click()
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
If ShowEdit.TextMatrix(ShowEdit.Row, ShowEdit.Col) = "(Print)" Then
    SetPage ReportView.View, "Daily Balance Report For " & ShowEdit.TextMatrix(0, 1) & " " & ShowEdit.TextMatrix(0, 2)
    Dim Pa As Integer, Cl As Long
    ReportView.View.RenderControl = ShowEdit.hwnd
    ReportView.View.EndDoc
    ReportView.Show 1
    Call MemoryCheck(ReportView)
    Exit Sub
End If
If ShowEdit.TextMatrix(ShowEdit.Row, ShowEdit.Col) = "Exit" Then
  ShowEdit.Visible = False
  ShowEdit.Left = -5000
  Exit Sub
End If
If ShowEdit.TextMatrix(ShowEdit.Row, ShowEdit.Col) = "Save" Then
  SaveEdit
  ShowEdit.Visible = False
  ShowEdit.Left = -5000
  Exit Sub
End If
If ShowEdit.TextMatrix(ShowEdit.Row, ShowEdit.Col) = "Rebuild Mix" Then
  StoreDetail DatePart("w", CDate(ShowEdit.TextMatrix(0, ShowEdit.Col))), False
  ShowEdit.Visible = False
  ShowEdit.Left = -5000
  Exit Sub
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Show Edit ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub ShowEdit_DblClick()
'If KeyAscii = 27 Then
'End If

End Sub
Private Sub SaveEdit()
EditRecord = False
 If See_Msg("Save Changes", 4, "Message") = 6 Then
    Dim EditTable As String
    Select Case EditItem
        Case Is = "Net Sales"
           EditTable = "Hourly"
        Case Is = "Paid Outs"
           EditTable = "Paid Outs"
        Case Is = "Discounts"
           EditTable = "Discounts"
        Case Is = "Charges", "NonCash"
           EditTable = "Tenders"
        Case Is = "Deposits"
           EditTable = "Deposits"
        Case Else
           EditTable = "Sales"
    End Select
       If EditTable = "Deposits" Then
        SaveChanges EditTable, EditDate
       Else
        SaveChanges EditTable, ShowEdit.TextMatrix(0, 1)
       End If
       EditItem = ""
 End If
 ShowEdit.Visible = False
 ShowEdit.Clear
 ShowEdit.Refresh
 Exit Sub
End Sub
Private Sub ShowEdit_EnterCell()
ClrC = True
End Sub
Private Sub ShowEdit_KeyPress(KeyAscii As Integer)
On Error Resume Next
'If EditType = "Deposit" Then Exit Sub
Dim result As Integer
EditRecord = True
With ShowEdit
If KeyAscii = 13 Then
    SummarySheet_Click
    Exit Sub
End If
Select Case .Row
    Case Is = 12, 13, 14
      Exit Sub
End Select
If ClrC = True Then
  ClrC = False
  .TextMatrix(.Row, .Col) = ""
End If
If KeyAscii = 8 Then
    If Len(.TextMatrix(.Row, .Col)) > 0 Then
        .TextMatrix(ShowEdit.Row, .Col) = _
        Mid(ShowEdit.TextMatrix(.Row, .Col), 1, Len(.TextMatrix(.Row, .Col)) - 1)
    End If
Else
If EditItem = "Deposits" Then
  .TextMatrix(14, 1) = 0
  .TextMatrix(.Row, .Col) = _
  .TextMatrix(.Row, .Col) & Chr(KeyAscii)
  For result = 1 To .rows - 1
  .TextMatrix(14, 1) = Format(Val(.TextMatrix(14, 1)) + Val(.TextMatrix(result, 1)), "0.00")
  If Val(.TextMatrix(result, 1)) = 0 Then Exit For
  Next result
Else
.TextMatrix(.Row, .Col) = _
.TextMatrix(.Row, .Col) & Chr(KeyAscii)
.TextMatrix(13, 1) = Format(Val(.TextMatrix(2, 1)) + Val(.TextMatrix(3, 1)) + Val(.TextMatrix(4, 1)) + Val(.TextMatrix(5, 1)), "0.00")
.TextMatrix(12, 1) = Format(Val(.TextMatrix(8, 1)) + Val(.TextMatrix(9, 1)) + Val(.TextMatrix(10, 1))) '+ Val(.TextMatrix(8, 1)), "0.00") '+ .TextMatrix(5, 1) - .TextMatrix(7, 1)
.TextMatrix(14, 1) = Format(Val(.TextMatrix(12, 1)) - Val(.TextMatrix(13, 1)), "0.00")
End If
End If
End With
End Sub
Private Sub ShowEdit_LeaveCell()
ClrC = True
End Sub
Private Sub SRS_Click()
End Sub

Private Sub SL_Click()
End Sub

Private Sub SList_Click()
Call See_Msg("Searching", 5, "")
Select Case PaymentType
   Case Is = "Sales Mix", "Department", "Category"
      MixAll SList.List(SList.ListIndex), PaymentType
End Select
End Sub
Private Sub MixAll(ServerName As String, ReportType As String)
On Error GoTo ErrorHandler
Dim SL      As ADODB.Connection
Dim SRd     As ADODB.Recordset
Dim Data    As String
ConnectADODB SvrPath.TerminalPath & ProductHistory, SL
        If ReportType = "Sales Mix" Then
                     Sql = "Select [server Name], [Menu Item],sum([Quantity]) as QSold, Sum([Item Price]) as TSold  from [Detail] "
                     Sql = Sql & " Where [Department Name] <> 'Gratuity' and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate) & " " & GetStartTime), "yyyy-mm-dd hh:nn:ss") & "'"
                     Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(SearchDates.Item(1).EndDate)) & " " & SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "'"
                     Sql = Sql & " and [Server Name] = '" & ServerName & "'"
                     Sql = Sql & " Group By  [Server Name],[Menu Item]"
                     SetPage ReportView.View, "Mix Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
        ElseIf ReportType = "Category" Then
                     Sql = "Select [server Name], [Category Name],sum([Quantity]) as QSold, Sum([Item Price]) as TSold from [Detail] "
                     Sql = Sql & " Where [Department Name] <> 'Gratuity' and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate) & " " & GetStartTime), "yyyy-mm-dd hh:nn:ss") & "'"
                     Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(SearchDates.Item(1).EndDate)) & " " & SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "'"
                     Sql = Sql & " and [Server Name] = '" & ServerName & "'"
                     Sql = Sql & " Group By  [Server Name],[Category Name]"
                     SetPage ReportView.View, "Category Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
        Else
                     Sql = "Select [server Name], [Department Name],sum([Quantity]) as QSold, Sum([Item Price]) as TSold  from [Detail] "
                     Sql = Sql & " Where [Department Name] <> 'Gratuity' and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate) & " " & GetStartTime), "yyyy-mm-dd hh:nn:ss") & "'"
                     Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(SearchDates.Item(1).EndDate)) & " " & SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "'"
                     Sql = Sql & " and [Server Name] = '" & ServerName & "'"
                     Sql = Sql & " Group By  [Server Name],[Department Name]"
                     SetPage ReportView.View, "Department Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
        End If
If Sql = "" Then Exit Sub   'if no match of crieteria then exit sub
Dim H$
H = "Server|Item Name|Quantity|Amount|Percentage;"
Dim F$
F = "+<1500|+<2000|+>1000|+>1500|+>1000|;"
Data = ServerName
Dim TSales, TTSales As Currency
Dim QSales As Long, TQty As Long
Dim TSvr As String
Dim TSSales As Currency
Dim MyPerc As Double
TSales = 0:  QSales = 0
TTSales = 0: TQty = 0
TSSales = 0
ConnectADORS ReturnDateString(Sql), SL, SRd
  Do While Not SRd.EOF
     TSSales = TSSales + Val(SRd!TSold & "")
    SRd.MoveNext
  Loop
  If Not SRd.BOF Then SRd.MoveFirst
  Do While Not SRd.EOF
        MyPerc = 0
        If Val(SRd!TSold & "") <> 0 And TSSales <> 0 Then
          MyPerc = (SRd!TSold / TSSales) * 100
        End If
        Data = Data & "" & "|"
        If ReportType = "Sales Mix" Then
          Data = Data & SRd![Menu Item] & "" & "|"
          Data = Data & Format(Val(SRd!QSold & ""), "##") & "|"
          Data = Data & Format(Val(SRd!TSold & ""), "0.00") & "|"
          Data = Data & Format(MyPerc, "#.00") & ";"
        ElseIf ReportType = "Category" Then
            Data = Data & SRd![Category Name] & "" & "|"
            Data = Data & Format(SRd!QSold, "##") & "|"
            Data = Data & Format(SRd!TSold, "0.00") & "|"
            Data = Data & Format(MyPerc, "#.00") & ";"
        Else
            Data = Data & SRd![Department Name] & "" & "|"
            Data = Data & Format(SRd!QSold, "##") & "|"
            Data = Data & Format(SRd!TSold, "0.00") & "|"
            Data = Data & Format(MyPerc, "#.00") & ";"
        End If
          TSales = TSales + SRd!TSold
          QSales = QSales + SRd!QSold
          TTSales = TTSales + (SRd!TSold)
          TQty = TQty + SRd!QSold
        SRd.MoveNext
     Loop
DbClose SL, SRd
Data = Data & "" & "|"
Data = Data & "Total" & "|"
Data = Data & Format(QSales, "###") & "|"
Data = Data & Format(TSales, "0.00") & "|"
Data = Data & "" & ";"
ReportView.View.TableBorder = tbTopBottom
ReportView.View.addTable F, H, Data, vbBlack, vbWhite
FinishJob 1
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Mix Reports")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub SRC_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        DoRecap
        Set SearchDates = Nothing
    End If
End If
End Sub
Private Sub DoRecap()
Dim TSold           As Currency
Dim QSold           As Integer
Dim Printall        As Integer
Dim Cdb As ADODB.Connection, Ctb As ADODB.Recordset, Heading As String
Dim PrinterPort     As String, PO As Currency
Dim RecordNo        As Integer
Dim CatList         As New Collection
Dim catdb           As ADODB.Connection
Dim CatTb           As ADODB.Recordset
Dim coupons()       As Variant
Dim LL              As Integer
Dim NSR             As Currency
Dim CCTip           As Currency
Dim TotalCharges    As Currency
Dim CollectNSR      As Currency
StartReport "Daily Sales Recap"
PageHeader = "Daily Sales Recap"
With ReportView.View
   .PenColor = &HC0C0C0
   .PenWidth = 200
   .PenStyle = psSolid
   .CurrentY = "2.75in"
   .DrawLine .CurrentX, .CurrentY, .PageWidth - .MarginRight, .CurrentY
   .CurrentX = "1.25in"
   .CurrentY = .CurrentY - 100
   .FontSize = 10
   .textColor = vbBlack
   .Text = "Daily Sales Recap" & vbCrLf
   .CurrentX = "1.25in"
   .Text = "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate & vbCrLf
   .PenColor = vbBlack
   .PenWidth = 0
End With
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Cdb
    Sql = "Select [Name], Sum([Amount]) as [Price],Sum([Quantity]) as TotalQty from [Department Sales] Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(SearchDates.Item(1).StartDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " and '" & Format(CDate(CDate(SearchDates.Item(1).EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " and [Amount] <> 0"
    Sql = Sql & " and [Name] Not In ('Gift Certificate','Gratuity','Service Charge','Paid In','Gift EPay New Issue','Gift EPay ReLoad','Account Payment','Account Deposit','Non Sales Discount','Discount','Coupon','Online Tip','Frequent Diner','Donation','Online Fee','Online Delivery')"
    Sql = Sql & " Group By [Name] "
    ConnectADORS Sql, Cdb, Ctb
DoLineBox "1in", "3in", "<2.75in|^0.75in|^.25in|^1.75in;", "DEPARTMENT SALES|QTY|  |DOLLARS;", ""
ReportView.View.CurrentY = "3.20in"
If Not Ctb.EOF Then
  Do While Not Ctb.EOF
      PrintLine "", SetNxtPos(1)
      PrintLine Ctb![Name] & "", SetNxtPos(4#)
      PrintLine Format(Ctb![TotalQty], "###"), GetCurX + SetNxtPos(0.75)
      PrintLine "", GetCurX + SetNxtPos(0.55)
      PrintLine Format(Ctb![Price], "0.00"), GetCurX + SetNxtPos(1.25)
      TSold = TSold + Ctb![Price]
      PrintLine "", -1
    Ctb.MoveNext
  Loop
End If
SetNothing Ctb
DrawLine
GetCurX
PrintLine "TOTAL ", 7632
PrintLine Format(TSold, "0.00"), -1
PrintLine "", -1
TSold = 0
PrintLine "PAYMENTS (Cash Not Included In Total)", 5760
PrintLine "QTY", 7632
PrintLine "DOLLARS", 9432
PrintLine "TIPS", -1
DrawLine
    Sql = "Select [Name], Sum([Amount]) as TotalAMT , Sum([Quantity]) as TotalQty, sum([Tip total]) as TotalTip from [Tenders]"
    Sql = Sql & " Where TRY_CAST([date] AS datetime) between '" & Format(CDate(SearchDates.Item(1).StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & "Group by [Name] Order by [Name] asc "
    ConnectADORS Sql, Cdb, Ctb
TSold = 0: CCTip = 0
If Not Ctb.EOF Then
  Do While Not Ctb.EOF
     PrintLine Ctb![Name], 5760
     PrintLine Ctb!TotalQty, 7632
     PrintLine Format(Ctb!TotalAmt, "0.00"), 9432
     PrintLine Format(Ctb!TotalTip, "0.00"), -1
     If Ctb![Name] <> "Cash Tenders" Then
      TSold = TSold + Val(Check_Null(Ctb!TotalAmt))
     End If
     CCTip = CCTip + Val(Check_Null(Ctb!TotalTip))
   Ctb.MoveNext
  Loop
    SetNothing Ctb
    DrawLine
    PrintLine "TOTAL ", 7632
    PrintLine Format(TSold, "0.00"), -1
    TotalCharges = (TSold)
    'CCTip = GetCCTips
    If CCTip > 0 Then
     PrintLine "TOTAL CREDIT CARD TIPS  ", 7632
     PrintLine Format(CCTip, "0.00"), -1
     PrintLine "(included in total)", -1
    End If
    PrintLine "", -1
End If
    Sql = "Select [Name], Sum([Amount]) as TotalAMT , Sum([Quantity]) as TotalQty from [Discounts]"
    Sql = Sql & " Where TRY_CAST([date] AS datetime) between '" & Format(CDate(SearchDates.Item(1).StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & "Group by [Name] Order by [Name] asc "
    ConnectADORS Sql, Cdb, Ctb
If Not Ctb.EOF Then
   PrintLine "", -1
   PrintLine "DISCOUNTS", 5760
   PrintLine "QTY", 7632
   PrintLine "DOLLARS", -1
   DrawLine
   TSold = 0
    Do While Not Ctb.EOF
        PrintLine Ctb![Name], 5760
        PrintLine Format(Ctb![TotalQty], "###"), 7632
        PrintLine Format(Ctb![TotalAmt], "0.00"), -1
        TSold = TSold + Ctb![TotalAmt]
     Ctb.MoveNext
    Loop
   SetNothing Ctb
  DrawLine
  PrintLine "TOTAL ", 7632
  PrintLine Format(TSold, "0.00"), -1
  PrintLine "", -1
End If
  TSold = 0
  
    Sql = "Select [Name], Sum([Amount]) as TotalAMT , Sum([Quantity]) as TotalQty from [Paid Outs]"
    Sql = Sql & " Where TRY_CAST([date] AS datetime) between '" & Format(CDate(SearchDates.Item(1).StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & "Group by [Name] Order by [Name] asc "
  ConnectADORS Sql, Cdb, Ctb
  If Not Ctb.EOF Then
    PrintLine "PAID OUTS", 7632
    PrintLine "DOLLARS", -1
    DrawLine
    TSold = 0
    Do While Not Ctb.EOF
          PrintLine Ctb![Name], 7632
          PrintLine Format(Val(Check_Null(Ctb!TotalAmt)), "0.00"), -1
          TSold = TSold + Val(Check_Null(Ctb!TotalAmt))
      Ctb.MoveNext
    Loop
    SetNothing Ctb
    DrawLine
    PrintLine "TOTAL ", 7632
    PrintLine Format(TSold, "0.00"), -1
    PO = TSold
    PrintLine "", -1
  End If

'End If
    Sql = "Select Sum([Exempt Sales]) as ESale, Sum([Charged Gratuity]) as TGrat"
    Sql = Sql & " ,Sum([Tenders]) as TTender, Sum([Paid Outs]) as TPaidOut, Sum([Deposits]) as TDeposits"
    Sql = Sql & " ,Sum([Net Sales]) as TNet ,Sum([Sales Tax]) as Stax, Sum([Special Tax]) as SpTax"
    Sql = Sql & " ,Sum([Liquor Tax]) as LTax, Sum([Deposits Collected])as TCollected, Sum([House Charges]) as HCharge"
    Sql = Sql & " ,Sum([Tip Charge Back]) as TCB"
    Sql = Sql & " ,Sum([Refunds]) as RFund"
    Sql = Sql & " ,Sum([Gross Sales]) as GSales"
    Sql = Sql & " ,Sum([Voids]) as VSales"
    Sql = Sql & " from [Sales] Where "
    Sql = Sql & " TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDates.Item(1).StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SearchDates.Item(1).EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  ConnectADORS Sql, Cdb, Ctb
If Not Ctb.EOF Then
    PrintLine "SALES", 7632
    PrintLine "DOLLARS", -1
    DrawLine
    PrintLine "Gross Sales  ", 7632: PrintLine Format(Ctb![GSales], "0.00"), -1
    PrintLine "Voided Sales ", 7632: PrintLine Format(Ctb![VSales], "0.00"), -1
    PrintLine "Sales Tax    ", 7632: PrintLine Format(Ctb![STAX], "0.00"), -1
    PrintLine "Special Tax  ", 7632: PrintLine Format(Ctb![SPTax], "0.00"), -1
    PrintLine SurChargeTxt, 7632: PrintLine Format(Ctb![LTax], "0.00"), -1
    If Input_Option(58) = "Yes" Then
      PrintLine "Non Sale Revenue", 7632: PrintLine Format(Val(Ctb![ESale] & ""), "0.00"), -1
      CollectNSR = Ctb![ESale]
    Else
      PrintLine "Non Sale Revenue", 7632: PrintLine Format(Ctb![ESale] - Ctb!TGrat, "0.00"), -1
      CollectNSR = Val(Ctb![ESale] & "") - Val(Ctb.fields("TGrat").Value & "")
    End If
    PrintLine "Refunds      ", 7632: PrintLine Format(Ctb![RFund], "0.00"), -1
    PrintLine "Net Sales    ", 7632: PrintLine Format((Ctb![TNet]), "0.00"), -1  '- Ctb![RFund]
    CashtoAcctFor = (Ctb![TNet] + (Ctb![STAX] + Ctb![SPTax]) + Ctb![TCB] + _
    (Ctb![LTax])) + CollectNSR
    DrawLine
    PrintLine "TOTAL TO ACCOUNT FOR ", 7632: PrintLine Format(CashtoAcctFor, "0.00"), -1
    PrintLine "CASH TO ACCOUNT FOR  ", 7632: PrintLine Format((CashtoAcctFor) - (TotalCharges + PO), "0.00"), -1
    If Check_Option(237) = 1 Then
      CashtoAcctFor = CashtoAcctFor + CCTip
    End If
    PrintLine "CASH INCLUDING (tips withheld) ", 7632: PrintLine Format(CashtoAcctFor, "0.00"), -1
    If CCTip > 0 Then
     PrintLine "TOTAL CREDIT CARD TIPS  ", 7632: PrintLine Format(CCTip, "0.00"), -1
    End If
    DrawLine
End If


ReportView.Caption = "Sales Recap Sheet"
FinishJob 1
End Sub
Private Function GetCCTips() As Currency
On Error GoTo ErrorHandler
Dim SL As ADODB.Connection, SRd As ADODB.Recordset
OpenDb SvrPath.TerminalPath & "MSaleData", SL
      Sql = "Select Sum([Tip Amount]) as TTip from [Server Payments] "
      Sql = Sql & " Where TRY_CAST([Date] AS date) Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
      Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
      'Sql = Sql & " and [Payment Type] = 'Credit Card'"
ConnectADORS Sql, SL, SRd
If Not SRd.EOF Then
 GetCCTips = Format(SRd![TTip], "0.00")
Else
 GetCCTips = 0
End If
DbClose SL, SRd
ErrorHandler:

End Function

Private Sub SSCommand1_Click()
On Error GoTo ErrorHandler
ReportView.View.Orientation = orLandscape
SetPage ReportView.View, "Weekly Summary Report"
Dim Pa As Integer, Cl As Long
ReportView.View.RenderControl = SummarySheet.hwnd
ReportView.View.EndDoc
ReportView.Show 1
Call MemoryCheck(ReportView)
'end with
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Display SS1 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

Exit Sub
End Sub
Public Sub ShiftWeeks(MyDate As String, WeekMove As String)
On Error GoTo ErrorHandler
Dim LP As Variant
Dim OverSales As Currency
If Not IsDate(MyDate) Then
 Select Case Scroll
    Case Is = -1
       wYear = wYear - 1
       Scroll = 53
    Case Is = 53
       wYear = wYear + 1
       Scroll = 1
 End Select
 If Not IsDate(MyDate) Then
    If IsDate(SummarySheet.TextMatrix(0, 1)) Then
     If WeekMove = "+" Then
      MyDate = CDate(SummarySheet.TextMatrix(0, 1)) + 7
     Else
      MyDate = CDate(SummarySheet.TextMatrix(0, 1)) - 7
     End If
    End If
 End If
 If wYear = 0 Then
    wYear = Year(HistoryDate.Text)
    If wYear = 0 Then wYear = Year(Date)
 End If
End If
SummarySheet.Clear
Dim RS           As ADODB.Recordset
Dim Cdb          As ADODB.Connection
Dim PDB          As ADODB.Connection
Dim StartingDate As String
Dim NODays       As Integer
SearchY.Caption = "Searching Year " & wYear
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Cdb) = False Then Exit Sub
If ConnectADODB(SvrPath.TerminalPath & ProductHistory, PDB) = False Then Exit Sub
If Not IsDate(MyDate) Then
   Sql = "Select * from [Sales] Where datepart('ww',cdate([date])," & WeekStartsOn & ", 3) = " & Scroll & " and year([date]) = " & wYear & " order by cdate([date]) desc"
  If ConnectADORS(Sql, Cdb, RS) = False Then Exit Sub
  If Not RS.EOF Then
   StartingDate = RS![Date]
  End If
End If
If IsDate(MyDate) Then
  StartingDate = MyDate
  NODays = 7 - WeekDay(CDate(MyDate), WeekStartsOn)
Else
  NODays = 7
End If
If Not IsDate(StartingDate) Then
   StartingDate = Date
   See_Msg "Date Not Found", 5, "For This Week"
End If
Sql = "Select * from [Sales] Where TRY_CAST([date] AS datetime) between '" & Format(CDate(CDate(StartingDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(StartingDate) + NODays), "yyyy-mm-dd hh:nn:ss") & "' order by TRY_CAST([Date] AS datetime) asc"
If ConnectADORS(Sql, Cdb, RS) = False Then Exit Sub
With SummarySheet
  .TextMatrix(1, 0) = "Gross Sales"
  .TextMatrix(2, 0) = "Net Sales"
  .TextMatrix(3, 0) = "Sales Tax"
  .TextMatrix(4, 0) = "Special Tax"
  .TextMatrix(5, 0) = SurChargeTxt
  .TextMatrix(6, 0) = "Voids"
  .TextMatrix(7, 0) = "Discounts"
  .TextMatrix(8, 0) = "Non Cash Tenders"
  .TextMatrix(9, 0) = "Paid Outs"
  .TextMatrix(10, 0) = "Bank Deposits"
  .TextMatrix(11, 0) = "Non Taxed Sales"
  .TextMatrix(12, 0) = "Ttl To Acct for"
  .TextMatrix(13, 0) = "Ttl Accounted for"
  .TextMatrix(14, 0) = "Over / Short"
  .TextMatrix(15, 0) = "Labor Dollars"
  .TextMatrix(16, 0) = "Labor Hours"
  .TextMatrix(17, 0) = "Labor Percentage"
  .TextMatrix(18, 0) = "Sales Percentage"
  .TextMatrix(19, 0) = "Go Sales"
  .TextMatrix(20, 0) = "Call In Sales"
  .TextMatrix(21, 0) = "Drive Thru Sales"
  .TextMatrix(22, 0) = "Gift Cert/Paid In"
  .TextMatrix(23, 0) = "Tip Fee Income"
  .TextMatrix(24, 0) = "Refunds"
  .TextMatrix(25, 0) = "Deposit Paid In"
  If Check_Option(237) = 1 Then
  .TextMatrix(26, 0) = "Tips Withheld"
  End If
End With
Dim DayInfo() As Currency, WTDsales As Currency
SummarySheet.TextMatrix(0, 0) = "Summary"
If Not RS.EOF Then
 SearchDate = Format(RS![Date], "m/d/yyyy")
 LocName.Caption = "Week of " & Format(SearchDate, "MM/DD/YYYY")
Else
 LocName.Caption = "No Data Found For Week # " & Scroll
End If
ReDim DayInfo(26, 8)
  WTDsales = 0
    If Not RS.EOF Then
     LocName.Caption = "Week of " & Format(SearchDate, "MM/DD/YYYY")
     StartingDate = RS![Date]
    End If
  Do While Not RS.EOF
    If Not IsDate(StartingDate) Then StartingDate = RS![Date]
     If RS![Date] & "" > "" And CDate(StartingDate) <= CDate(RS![Date]) Then
        If RS![Memo] & "" > "" Then
         SummarySheet.Cell(flexcpForeColor, 0, WeekDay(CDate(RS![Date]), WeekStartsOn), 0, WeekDay(CDate(RS![Date]), WeekStartsOn)) = vbRed
        Else
         SummarySheet.Cell(flexcpForeColor, 0, WeekDay(CDate(RS![Date]), WeekStartsOn), 0, WeekDay(CDate(RS![Date]), WeekStartsOn)) = vbBlack
        End If
        SummarySheet.TextMatrix(0, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Format(RS![Date], "  mm/dd/yy")
        DayInfo(0, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Gross Sales])))
        DayInfo(1, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Net Sales])))
        WTDsales = WTDsales + Val(RS![Net Sales] & "")
        DayInfo(2, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Sales Tax])))
        DayInfo(3, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Special Tax])))
        DayInfo(4, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Liquor Tax])))
        DayInfo(5, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Voids])))
        DayInfo(6, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Discounts])))
        '[NonCashTenders
        'DayInfo(7, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Non_CashTenders(pDb, RS.fields("Date").Value, RS.fields("Date").Value)
        DayInfo(7, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Tenders])))
        DayInfo(8, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Paid Outs])))
        DayInfo(9, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Deposits])))
        DayInfo(10, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![None Taxed Sales])))
        If Check_Option(237) = 1 Then
        DayInfo(25, WeekDay(CDate(RS![Date]), WeekStartsOn)) = GetTipwitheld(RS![Date])
        Else
        DayInfo(25, WeekDay(CDate(RS![Date]), WeekStartsOn)) = 0
        End If
        OverSales = Val(RS![Exempt Sales] & "") - Val(RS![Charged Gratuity] & "")
        If OverSales < 0 Then OverSales = 0
        DayInfo(12, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(Check_Null(RS![Tenders])) + Val(Check_Null(RS![Paid Outs])) + Val(Check_Null(RS![Deposits])) - (Val(RS![Deposits Collected] & "") + Val(RS![House Charges] & "") + Val(RS![Tip Charge Back] & ""))
        DayInfo(11, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(Check_Null(RS![Net Sales])) + Val(Check_Null(RS![Sales Tax])) + Val(Check_Null(RS![Special Tax])) + Val(Check_Null(RS![Liquor Tax])) + OverSales
        'OverSales is Amount of Exempt sales less gratuity charged
        'will cause over short to be right other wise will be over
        'amount of exempt sales.
        DayInfo(13, WeekDay(CDate(RS![Date]), WeekStartsOn)) = (Val(DayInfo(12, WeekDay(CDate(RS![Date]), WeekStartsOn))) - (Val(DayInfo(11, WeekDay(CDate(RS![Date]), WeekStartsOn))) + DayInfo(25, WeekDay(CDate(RS![Date]), WeekStartsOn)))) '- OverSales
        DayInfo(14, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Labor Dollars])))
        If Val(RS![Net Sales] & "") > 0 And Val(Check_Null(RS![Labor Dollars])) > 0 Then
          LP = Format(Val(Check_Null(RS![Labor Dollars])) / Val(RS![Net Sales] & "") * 100, "0.00")
        End If
        DayInfo(15, WeekDay(CDate(RS![Date]), WeekStartsOn)) = CStr(Val(Check_Null(RS![Labor Hours])))
        DayInfo(16, WeekDay(CDate(RS![Date]), WeekStartsOn)) = LP
        DayInfo(18, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(RS![Go Sales] & "")
        DayInfo(19, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(RS![Call In Sales] & "")
        DayInfo(20, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(RS![Drive Thru Sales] & "")
        DayInfo(21, WeekDay(CDate(RS![Date]), WeekStartsOn)) = OverSales   'Val(Rs![Exempt Sales] & "")
        DayInfo(22, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(RS![Tip Charge Back] & "")
        DayInfo(23, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(RS![Refunds] & "")
        DayInfo(24, WeekDay(CDate(RS![Date]), WeekStartsOn)) = Val(RS![Deposits Collected] & "")
        
    End If
  RS.MoveNext
  Loop
DbClose Cdb, RS
DbClose PDB, Nothing
Dim FI As Integer, DN As Integer
Dim LabDoll As Currency
LabDoll = 0
For DN = 1 To 7
    For FI = 0 To 16
            If FI = 5 Then
            ' If Val(DayInfo(5, DN)) >= Val(Input_Option(78)) Then
            '  SummarySheet.Cell(flexcpForeColor, 6, DN, 6, DN) = vbRed
            ' Else
              SummarySheet.Cell(flexcpForeColor, 6, DN, 6, DN) = vbBlack
            ' End If
            End If
            If FI = 6 Then
             '   If Val(DayInfo(6, DN)) >= Val(Input_Option(79)) Then
             '     SummarySheet.Cell(flexcpForeColor, 7, DN, 7, DN) = vbRed
             '   Else
                  SummarySheet.Cell(flexcpForeColor, 7, DN, 7, DN) = vbBlack
             '   End If
            End If
            If FI = 16 Then
              '  If Val(DayInfo(16, DN)) >= Val(Input_Option(80)) Then
               '   SummarySheet.Cell(flexcpForeColor, 17, DN, 17, DN) = vbRed
                'Else
                  SummarySheet.Cell(flexcpForeColor, 17, DN, 17, DN) = vbBlack
                'End If
            End If
            SummarySheet.TextMatrix(FI + 1, DN) = Format(DayInfo(FI, DN), "0.00")
            If FI = 14 Then LabDoll = LabDoll + DayInfo(14, DN)
    Next FI
    For FI = 18 To 25
            SummarySheet.TextMatrix(FI + 1, DN) = Format(DayInfo(FI, DN), "0.00")
    Next FI
Next DN
Dim LineTotal As Currency
SummarySheet.TextMatrix(0, 8) = "     Total"
For FI = 0 To 16
LineTotal = 0
For DN = 1 To 7
  LineTotal = LineTotal + Val(DayInfo(FI, DN))
Next DN
If FI = 16 Then
    If LabDoll > 0 And Val(SummarySheet.TextMatrix(2, 8)) > 0 Then SummarySheet.TextMatrix(FI + 1, 8) = Format((LabDoll / SummarySheet.TextMatrix(2, 8)) * 100, "0.00")
Else
    SummarySheet.TextMatrix(FI + 1, 8) = Format(LineTotal, "0.00")
End If
Next FI
For FI = 18 To 25
LineTotal = 0
For DN = 1 To 7
  LineTotal = LineTotal + Val(DayInfo(FI, DN))
Next DN
SummarySheet.TextMatrix(FI + 1, 8) = Format(LineTotal, "0.00")
Next FI
With SummarySheet
For DN = 1 To 7
  If WTDsales > 0 And Val(DayInfo(1, DN)) > 0 Then
  .TextMatrix(18, DN) = Format((Val(DayInfo(1, DN)) / WTDsales) * 100, "0.00")
  .TextMatrix(18, 8) = Format(Val(.TextMatrix(18, 8)) + Val(.TextMatrix(18, DN)), "0.00")
  End If
Next DN
End With
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Shift Weeks ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub CompareSales(SelectedDate As Boolean)
On Error GoTo ErrorHandler
Dim LP As Variant
Dim Cf As Long
Dim OverSales As Currency
Dim GDay As Integer
Dim Db As ADODB.Connection
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
'If SearchSpan = -1 Then Exit Sub
If Trim(Compare.TextMatrix(0, 1)) = "Total" Then Exit Sub
wYear = Year(HistoryDate.Text)
If CompareByTotal.Text = "Week" Then
   Compare.TextMatrix(0, 1) = SummarySheet.TextMatrix(0, 1)
   GDay = 7
Else
   GDay = 0
End If
If wYear = 0 Then wYear = Year(Date)
OpenDb SvrPath.TerminalPath & ClosingHistory, Cdb
If SelectedDate = False Then
If Compare.TextMatrix(0, 1) = "" Then
   Call See_Msg("No Date Found", 5, "")
   DbClose Cdb, RS
   Exit Sub
End If
If GDay > 0 Then
  Sql = "Select * from [Sales] Where datepart('ww',cdate([date])," & WeekStartsOn & ",3) = " & DatePart("ww", CDate(Compare.TextMatrix(0, 1)) + SearchSpan, WeekStartsOn, 3) & " and year([date]) = '" & wYear & "'"
Else
'  Sql = "Select * from [Sales] Where TRY_CAST([date] AS datetime) = '" & CDate(Compare.TextMatrix(0, 1)) + SearchSpan & "' and year([date]) = '" & wYear & "' and not isnull([date])"
  Sql = "Select * from [Sales] Where TRY_CAST([date] AS datetime) = '" & CDate(Compare.TextMatrix(0, 1)) + SearchSpan & "'"

End If
Else
If SearchHow.Text = "Week" Then
  Sql = "Select * from [Sales] Where datepart('ww',cdate([date])," & WeekStartsOn & ",1) = " & DatePart("ww", CDate(Compare.TextMatrix(0, 2)), WeekStartsOn, 3) & " and year([date]) = '" & wYear & "'"
Else
  'Sql = "Select * from [Sales] Where TRY_CAST([date] AS datetime) = '" & CDate(Compare.TextMatrix(0, 2)) & "' and year([date]) = '" & wYear & "' and not isnull([date])"
  Sql = "Select * from [Sales] Where TRY_CAST([date] AS datetime) = '" & CDate(Compare.TextMatrix(0, 2)) & "'"

End If
End If
ConnectADORS Sql, Cdb, RS
Dim DayInfo() As Currency, WTDsales As Currency
If SelectedDate = True Then
  CompareDisp.Caption = Get_DayofWeek(DatePart("w", CDate(Compare.TextMatrix(0, 1)))) & " / " & Get_DayofWeek(DatePart("w", CDate(Compare.TextMatrix(0, 2))))
Else
  CompareDisp.Caption = Get_DayofWeek(DatePart("w", CDate(Compare.TextMatrix(0, 1)))) & " / " & Get_DayofWeek(DatePart("w", CDate(Compare.TextMatrix(0, 1)) + SearchSpan))
End If
ReDim DayInfo(21, 8)
  WTDsales = 0
  If Not RS.EOF Then
      If SelectedDate = False Then
        Compare.TextMatrix(0, 2) = Format(CDate(Compare.TextMatrix(0, 1)) + SearchSpan, "MM/DD/YY")
      End If
  Else
      If SelectedDate = False Then
        Compare.TextMatrix(0, 2) = Format(CDate(Compare.TextMatrix(0, 1)) + SearchSpan, "MM/DD/YY")
      End If
  End If
    With Compare
      For Cf = 1 To 21
        .TextMatrix(Cf, 2) = ""
      Next Cf
    End With
 
  Do While Not RS.EOF
    With Compare
    .TextMatrix(1, 2) = Val(.TextMatrix(1, 2)) + Format(CStr(Val(Check_Null(RS![Gross Sales]))), "0.00")
    .TextMatrix(2, 2) = Val(.TextMatrix(2, 2)) + Format(CStr(Val(Check_Null(RS![Net Sales]))), "0.00")
    .TextMatrix(3, 2) = Val(.TextMatrix(3, 2)) + Format(CStr(Val(Check_Null(RS![Sales Tax]))), "0.00")
    .TextMatrix(4, 2) = Val(.TextMatrix(4, 2)) + Format(CStr(Val(Check_Null(RS![Special Tax]))), "0.00")
    .TextMatrix(5, 2) = Val(.TextMatrix(5, 2)) + Format(CStr(Val(Check_Null(RS![Liquor Tax]))), "0.00")
    .TextMatrix(6, 2) = Val(.TextMatrix(6, 2)) + Format(CStr(Val(Check_Null(RS![Voids]))), "0.00")
    .TextMatrix(7, 2) = Val(.TextMatrix(7, 2)) + Format(CStr(Val(Check_Null(RS![Discounts]))), "0.00")
    .TextMatrix(8, 2) = Val(.TextMatrix(8, 2)) + Format(CStr(Val(Check_Null(RS![Tenders]))), "0.00")
    .TextMatrix(9, 2) = Val(.TextMatrix(9, 2)) + Format(CStr(Val(Check_Null(RS![Paid Outs]))), "0.00")
    .TextMatrix(10, 2) = Val(.TextMatrix(10, 2)) + Format(CStr(Val(Check_Null(RS![Deposits]))), "0.00")
    .TextMatrix(11, 2) = Val(.TextMatrix(11, 2)) + Format(CStr(Val(Check_Null(RS![None Taxed Sales]))), "0.00")
    OverSales = Val(RS![Exempt Sales] & "") - Val(RS![Charged Gratuity] & "")
    If OverSales < 0 Then OverSales = 0
    .TextMatrix(12, 2) = Val(.TextMatrix(12, 2)) + Format(Val(Check_Null(RS![Tenders])) + Val(Check_Null(RS![Paid Outs])) + Val(Check_Null(RS![Deposits])), "0.00")
    .TextMatrix(13, 2) = Val(.TextMatrix(13, 2)) + Format(Val(Check_Null(RS![Net Sales])) + Val(Check_Null(RS![Sales Tax])) + Val(Check_Null(RS![Special Tax])) + Val(Check_Null(RS![Liquor Tax])) + OverSales, "0.00")
    .TextMatrix(14, 2) = Val(.TextMatrix(14, 2)) + Format((Val(DayInfo(12, WeekDay(CDate(RS![Date]), WeekStartsOn))) - Val(DayInfo(11, WeekDay(CDate(RS![Date]), WeekStartsOn)))), "0.00") '- OverSales
    .TextMatrix(15, 2) = Val(.TextMatrix(15, 2)) + Format(CStr(Val(Check_Null(RS![Labor Dollars]))), "0.00")
    If Val(RS![Net Sales] & "") > 0 And Val(Check_Null(RS![Labor Dollars])) > 0 Then
      LP = Format(Val(Check_Null(RS![Labor Dollars])) / Val(RS![Net Sales] & "") * 100, "0.00")
    End If
    .TextMatrix(16, 2) = Val(.TextMatrix(16, 2)) + Format(CStr(Val(Check_Null(RS![Labor Hours]))), "0.00")
    .TextMatrix(17, 2) = Val(.TextMatrix(17, 2)) + LP
    .TextMatrix(18, 2) = Val(.TextMatrix(18, 2)) + Format(Val(RS![Go Sales] & ""), "0.00")
    .TextMatrix(19, 2) = Val(.TextMatrix(19, 2)) + Format(Val(RS![Call In Sales] & ""), "0.00")
    .TextMatrix(20, 2) = Val(.TextMatrix(20, 2)) + Format(Val(RS![Drive Thru Sales] & ""), "0.00")
    .TextMatrix(21, 2) = Val(.TextMatrix(21, 2)) + Format(OverSales, "0.00") 'Val(Rs![Exempt Sales] & "")
    End With
    RS.MoveNext
 Loop
DbClose Cdb, RS
For Cf = 1 To 21
  With Compare
     .TextMatrix(Cf, 3) = Format(Val(.TextMatrix(Cf, 1)) - Val(.TextMatrix(Cf, 2)), "0.00")
     If Val(.TextMatrix(Cf, 1)) > 0 And Abs(Val(.TextMatrix(Cf, 3))) > 0 Then
        .TextMatrix(Cf, 4) = Format(((.TextMatrix(Cf, 3)) / .TextMatrix(Cf, 1) * 100), "0.00")
     Else
        .TextMatrix(Cf, 4) = "0.00"
     End If
     If Val(.TextMatrix(Cf, 4)) < 0 Then
       .Cell(flexcpForeColor, Cf, 4, Cf, 4) = vbRed
     Else
       .Cell(flexcpForeColor, Cf, 4, Cf, 4) = vbBlack
     End If
  End With
Next Cf
SelectedDate = False
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Compare Sales ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand10_Click()
Select Case SearchHow.Text
   Case Is = "Day"
     AdvSearch = 1
   Case Is = "Week"
     AdvSearch = 7
   Case Is = "Year"
     AdvSearch = 365
End Select
SearchSpan = SearchSpan - AdvSearch
CompareSales False
End Sub
Private Sub SSCommand11_Click()
Select Case SearchHow.Text
   Case Is = "Day"
     AdvSearch = 1
   Case Is = "Week"
     AdvSearch = 7
   Case Is = "Year"
     AdvSearch = 365
End Select
SearchSpan = SearchSpan + AdvSearch
CompareSales False

End Sub
Private Sub SSCommand13_Click()
Dim Pl As Integer
SetPage ReportView.View, "Sales Review"
PrintLine CompareDisp.Caption, 3000
With Compare
DrawLine
 PrintLine CompareDisp.Caption, 3000
 PrintLine .TextMatrix(0, 1), 4500
 PrintLine .TextMatrix(0, 2), 6000
 PrintLine .TextMatrix(0, 3), 7500
 PrintLine .TextMatrix(0, 4), -1
DrawLine
For Pl = 1 To 21
 PrintLine .TextMatrix(Pl, 0), 3000
 PrintLine .TextMatrix(Pl, 1), 4500
 PrintLine .TextMatrix(Pl, 2), 6000
 PrintLine .TextMatrix(Pl, 3), 7500
 PrintLine .TextMatrix(Pl, 4), -1
Next Pl
End With
DrawLine
FinishJob 1
End Sub

Private Sub SSCommand14_Click()
Memo.Visible = False
Memo.Left = -4000
End Sub

Private Sub SSCommand15_Click()
If Not IsDate(MDate) Then Exit Sub
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
OpenDb SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select * from [Sales] Where TRY_CAST([Date] AS datetime) = '" & CDate(MDate) & "'"
OpenTb Sql, RS, Db
 If RS Is Nothing Then
    DbClose Db, RS
    Exit Sub
 End If
If Not RS.EOF Then
  RS![Memo] = SumMemo.Text & ""
  RS.Update
End If
DbClose Db, RS
See_Msg "Saved...", 5, ""
End Sub
Private Sub SSCommand2_Click()
Scroll = Scroll + 1
'If Scroll > 52 Then Scroll = 52
ShiftWeeks "", "+"
End Sub

Private Sub SSCommand23_Click()
With ServerList
   .Left = -4000
  .Visible = False
End With

End Sub

Private Sub SSCommand3_Click()
Scroll = Scroll - 1
'If Scroll < 1 Then Scroll = 1
ShiftWeeks "", "-"
End Sub

Private Sub SSCommand4_Click()
If Trim(SummarySheet.TextMatrix(0, SummarySheet.Col)) > "" Then
  RePrintSales SummarySheet.TextMatrix(0, SummarySheet.Col)
Else
  Call See_Msg("Touch Date First", 5, "")
End If
End Sub

Private Sub SSCommand5_Click()
If Trim(StoreDate.Text) = "" Then Exit Sub
If Label3.Caption = "Change Deposit" Then
    Dim OS As Currency
    Dim Cdb As ADODB.Connection, RS As ADODB.Recordset
    If IsDate(SummarySheet.TextMatrix(0, SummarySheet.Col)) = False Then Exit Sub
    OpenDb SvrPath.TerminalPath & ClosingHistory, Cdb
    Sql = "Select * from [Sales] where TRY_CAST([Date] AS datetime) = '" & CDate(SummarySheet.TextMatrix(0, SummarySheet.Col)) & "'"
    ConnectADORS Sql, Cdb, RS
     If Not RS.EOF Then
       SummarySheet.TextMatrix(10, SummarySheet.Col) = Format(StoreDate.Text, "0.00")
       OS = (Val(SummarySheet.TextMatrix(8, SummarySheet.Col)) + Val(SummarySheet.TextMatrix(9, SummarySheet.Col)) + Val(SummarySheet.TextMatrix(10, SummarySheet.Col))) - Val(SummarySheet.TextMatrix(12, SummarySheet.Col))
       SummarySheet.TextMatrix(14, SummarySheet.Col) = Format(OS, "0.00")
       RS![Deposits] = Format(StoreDate.Text, "0.00")
       RS.fields("Over Short").Value = Format(OS, "0.0")
       RS.Update
    End If
    DbClose Cdb, RS
Else
    With SummarySheet
      .TextMatrix(0, .Col) = Format(StoreDate.Text, "MM/DD/YYYY")
    End With
End If
SSCommand6_Click
End Sub

Private Sub SSCommand6_Click()
NewDate.Visible = False
NewDate.Left = -4000
End Sub

Private Sub SSCommand7_Click()

End Sub

Private Sub SSCommand8_Click()
End Sub
Private Sub SSCommand9_Click()
CompareHistory.Visible = False
CompareHistory.Left = -5000
End Sub
Public Sub DateChange(OldDate As String, NewDate As String)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & ClosingHistory, Db
    Sql = "Update [Category Sales] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Department Sales] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Deposits] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Discounts] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Hourly Sales] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Labor] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Meal Periods] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Non Sales Items] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Order Destination] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Paid Outs] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Sales] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Sales Tax] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
    Sql = "Update [Tenders] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute Sql
SetNothing Db
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
    Sql = "Update [Detail] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute ReturnDateString(Sql)
    Sql = "Update [Voids] Set [Date] = '" & NewDate & "'"
    Sql = Sql & " where TRY_CAST([Date] AS datetime) = '" & Format(CDate(CDate(OldDate)), "yyyy-mm-dd hh:nn:ss") & "'"
 Db.Execute ReturnDateString(Sql)
SetNothing Db
See_Msg "Update Complete", 5, ""
End Sub

Private Sub SummarySheet_Click()
On Error GoTo ErrorHandler
'Start Edit Deposit********************************************
If SummarySheet.RowSel = 10 And IsDate(SummarySheet.TextMatrix(0, SummarySheet.Col)) Then
   EditItem = "Deposits"
   EditDayData SummarySheet.TextMatrix(0, SummarySheet.Col)
   Exit Sub
End If

'edit daily Cash deposit.**************************************

If ShowEdit.Visible And SummarySheet.RowSel <> 49 Then
 ShowEdit.Visible = False
 EditRecord = False
 Exit Sub
End If
If SummarySheet.TextMatrix(0, SummarySheet.Col) = "" Then
    If SummarySheet.Row = 1 Then
    'EditDayData SearchDate
    CenterObject Me, NewDate
    Label3.Caption = "Enter Date"
    KeyCal.Caption = "Show Keyboard"
    StoreDate.Text = ""
    NewDate.ZOrder 0
    NewDate.Visible = True
    Exit Sub
    End If
End If
Dim Data()
Dim Labels()
Dim Colors()

If SummarySheet.RowSel = 15 Then
  If SummarySheet.Col > 0 And SummarySheet.Col < 9 Then
  With SummarySheet
    Data = Array(Val(.TextMatrix(15, SummarySheet.Col)), Val(.TextMatrix(2, SummarySheet.Col)) - Val(.TextMatrix(15, SummarySheet.Col)))
    Labels = Array("Labor Dollars", "Net Sales")
  End With
    Charts.DisplayGrid.Visible = False
    Charts.SetData Data, Labels, Colors, "Sales / Labor" & " for " & LocName.Caption
    Charts.Show 1
  End If
End If

If SummarySheet.RowSel = 7 Then
  If SummarySheet.Col > 0 And SummarySheet.Col < 9 Then
  With SummarySheet
    Data = Array(Val(.TextMatrix(7, SummarySheet.Col)), Val(.TextMatrix(2, SummarySheet.Col)))
    Labels = Array("Discount Dollars", "Net Sales")
  End With
    Charts.DisplayGrid.Visible = False
    Charts.SetData Data, Labels, Colors, "Sales / Discounts" & " for " & LocName.Caption
    Charts.Show 1
  End If
End If
If SummarySheet.RowSel = 8 Then
  If SummarySheet.Col > 0 And SummarySheet.Col < 9 Then
  With SummarySheet
    Data = Array(Val(.TextMatrix(8, SummarySheet.Col)), Val(.TextMatrix(2, SummarySheet.Col)))
    Labels = Array("Charged Sales", "Cash Sales")
  End With
    Charts.DisplayGrid.Visible = False
    Charts.SetData Data, Labels, Colors, "Sales / Charges" & " for " & LocName.Caption
    Charts.Show 1
  End If
End If
'EditItem = ShowEdit.TextMatrix(ShowEdit.Row, 0)
'SearchDate = SummarySheet.TextMatrix(0, SummarySheet.Col)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Summary Sheet ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub SummarySheet_DblClick()
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
With SummarySheet
MDate = Trim(.TextMatrix(0, .Col))
End With
If IsDate(MDate) Then
OpenDb SvrPath.TerminalPath & ClosingHistory, Db
Sql = "Select * from [Sales] Where TRY_CAST([Date] AS datetime) = '" & CDate(MDate) & "'"
OpenTb Sql, RS, Db
 If RS Is Nothing Then
    DbClose Db, RS
    Exit Sub
 End If
 SumMemo.Text = ""
 If Not RS.EOF Then
   SumMemo.Text = RS![Memo] & ""
 End If
CenterObject Me, Memo
Memo.ZOrder 0
Memo.Visible = True
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Summary Sheet DBL Click ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub SummarySheet_GotFocus()
CC = True
End Sub
Private Sub SummarySheet_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler:
Dim Cter As Integer
Dim RC   As Long
If Button = 1 Then
    Dim RS As ADODB.Recordset
    Dim Cdb As ADODB.Connection
    Dim Data() As Variant
    Dim Labels() As Variant
    Dim Colors() As Variant
    SearchDate = SummarySheet.TextMatrix(0, SummarySheet.Col)
    Dim TTL As Currency
    Dim CT  As Integer
    'Dim Avgs As Currency
    If Not IsDate(SearchDate) Then Exit Sub
    Charts.DisplayGrid.Clear flexClearEverywhere, flexClearText
    If Trim(SearchDate) = "" Then Exit Sub
    ConnectADODB SvrPath.TerminalPath & ClosingHistory, Cdb
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Paid Outs" Then
        Charts.DisplayGrid.cols = 3
        Sql = "Select * from [Paid Outs] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'"
      ConnectADORS Sql, Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        Charts.DisplayGrid.TextMatrix(0, 0) = "Name"
        Charts.DisplayGrid.TextMatrix(0, 1) = "Desc"
        Charts.DisplayGrid.TextMatrix(0, 2) = "$"
        Charts.DisplayGrid.rows = RS.RecordCount + 2
        TTL = 0
        ReDim FieldName(RS.fields.count - 1)
        For Cnt = 0 To RC - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS![Name] & ""
            Labels(Cnt) = RS![Name] & ""
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = RS![Type] & ""
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = RS(2)
            TTL = TTL + Val(RS(2) & "")
            Data(Cnt) = RS(2)
            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = ""
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = "Total"
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = Format(TTL, "0.00")
        Charts.DisplayGrid.ColFormat(2) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1500 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1500 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        SetNothing RS
        SetNothing Cdb
        Charts.SetData Data, Labels, Colors, "Paid Outs" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
'*******************************
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Sales Percentage" Then
        Charts.DisplayGrid.cols = 3
        ConnectADORS "Select * from [Order Destination] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        Charts.DisplayGrid.rows = RC + 1
        Charts.DisplayGrid.cols = RS.fields.count
        ReDim FieldName(RS.fields.count - 1)
        For Cnt = 0 To RS.fields.count - 2
          If OKTOUSE(RS.fields(Cnt + 1).Name) Then
           Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
          End If
        Next
        'Data = ""
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        For Cnt = 0 To RS.RecordCount - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS(1)
            Labels(Cnt) = RS(1)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = RS(2)
            Data(Cnt) = RS(3)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = Format(RS(3), "0.00")
            RS.MoveNext
        Next
        Charts.DisplayGrid.ColFormat(2) = "(#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1600 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 900 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1000 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        DbClose Cdb, RS
        Charts.SetData Data, Labels, Colors, "Order Destinations By Dollars" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
'******************************
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Labor Percentage" Then
        Charts.DisplayGrid.cols = 3
        ConnectADORS "Select * from [Labor] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        Charts.DisplayGrid.rows = RS.RecordCount + 1
        Charts.DisplayGrid.cols = RS.fields.count
        ReDim FieldName(RS.fields.count - 1)
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        For Cnt = 0 To RS.fields.count - 2
          If OKTOUSE(RS.fields(Cnt + 1).Name) Then
            Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
          End If
        Next
        For Cnt = 0 To RC - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS(1)
            Labels(Cnt) = RS(1) '& " (" & Format(RS(2), "$0.00") & ")"
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(RS(2), "0.00")
            Data(Cnt) = RS(2)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = Format(RS(3), "%0.00")
            RS.MoveNext
        Next
        Charts.DisplayGrid.ColFormat(2) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        Charts.DisplayGrid.ScrollBars = flexScrollBarVertical
        DbClose Cdb, RS
        Charts.SetData Data, Labels, Colors, "Store Labor" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Sales Tax" Then
        ConnectADORS "Select * from [Department Sales] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        Charts.DisplayGrid.rows = RC + 2
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        ReDim FieldName(RS.fields.count - 1)
        Charts.DisplayGrid.cols = RS.fields.count
        TTL = 0
        For Cnt = 0 To RS.fields.count - 2
            Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
        Next
        For Cnt = 0 To RC - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = Check_Null(RS(1))
            Labels(Cnt) = RS(1) & ""
            Data(Cnt) = RS(2) & ""
            TTL = TTL + Val(RS(2) & "")
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(Val(Check_Null(RS(2))), "0.00")
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = Val(Check_Null(RS(3)))
            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(TTL, "0.00")
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = ""
        
        Charts.DisplayGrid.ColFormat(2) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.ColFormat(2) = "(###)"
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        DbClose Cdb, RS
        Charts.SetData Data, Labels, Colors, "Department Sales" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Deposits" Then
        ConnectADORS "Select * from [Deposits] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        Charts.DisplayGrid.rows = RS.RecordCount + 1
        ReDim FieldName(RS.fields.count - 1)
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        For Cnt = 0 To RS.fields.count - 1
            Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt).Name
        Next
        TTL = 0
        For Cnt = 0 To RC - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = "Deposit " & CStr(Cnt + 1)
            Labels(Cnt) = Charts.DisplayGrid.TextMatrix(Cnt + 1, 0)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = RS(1)
            Data(Cnt) = RS(1)
            TTL = TTL + Val(RS(1) & "")
            'Charts.DisplayGrid.Cols = 2
            Charts.DisplayGrid.cols = RS.fields.count
            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(TTL, "0.00")
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = ""
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 900 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        DbClose Cdb, RS
        Charts.SetData Data, Labels, Colors, "Deposits" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Net Sales" Then
        If Not IsDate(SearchDate) Then Exit Sub
        ConnectADORS "Select * from [Meal Periods] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        Charts.DisplayGrid.rows = RS.RecordCount + 1
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        Charts.DisplayGrid.cols = 4
        Charts.DisplayGrid.cols = RS.fields.count
        ReDim FieldName(RS.fields.count - 1)
        For Cnt = 1 To RS.fields.count - 2
            If OKTOUSE(RS.fields(Cnt + 1).Name) Then
              Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
            End If
        Next
        Charts.DisplayGrid.TextMatrix(0, 2) = "Cust. Count"
        Charts.DisplayGrid.TextMatrix(0, 3) = "Cust. Average"
        Charts.DisplayGrid.rows = RC + 1
        DPR = 0: TTL = 0: CT = 0
        For Cnt = 0 To RC - 1
            If RS(2) > 0 Then
            Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = Check_Null(RS(1))
            Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = RS(2)
            Charts.DisplayGrid.TextMatrix(DPR + 1, 2) = RS(3)
            Charts.DisplayGrid.TextMatrix(DPR + 1, 3) = Format(RS(2) / RS(3), "0.00")
            Labels(DPR) = RS(1) & ""
            Data(DPR) = RS(2) & ""
            TTL = TTL + Val(RS(2) & "")
            CT = CT + Val(RS(3) & "")
            DPR = DPR + 1
            End If
            Charts.DisplayGrid.cols = 4
            Charts.DisplayGrid.cols = RS.fields.count
            RS.MoveNext
        Next
        If DPR = 0 Then
         See_Msg "No Data...", 5, ""
         Exit Sub
        End If
        Charts.DisplayGrid.rows = DPR + 2
        Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = Format(TTL, "0.00")
        Charts.DisplayGrid.TextMatrix(DPR + 1, 2) = CT
        Charts.DisplayGrid.TextMatrix(DPR + 1, 3) = Format(TTL / CT, "0.00")
        
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColFormat(2) = "(###)"
        Charts.DisplayGrid.ColWidth(0) = 900 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(3) = 1300 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        
        DbClose Cdb, RS
        Charts.SetData Data, Labels, Colors, "Meal Periods" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Refunds" Then
        ConnectADORS "Select * from [Refunds] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        ReDim FieldName(RS.fields.count - 2)
        Charts.DisplayGrid.cols = 2
        Charts.DisplayGrid.cols = RS.fields.count
        Charts.DisplayGrid.rows = 25
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        For Cnt = 0 To 1
           If OKTOUSE(RS.fields(Cnt + 1).Name) Then
            Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
           End If
        Next
        Charts.DisplayGrid.rows = RC + 2
        DPR = 0
        For Cnt = 0 To RC - 1
            If RS(2) <> 0 Then
            Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = Check_Null(RS(1))
            Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = Format(RS(2), "0.00")
            Labels(DPR) = RS(1) & ""
            TTL = TTL + Val(RS(2) & "")
            Data(DPR) = RS(2) & ""
            DPR = DPR + 1
            End If
            'Charts.DisplayGrid.Cols = 2
            'Charts.DisplayGrid.Cols = RS.Fields.count
            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = Format(TTL, "0.00")
        Charts.DisplayGrid.TextMatrix(DPR + 1, 2) = ""
        DbClose Cdb, RS
        If DPR = 0 Then
         See_Msg "No Data...", 5, ""
         Exit Sub
        End If

        Charts.DisplayGrid.rows = DPR + 1
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 2000 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 900 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        Charts.SetData Data, Labels, Colors, "Refunds" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Gross Sales" Then
        ConnectADORS "Select * from [Hourly Sales] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        
        ReDim FieldName(RS.fields.count - 1)
        Charts.DisplayGrid.cols = 4
        Charts.DisplayGrid.cols = RS.fields.count
        Charts.DisplayGrid.rows = 25
        ReDim Data(RS.RecordCount - 1)
        ReDim Labels(RS.RecordCount - 1)
        For Cnt = 0 To RS.fields.count - 2
             If OKTOUSE(RS.fields(Cnt + 1).Name) Then
                Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
             End If
        Next
        Charts.DisplayGrid.rows = RC + 3
        DPR = 0: TTL = 0
        If Not RS.BOF Then RS.MoveFirst
        'For Cnt = 0 To RS.RecordCount - 1
        Do While Not RS.EOF
            If RS(2) > 0 Then
            Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = Check_Null(RS(1))
            Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = RS(2)
            TTL = TTL + Val(RS(2) & "")
            Charts.DisplayGrid.TextMatrix(DPR + 1, 2) = Format(RS(3), "###")
            Labels(DPR) = RS(1) & ""
            Data(DPR) = RS(2) & ""
            'TTL = TTL + Val(RS(2) & "")
            DPR = DPR + 1
            End If
            Charts.DisplayGrid.cols = 3
            Charts.DisplayGrid.cols = RS.fields.count
            RS.MoveNext
          Loop
        'Next
        Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = Format(TTL, "0.00")
        Charts.DisplayGrid.TextMatrix(DPR + 1, 2) = ""
        DbClose Cdb, RS
        If DPR = 1 Then Exit Sub
        If DPR = 0 Then
         See_Msg "No Data...", 5, ""
         Exit Sub
        End If
        Charts.DisplayGrid.rows = DPR + 2
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 900 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        Charts.SetData Data, Labels, Colors, "Hourly Sales" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Voids" Then
        SetNothing Cdb
        ConnectADODB SvrPath.TerminalPath & ProductHistory, Cdb
        Sql = "Select [Item Name], sum([Amount]) as TotalAmt from [Voids] where [Amount] > 0 and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SearchDate) & " " & GetStartTime), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(SetEndingDate(SearchDate) & " " & SetEndingTime), "yyyy-mm-dd hh:nn:ss") & "' group by [Item Name]"
        ConnectADORS ReturnDateString(Sql), Cdb, RS
        RC = RecordsCount(RS)
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        ReDim FieldName(RS.fields.count - 1)
        Charts.DisplayGrid.cols = 2
        Charts.DisplayGrid.cols = RS.fields.count
        Charts.DisplayGrid.rows = 100
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        Charts.DisplayGrid.rows = RC + 4
        DPR = 0: TTL = 0
        For Cnt = 0 To RC - 1
            On Error Resume Next
            Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = Check_Null(RS![Item Name])
            Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = Format(RS![TotalAmt], "0.00")
            Labels(DPR) = Check_Null(RS![Item Name])
            Data(DPR) = Val(RS![TotalAmt] & "")
            TTL = TTL + Val(RS![TotalAmt] & "")
            DPR = DPR + 1
            Charts.DisplayGrid.cols = 2
            Charts.DisplayGrid.cols = RS.fields.count
            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(DPR + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(DPR + 1, 1) = Format(TTL, "0.00")
        DbClose Cdb, RS
        If DPR = 1 Then Exit Sub
        If DPR = 0 Then
         See_Msg "No Data...", 5, ""
         Exit Sub
        End If
        'Charts.DisplayGrid.Rows = Dpr + 3
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 2500 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        'Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        Charts.SetData Data, Labels, Colors, "Void Sales" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If


    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Discounts" Then
        ConnectADORS "Select * from [Discounts] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        Charts.DisplayGrid.cols = 3
        Charts.DisplayGrid.cols = RS.fields.count
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        
        Charts.DisplayGrid.rows = RC + 2
        ReDim FieldName(RS.fields.count - 1)
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        For Cnt = 0 To RS.fields.count - 2
          If OKTOUSE(RS.fields(Cnt + 1).Name) Then
            Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
          End If
        Next
        For Cnt = 0 To RC - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS(1) & ""
            Labels(Cnt) = RS(1) & ""
            Data(Cnt) = RS(2) & ""
            TTL = TTL + Val(RS(2) & "")
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = RS(2) & ""
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = Format(RS(3), "###")
            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(TTL, "0.00")
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = ""
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1600 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.ColFormat(2) = "(###)"
        Charts.DisplayGrid.width = Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        
        DbClose Cdb, RS
        
        Charts.SetData Data, Labels, Colors, "Discounts" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
    If SummarySheet.TextMatrix(SummarySheet.Row, 0) = "Non Cash Tenders" Then
        ConnectADORS "Select * from [Tenders] where TRY_CAST([Date] AS datetime) = '" & CDate(SearchDate) & "'", Cdb, RS
        RC = RecordsCount(RS)
        Charts.DisplayGrid.cols = 3
        If RC = 0 Then SetNothing RS: Cdb.Close: Set Cdb = Nothing: Exit Sub
        Charts.DisplayGrid.rows = RC + 2
        ReDim FieldName(RS.fields.count - 1)
        ReDim Data(RC - 1)
        ReDim Labels(RC - 1)
        For Cnt = 0 To Charts.DisplayGrid.cols - 1 'RS.Fields.count - 3
           If OKTOUSE(RS.fields(Cnt + 1).Name) Then
            Charts.DisplayGrid.TextMatrix(0, Cnt) = RS.fields(Cnt + 1).Name
           End If
        Next
        For Cnt = 0 To RC - 1
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = RS(1)
            Labels(Cnt) = RS(1)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = RS(2)
            TTL = TTL + Val(RS(2) & "")
            Data(Cnt) = RS(2)
            Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = RS(3)
            

            RS.MoveNext
        Next
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 0) = "Total"
        Charts.DisplayGrid.TextMatrix(Cnt + 1, 1) = Format(TTL, "0.00")
        'Charts.DisplayGrid.TextMatrix(Cnt + 1, 2) = ""
        Charts.DisplayGrid.ColFormat(1) = "($#.00)"
        Charts.DisplayGrid.ColWidth(0) = 1600 * WidScale1024
        Charts.DisplayGrid.ColWidth(1) = 1300 * WidScale1024
        Charts.DisplayGrid.ColWidth(2) = 1300 * WidScale1024
        Charts.DisplayGrid.width = (Charts.DisplayGrid.ColWidth(0) + Charts.DisplayGrid.ColWidth(1)) + Charts.DisplayGrid.ColWidth(2) + 300
        Charts.DisplayGrid.Height = (Charts.DisplayGrid.RowHeight(1) * Charts.DisplayGrid.rows) + 200
        
        DbClose Cdb, RS
        
        Charts.SetData Data, Labels, Colors, "Non Cash Tenders" & " for " & SearchDate
        Charts.Show 1
        Exit Sub
    End If
End If
If Button = 2 Then
SearchSpan = 0
If SummarySheet.MouseCol = 8 Then Exit Sub
With Compare
  .Clear
  .TextMatrix(0, 0) = ""
  .TextMatrix(0, 1) = "Current"
  .TextMatrix(0, 2) = "Past"
  .TextMatrix(0, 3) = "+/- $"
  .TextMatrix(0, 4) = "+/- %"
  .TextMatrix(1, 0) = "Gross Sales"
  .TextMatrix(2, 0) = "Net Sales"
  .TextMatrix(3, 0) = "Sales Tax"
  .TextMatrix(4, 0) = "Special Tax"
  .TextMatrix(5, 0) = SurChargeTxt
  .TextMatrix(6, 0) = "Voids"
  .TextMatrix(7, 0) = "Discounts"
  .TextMatrix(8, 0) = "Non Cash Payments"
  .TextMatrix(9, 0) = "Paid Outs"
  .TextMatrix(10, 0) = "Deposits"
  .TextMatrix(11, 0) = "Non Taxed Sales"
  .TextMatrix(13, 0) = "Ttl Accounted for"
  .TextMatrix(12, 0) = "Ttl To Acct for"
  .TextMatrix(14, 0) = "Over / Short"
  .TextMatrix(15, 0) = "Labor Dollars"
  .TextMatrix(16, 0) = "Labor Hours"
  .TextMatrix(17, 0) = "Labor Percentage"
  .TextMatrix(18, 0) = "Sales Percentage"
  .TextMatrix(19, 0) = "Go Sales"
  .TextMatrix(20, 0) = "Call In Sales"
  .TextMatrix(21, 0) = "Drive Thru Sales"
  '.TextMatrix(22, 0) = "Gift Cert/Paid In"
End With
CompareByTotal.Text = "Day"
SearchHow.Text = "Day"
Dim LR As Integer
   With Compare
    For LR = 0 To 21
         If Val(SummarySheet.TextMatrix(1, SummarySheet.MouseCol)) = 0 Then Exit Sub
         If SummarySheet.TextMatrix(LR, SummarySheet.MouseCol) = "" Then Exit For
         .TextMatrix(LR, 1) = SummarySheet.TextMatrix(LR, SummarySheet.MouseCol)
    Next LR
   End With
With CompareHistory
  CenterObject Me, CompareHistory
   '.Left = 1320 * WidScale1024
   '.Top = 120 * TopScale800x600
   .ZOrder 0
   .Visible = True
End With
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Summary Mouse Down ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Function OKTOUSE(DataIn As String) As Boolean
               OKTOUSE = False
               Select Case UCase(DataIn & "")
                   Case Is = "GEN_SEND XML", "GEN_RETURN XML", "S_GUID", "GEN_SIGN"
                   Case Is = "CONFLICTROWGUID", "S_COLLINEAGE", "S_GENERATION", "S_LINEAGE"
                   Case Else
                      OKTOUSE = True
               End Select
End Function

Private Sub T10_Click()
End Sub

Private Sub TBFR_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Friday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Friday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Friday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Friday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub
Private Sub TBM_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Monday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Monday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Monday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Monday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub
Private Sub TBSAT_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Saturday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Saturday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Saturday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Saturday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub
Private Sub TBSUN_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Sunday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Sunday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Sunday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Sunday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub
Private Sub TBT_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Tuesday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Tuesday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Tuesday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Tuesday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub

Private Sub TBT1_Click(Index As Integer)

End Sub

Private Sub TBTH_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Thursday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Thursday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Thursday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Thursday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub
Private Sub TBW_Click(Index As Integer)
Dim TxStr As String
 Select Case Index
    Case Is = 0
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Wednesday'"
     Sql = Sql & " Order By [Net Sales] desc;"
     TxStr = "Wednesday Top 10"
    Case 1
     Sql = "Select * from [Sales]"
     Sql = Sql & " Where not isnull([date]) and format(TRY_CAST([Date] AS datetime),'dddd') = 'Wednesday'"
     Sql = Sql & " Order By [Net Sales] asc;"
     TxStr = "Wednesday Bottom 10"
 End Select
 GetRecords TxStr, Sql
End Sub
Private Sub Tdrs_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        TenderHistory
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub ThurRS_Click()
End Sub

Private Sub TRS_Click()

End Sub

Private Sub TW_Click()

End Sub

Private Sub VRP_Click()
End Sub

Private Sub VRPT_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        VoidHistoryItemReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub WedRS_Click()

End Sub

Private Sub VS_Click()
  ShowGraph 6, "Voided Sales" & " " & LocName.Caption
End Sub

Private Sub WDS_Click()
  ShowGraph 7, "Discounts" & " " & LocName.Caption
End Sub

Private Sub WKSales_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
        Sales_History "Weekly", SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
        Set SearchDates = Nothing
    End If
End If

End Sub

Private Sub WRS_Click()

End Sub

Private Sub WLH_Click()
  ShowGraph 16, "Labor Hours Used" & " " & LocName.Caption
End Sub
Private Sub ShowGraph(Cr As Integer, GraphTitle As String)
  Dim Data()
  Dim Labels()
  Dim Colors()
  Colors = Array(&HB8BC9C, &HECF0B9, &H999966, &H333366, &HC3C3E6, &H594330, &HA0BDC4)
  With SummarySheet
    Data = Array(Val(.TextMatrix(Cr, 1)), Val(.TextMatrix(Cr, 2)), Val(.TextMatrix(Cr, 3)), Val(.TextMatrix(Cr, 4)), Val(.TextMatrix(Cr, 5)), Val(.TextMatrix(Cr, 6)), Val(.TextMatrix(Cr, 7)))
    Labels = Array(Format(.TextMatrix(0, 1), "dddd"), Format(.TextMatrix(0, 2), "dddd"), Format(.TextMatrix(0, 3), "dddd"), Format(.TextMatrix(0, 4), "dddd"), Format(.TextMatrix(0, 5), "dddd"), Format(.TextMatrix(0, 6), "dddd"), Format(.TextMatrix(0, 7), "dddd"))
  End With
    Charts.DisplayGrid.Visible = False
    Charts.SetData Data, Labels, Colors, GraphTitle
    Charts.Show 1
End Sub
Private Sub ShowBARGraph(Cr As Integer, GraphTitle As String)
  Dim Data()
  Dim Labels()
  Dim Colors()
 ' colors = Array(&HB8BC9C, &HECF0B9, &H999966, &H333366, &HC3C3E6, &H594330, &HA0BDC4)
  Colors = Array(&HB8BC9C, &HA0BDC4, &H999966, &H333366, &HC3C3E6)
  With SummarySheet
    Data = Array(Val(.TextMatrix(Cr, 1)), Val(.TextMatrix(Cr, 2)), Val(.TextMatrix(Cr, 3)), Val(.TextMatrix(Cr, 4)), Val(.TextMatrix(Cr, 5)), Val(.TextMatrix(Cr, 6)), Val(.TextMatrix(Cr, 7)))
    Labels = Array(Format(.TextMatrix(0, 1), "dddd"), Format(.TextMatrix(0, 2), "dddd"), Format(.TextMatrix(0, 3), "dddd"), Format(.TextMatrix(0, 4), "dddd"), Format(.TextMatrix(0, 5), "dddd"), Format(.TextMatrix(0, 6), "dddd"), Format(.TextMatrix(0, 7), "dddd"))
  End With
    Charts.DisplayGrid.Visible = False
    Charts.SetBarData Data, Labels, Colors, GraphTitle
    Charts.Show 1
End Sub

Private Sub WLP_Click()
  ShowGraph 17, "Labor Percentage " & LocName.Caption
End Sub

Private Sub WRP_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
     Get_WasteReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    End If
End If
End Sub

Private Sub WSG_Click()
  ShowGraph 2, "Net Sales " & LocName.Caption
End Sub

Private Sub XCPT_Click()
If GetConfigInfo("Misc", "Compeat Export Is", "Off", "Functions.INI") = "ENABLED" Then
With Review
  CenterObject Me, Review
  .ZOrder 0
  .Visible = True
  .Value = Format(Date, "mm/dd/yyyy")
End With
Do While Review.Visible = True
  DoEvents
  Sleep 100
Loop
If Not IsDate(SearchDate) Then Exit Sub
  See_Msg "Export to Compeat " & Closing_Report.SendCompeat(SearchDate), 5, ""
  
Else
  See_Msg "Not Set Up", 5, ""
End If
End Sub
