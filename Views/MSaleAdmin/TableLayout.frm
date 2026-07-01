VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Begin VB.Form TableLayout 
   Appearance      =   0  'Flat
   BackColor       =   &H00100F0E&
   BorderStyle     =   0  'None
   Caption         =   "Table Layout"
   ClientHeight    =   11535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15375
   FillColor       =   &H00100F0E&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00100F0E&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11535
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ShowChecks 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   -9240
      ScaleHeight     =   6315
      ScaleWidth      =   11835
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   11895
      Begin VSFlex8DAOCtl.VSFlexGrid TData 
         Height          =   4935
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   11415
         _cx             =   1990348455
         _cy             =   1990337025
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   100
         Cols            =   5
         FixedRows       =   0
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
      Begin BTNENHLib4.BtnEnh SSCommand11 
         Height          =   555
         Left            =   10320
         TabIndex        =   52
         Top             =   5520
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   979
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":0000
         textLT          =   "TableLayout.frx":0068
         textCT          =   "TableLayout.frx":0080
         textRT          =   "TableLayout.frx":0098
         textLM          =   "TableLayout.frx":00B0
         textRM          =   "TableLayout.frx":00C8
         textLB          =   "TableLayout.frx":00E0
         textCB          =   "TableLayout.frx":00F8
         textRB          =   "TableLayout.frx":0110
         colorBack       =   "TableLayout.frx":0128
         colorIntern     =   "TableLayout.frx":0152
         colorMO         =   "TableLayout.frx":017C
         colorFocus      =   "TableLayout.frx":01A6
         colorDisabled   =   "TableLayout.frx":01D0
         colorPressed    =   "TableLayout.frx":01FA
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
      Begin VB.Label CheckStatus 
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
         Height          =   375
         Left            =   2880
         TabIndex        =   70
         Top             =   5880
         Width           =   3975
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   120
         Top             =   120
         Width           =   11655
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Table Status"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   120
         Width           =   8415
      End
      Begin VB.Label PCHK 
         BackStyle       =   0  'Transparent
         Caption         =   "Presented Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label SCHK 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Seated Tables"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label FCHK 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Tables With Food"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label OCHK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Tables Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   5520
         Width           =   1695
      End
   End
   Begin VB.Frame TBLUTIL 
      BackColor       =   &H00404040&
      Caption         =   "Import/Export"
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
      Height          =   1575
      Left            =   1440
      TabIndex        =   65
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Export To Access"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":0224
         textLT          =   "TableLayout.frx":02A4
         textCT          =   "TableLayout.frx":02BC
         textRT          =   "TableLayout.frx":02D4
         textLM          =   "TableLayout.frx":02EC
         textRM          =   "TableLayout.frx":0304
         textLB          =   "TableLayout.frx":031C
         textCB          =   "TableLayout.frx":0334
         textRB          =   "TableLayout.frx":034C
         colorBack       =   "TableLayout.frx":0364
         colorIntern     =   "TableLayout.frx":038E
         colorMO         =   "TableLayout.frx":03B8
         colorFocus      =   "TableLayout.frx":03E2
         colorDisabled   =   "TableLayout.frx":040C
         colorPressed    =   "TableLayout.frx":0436
         Style           =   2
      End
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Import From Access"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":0460
         textLT          =   "TableLayout.frx":04E4
         textCT          =   "TableLayout.frx":04FC
         textRT          =   "TableLayout.frx":0514
         textLM          =   "TableLayout.frx":052C
         textRM          =   "TableLayout.frx":0544
         textLB          =   "TableLayout.frx":055C
         textCB          =   "TableLayout.frx":0574
         textRB          =   "TableLayout.frx":058C
         colorBack       =   "TableLayout.frx":05A4
         colorIntern     =   "TableLayout.frx":05CE
         colorMO         =   "TableLayout.frx":05F8
         colorFocus      =   "TableLayout.frx":0622
         colorDisabled   =   "TableLayout.frx":064C
         colorPressed    =   "TableLayout.frx":0676
         Style           =   2
      End
      Begin BTNENHLib4.BtnEnh BtnEnh10 
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   1080
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "Exit"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":06A0
         textLT          =   "TableLayout.frx":0708
         textCT          =   "TableLayout.frx":0720
         textRT          =   "TableLayout.frx":0738
         textLM          =   "TableLayout.frx":0750
         textRM          =   "TableLayout.frx":0768
         textLB          =   "TableLayout.frx":0780
         textCB          =   "TableLayout.frx":0798
         textRB          =   "TableLayout.frx":07B0
         colorBack       =   "TableLayout.frx":07C8
         colorIntern     =   "TableLayout.frx":07F2
         colorMO         =   "TableLayout.frx":081C
         colorFocus      =   "TableLayout.frx":0846
         colorDisabled   =   "TableLayout.frx":0870
         colorPressed    =   "TableLayout.frx":089A
         Style           =   2
      End
   End
   Begin VB.PictureBox HostessMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF9F8&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   12720
      ScaleHeight     =   7215
      ScaleWidth      =   2610
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2609
      Begin VB.ListBox TblList 
         Appearance      =   0  'Flat
         Height          =   2190
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin BTNENHLib4.BtnEnh SSCommand13 
         Height          =   675
         Left            =   120
         TabIndex        =   53
         Top             =   4320
         Width           =   2385
         _Version        =   589828
         _ExtentX        =   4209
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Select Server"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         CornerFactor    =   10
         Surface         =   14
         BackColorContainer=   16447992
         ButtonRaiseFactor=   0
         SmoothEdges     =   1
         CaptionWordWrapPerc=   93
         LogPixels       =   96
         SpecialEffectFactor=   2
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":08C4
         textLT          =   "TableLayout.frx":093E
         textCT          =   "TableLayout.frx":0956
         textRT          =   "TableLayout.frx":096E
         textLM          =   "TableLayout.frx":0986
         textRM          =   "TableLayout.frx":099E
         textLB          =   "TableLayout.frx":09B6
         textCB          =   "TableLayout.frx":09CE
         textRB          =   "TableLayout.frx":09E6
         colorBack       =   "TableLayout.frx":09FE
         colorIntern     =   "TableLayout.frx":0A28
         colorMO         =   "TableLayout.frx":0A52
         colorFocus      =   "TableLayout.frx":0A7C
         colorDisabled   =   "TableLayout.frx":0AA6
         colorPressed    =   "TableLayout.frx":0AD0
         TextCaptionAlignment=   0
         HollowFrame     =   -1  'True
         VistaColorGlossyEffectUpper=   4272160
         VistaColorGlossyEffectLower=   12953225
         VistaOpacityGlossyEffectLower=   75
         VistaColorOuterBorder=   4272160
         VistaColorInnerBorder=   4272160
         VistaColorInnerBorderPressed=   4272160
         VistaColorMiddleBorder=   4272160
         VistaOpacityGlossyEmphLower=   0
         VistaColorGlossyEmphUpperPressed=   2563610
         VistaColorGlossyEmphLowerPressed=   2563610
         VistaGlossyShape=   1
         VistaOuterBorderWidth=   1
         VistaMiddleBorderWidth=   1
         VistaInnerBorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand12 
         Height          =   675
         Left            =   120
         TabIndex        =   54
         Top             =   5040
         Width           =   2386
         _Version        =   589828
         _ExtentX        =   4209
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Assign Section"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         CornerFactor    =   10
         Surface         =   14
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         SmoothEdges     =   1
         CaptionWordWrapPerc=   93
         LogPixels       =   96
         SpecialEffectFactor=   2
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":0AFA
         textLT          =   "TableLayout.frx":0B76
         textCT          =   "TableLayout.frx":0B8E
         textRT          =   "TableLayout.frx":0BA6
         textLM          =   "TableLayout.frx":0BBE
         textRM          =   "TableLayout.frx":0BD6
         textLB          =   "TableLayout.frx":0BEE
         textCB          =   "TableLayout.frx":0C06
         textRB          =   "TableLayout.frx":0C1E
         colorBack       =   "TableLayout.frx":0C36
         colorIntern     =   "TableLayout.frx":0C60
         colorMO         =   "TableLayout.frx":0C8A
         colorFocus      =   "TableLayout.frx":0CB4
         colorDisabled   =   "TableLayout.frx":0CDE
         colorPressed    =   "TableLayout.frx":0D08
         TextCaptionAlignment=   0
         HollowFrame     =   -1  'True
         VistaColorGlossyEffectUpper=   4272160
         VistaColorGlossyEffectLower=   12953225
         VistaOpacityGlossyEffectLower=   75
         VistaColorOuterBorder=   4272160
         VistaColorInnerBorder=   4272160
         VistaColorInnerBorderPressed=   4272160
         VistaColorMiddleBorder=   4272160
         VistaOpacityGlossyEmphLower=   0
         VistaColorGlossyEmphUpperPressed=   2563610
         VistaColorGlossyEmphLowerPressed=   2563610
         VistaGlossyShape=   1
         VistaOuterBorderWidth=   1
         VistaMiddleBorderWidth=   1
         VistaInnerBorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand4 
         Height          =   555
         Left            =   360
         TabIndex        =   55
         Top             =   3360
         Width           =   735
         _Version        =   589828
         _ExtentX        =   1296
         _ExtentY        =   979
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Shape           =   8
         CornerFactor    =   10
         Surface         =   2
         ButtonRaiseFactor=   0
         SmoothEdges     =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":0D32
         textLT          =   "TableLayout.frx":0D4A
         textCT          =   "TableLayout.frx":0D62
         textRT          =   "TableLayout.frx":0D7A
         textLM          =   "TableLayout.frx":0D92
         textRM          =   "TableLayout.frx":0DAA
         textLB          =   "TableLayout.frx":0DC2
         textCB          =   "TableLayout.frx":0DDA
         textRB          =   "TableLayout.frx":0DF2
         colorBack       =   "TableLayout.frx":0E0A
         colorIntern     =   "TableLayout.frx":0E34
         colorMO         =   "TableLayout.frx":0E5E
         colorFocus      =   "TableLayout.frx":0E88
         colorDisabled   =   "TableLayout.frx":0EB2
         colorPressed    =   "TableLayout.frx":0EDC
         Orientation     =   1
         TextCaptionAlignment=   0
         HollowFrame     =   -1  'True
         LightDirection  =   8
         ShapeHeadFactor =   43
         ShapeLineFactor =   55
         SurfaceTransparentFactor=   237
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
         Office2003ColorNorth=   7364430
         Office2003ColorSouth=   6443844
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   7364430
         Office2003ColorSouthMouseOver=   6443844
         Office2003ColorBorderMouseOver=   12632256
         Office2003ColorNorthPressed=   10258290
         Office2003ColorSouthPressed=   10258290
         Office2003ColorBorderPressed=   16739100
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand16 
         Height          =   555
         Left            =   1560
         TabIndex        =   56
         Top             =   3360
         Width           =   735
         _Version        =   589828
         _ExtentX        =   1296
         _ExtentY        =   979
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Shape           =   8
         CornerFactor    =   10
         Surface         =   2
         ButtonRaiseFactor=   0
         SmoothEdges     =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":0F06
         textLT          =   "TableLayout.frx":0F1E
         textCT          =   "TableLayout.frx":0F36
         textRT          =   "TableLayout.frx":0F4E
         textLM          =   "TableLayout.frx":0F66
         textRM          =   "TableLayout.frx":0F7E
         textLB          =   "TableLayout.frx":0F96
         textCB          =   "TableLayout.frx":0FAE
         textRB          =   "TableLayout.frx":0FC6
         colorBack       =   "TableLayout.frx":0FDE
         colorIntern     =   "TableLayout.frx":1008
         colorMO         =   "TableLayout.frx":1032
         colorFocus      =   "TableLayout.frx":105C
         colorDisabled   =   "TableLayout.frx":1086
         colorPressed    =   "TableLayout.frx":10B0
         Orientation     =   3
         TextCaptionAlignment=   0
         HollowFrame     =   -1  'True
         LightDirection  =   8
         ShapeHeadFactor =   43
         ShapeLineFactor =   55
         SurfaceTransparentFactor=   237
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
         Office2003ColorNorth=   7364430
         Office2003ColorSouth=   6443844
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   7364430
         Office2003ColorSouthMouseOver=   6443844
         Office2003ColorBorderMouseOver=   12632256
         Office2003ColorNorthPressed=   10258290
         Office2003ColorSouthPressed=   10258290
         Office2003ColorBorderPressed=   16739100
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand14 
         Height          =   645
         Left            =   120
         TabIndex        =   57
         Top             =   6480
         Width           =   2385
         _Version        =   589828
         _ExtentX        =   4207
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Hide"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         CornerFactor    =   10
         Surface         =   15
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   85
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":10DA
         textLT          =   "TableLayout.frx":1142
         textCT          =   "TableLayout.frx":115A
         textRT          =   "TableLayout.frx":1172
         textLM          =   "TableLayout.frx":118A
         textRM          =   "TableLayout.frx":11A2
         textLB          =   "TableLayout.frx":11BA
         textCB          =   "TableLayout.frx":11D2
         textRB          =   "TableLayout.frx":11EA
         colorBack       =   "TableLayout.frx":1202
         colorIntern     =   "TableLayout.frx":122C
         colorMO         =   "TableLayout.frx":1256
         colorFocus      =   "TableLayout.frx":1280
         colorDisabled   =   "TableLayout.frx":12AA
         colorPressed    =   "TableLayout.frx":12D4
         TextCaptionAlignment=   0
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectUpper=   8454143
         VistaColorGlossyEffectLower=   33023
         VistaColorOuterBorder=   16512
         VistaColorInnerBorder=   65535
         VistaColorMiddleBorder=   33023
         VistaColorGlowBack=   16512
         VistaColorGlow  =   65535
         VistaLook       =   1
         VistaGlossyShape=   3
         VistaGlossySmoothFactor=   1
         SpotlightType   =   1
         Office2003ColorNorth=   3290088
         Office2003ColorSouth=   3290088
         Office2003ColorBorder=   3290088
         Office2003ColorNorthMouseOver=   3290088
         Office2003ColorSouthMouseOver=   3290088
         Office2003ColorBorderMouseOver=   1181949
         Office2003ColorNorthPressed=   2167785
         Office2003ColorSouthPressed=   853993
         Office2003ColorBorderPressed=   1181949
         Office2003BorderWidth=   3
      End
      Begin VB.Label SvrsName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   112
         TabIndex        =   13
         Top             =   120
         Width           =   2386
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tables Assigned"
         ForeColor       =   &H00100F0E&
         Height          =   3615
         Left            =   105
         TabIndex        =   14
         Top             =   480
         Width           =   2385
      End
   End
   Begin VB.PictureBox ServerMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00100F0E&
      BorderStyle     =   0  'None
      FillColor       =   &H00100F0E&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   15135
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   15135
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   855
         Left            =   1470
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Split Check"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":12FE
         textLT          =   "TableLayout.frx":1374
         textCT          =   "TableLayout.frx":138C
         textRT          =   "TableLayout.frx":13A4
         textLM          =   "TableLayout.frx":13BC
         textRM          =   "TableLayout.frx":13D4
         textLB          =   "TableLayout.frx":13EC
         textCB          =   "TableLayout.frx":1404
         textRB          =   "TableLayout.frx":141C
         colorBack       =   "TableLayout.frx":1434
         colorIntern     =   "TableLayout.frx":145E
         colorMO         =   "TableLayout.frx":1488
         colorFocus      =   "TableLayout.frx":14B2
         colorDisabled   =   "TableLayout.frx":14DC
         colorPressed    =   "TableLayout.frx":1506
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   8404992
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand22 
         Height          =   855
         Left            =   2820
         TabIndex        =   34
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":1530
         textLT          =   "TableLayout.frx":159A
         textCT          =   "TableLayout.frx":15B2
         textRT          =   "TableLayout.frx":15CA
         textLM          =   "TableLayout.frx":15E2
         textRM          =   "TableLayout.frx":15FA
         textLB          =   "TableLayout.frx":1612
         textCB          =   "TableLayout.frx":162A
         textRB          =   "TableLayout.frx":1642
         colorBack       =   "TableLayout.frx":165A
         colorIntern     =   "TableLayout.frx":1684
         colorMO         =   "TableLayout.frx":16AE
         colorFocus      =   "TableLayout.frx":16D8
         colorDisabled   =   "TableLayout.frx":1702
         colorPressed    =   "TableLayout.frx":172C
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   12870144
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh OpenTable 
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Add Check To Table"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":1756
         textLT          =   "TableLayout.frx":17DA
         textCT          =   "TableLayout.frx":17F2
         textRT          =   "TableLayout.frx":180A
         textLM          =   "TableLayout.frx":1822
         textRM          =   "TableLayout.frx":183A
         textLB          =   "TableLayout.frx":1852
         textCB          =   "TableLayout.frx":186A
         textRB          =   "TableLayout.frx":1882
         colorBack       =   "TableLayout.frx":189A
         colorIntern     =   "TableLayout.frx":18C4
         colorMO         =   "TableLayout.frx":18EE
         colorFocus      =   "TableLayout.frx":1918
         colorDisabled   =   "TableLayout.frx":1942
         colorPressed    =   "TableLayout.frx":196C
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   12870144
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh Status 
         Height          =   855
         Left            =   1475
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Table Status"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":1996
         textLT          =   "TableLayout.frx":1A0E
         textCT          =   "TableLayout.frx":1A26
         textRT          =   "TableLayout.frx":1A3E
         textLM          =   "TableLayout.frx":1A56
         textRM          =   "TableLayout.frx":1A6E
         textLB          =   "TableLayout.frx":1A86
         textCB          =   "TableLayout.frx":1A9E
         textRB          =   "TableLayout.frx":1AB6
         colorBack       =   "TableLayout.frx":1ACE
         colorIntern     =   "TableLayout.frx":1AF8
         colorMO         =   "TableLayout.frx":1B22
         colorFocus      =   "TableLayout.frx":1B4C
         colorDisabled   =   "TableLayout.frx":1B76
         colorPressed    =   "TableLayout.frx":1BA0
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
      Begin BTNENHLib4.BtnEnh SSCommand7 
         Height          =   855
         Left            =   5520
         TabIndex        =   35
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Apply Tip"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":1BCA
         textLT          =   "TableLayout.frx":1C3C
         textCT          =   "TableLayout.frx":1C54
         textRT          =   "TableLayout.frx":1C6C
         textLM          =   "TableLayout.frx":1C84
         textRM          =   "TableLayout.frx":1C9C
         textLB          =   "TableLayout.frx":1CB4
         textCB          =   "TableLayout.frx":1CCC
         textRB          =   "TableLayout.frx":1CE4
         colorBack       =   "TableLayout.frx":1CFC
         colorIntern     =   "TableLayout.frx":1D26
         colorMO         =   "TableLayout.frx":1D50
         colorFocus      =   "TableLayout.frx":1D7A
         colorDisabled   =   "TableLayout.frx":1DA4
         colorPressed    =   "TableLayout.frx":1DCE
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   8404992
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand15 
         Height          =   855
         Left            =   4170
         TabIndex        =   36
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Waiting List"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":1DF8
         textLT          =   "TableLayout.frx":1E70
         textCT          =   "TableLayout.frx":1E88
         textRT          =   "TableLayout.frx":1EA0
         textLM          =   "TableLayout.frx":1EB8
         textRM          =   "TableLayout.frx":1ED0
         textLB          =   "TableLayout.frx":1EE8
         textCB          =   "TableLayout.frx":1F00
         textRB          =   "TableLayout.frx":1F18
         colorBack       =   "TableLayout.frx":1F30
         colorIntern     =   "TableLayout.frx":1F5A
         colorMO         =   "TableLayout.frx":1F84
         colorFocus      =   "TableLayout.frx":1FAE
         colorDisabled   =   "TableLayout.frx":1FD8
         colorPressed    =   "TableLayout.frx":2002
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   8404992
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand17 
         Height          =   855
         Left            =   6870
         TabIndex        =   37
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Join Table"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":202C
         textLT          =   "TableLayout.frx":20A0
         textCT          =   "TableLayout.frx":20B8
         textRT          =   "TableLayout.frx":20D0
         textLM          =   "TableLayout.frx":20E8
         textRM          =   "TableLayout.frx":2100
         textLB          =   "TableLayout.frx":2118
         textCB          =   "TableLayout.frx":2130
         textRB          =   "TableLayout.frx":2148
         colorBack       =   "TableLayout.frx":2160
         colorIntern     =   "TableLayout.frx":218A
         colorMO         =   "TableLayout.frx":21B4
         colorFocus      =   "TableLayout.frx":21DE
         colorDisabled   =   "TableLayout.frx":2208
         colorPressed    =   "TableLayout.frx":2232
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   8404992
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand8 
         Height          =   855
         Left            =   13620
         TabIndex        =   38
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Cancel"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":225C
         textLT          =   "TableLayout.frx":22C8
         textCT          =   "TableLayout.frx":22E0
         textRT          =   "TableLayout.frx":22F8
         textLM          =   "TableLayout.frx":2310
         textRM          =   "TableLayout.frx":2328
         textLB          =   "TableLayout.frx":2340
         textCB          =   "TableLayout.frx":2358
         textRB          =   "TableLayout.frx":2370
         colorBack       =   "TableLayout.frx":2388
         colorIntern     =   "TableLayout.frx":23B2
         colorMO         =   "TableLayout.frx":23DC
         colorFocus      =   "TableLayout.frx":2406
         colorDisabled   =   "TableLayout.frx":2430
         colorPressed    =   "TableLayout.frx":245A
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
      Begin BTNENHLib4.BtnEnh SSCommand10 
         Height          =   855
         Left            =   8220
         TabIndex        =   39
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Close"
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
         BackColorContainer=   4194304
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":2484
         textLT          =   "TableLayout.frx":24EE
         textCT          =   "TableLayout.frx":2506
         textRT          =   "TableLayout.frx":251E
         textLM          =   "TableLayout.frx":2536
         textRM          =   "TableLayout.frx":254E
         textLB          =   "TableLayout.frx":2566
         textCB          =   "TableLayout.frx":257E
         textRB          =   "TableLayout.frx":2596
         colorBack       =   "TableLayout.frx":25AE
         colorIntern     =   "TableLayout.frx":25D8
         colorMO         =   "TableLayout.frx":2602
         colorFocus      =   "TableLayout.frx":262C
         colorDisabled   =   "TableLayout.frx":2656
         colorPressed    =   "TableLayout.frx":2680
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
         Office2003ColorNorth=   198
         Office2003ColorSouth=   128
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   198
         Office2003ColorSouthMouseOver=   198
         Office2003ColorBorderMouseOver=   1812708
         Office2003ColorNorthPressed=   198
         Office2003ColorSouthPressed=   128
         Office2003ColorBorderPressed=   1812708
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand3 
         Height          =   855
         Left            =   10920
         TabIndex        =   40
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Server Menu"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":26AA
         textLT          =   "TableLayout.frx":2720
         textCT          =   "TableLayout.frx":2738
         textRT          =   "TableLayout.frx":2750
         textLM          =   "TableLayout.frx":2768
         textRM          =   "TableLayout.frx":2780
         textLB          =   "TableLayout.frx":2798
         textCB          =   "TableLayout.frx":27B0
         textRB          =   "TableLayout.frx":27C8
         colorBack       =   "TableLayout.frx":27E0
         colorIntern     =   "TableLayout.frx":280A
         colorMO         =   "TableLayout.frx":2834
         colorFocus      =   "TableLayout.frx":285E
         colorDisabled   =   "TableLayout.frx":2888
         colorPressed    =   "TableLayout.frx":28B2
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand1 
         Height          =   855
         Left            =   9570
         TabIndex        =   41
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Change Section"
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
         BackColorContainer=   4194304
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":28DC
         textLT          =   "TableLayout.frx":2958
         textCT          =   "TableLayout.frx":2970
         textRT          =   "TableLayout.frx":2988
         textLM          =   "TableLayout.frx":29A0
         textRM          =   "TableLayout.frx":29B8
         textLB          =   "TableLayout.frx":29D0
         textCB          =   "TableLayout.frx":29E8
         textRB          =   "TableLayout.frx":2A00
         colorBack       =   "TableLayout.frx":2A18
         colorIntern     =   "TableLayout.frx":2A42
         colorMO         =   "TableLayout.frx":2A6C
         colorFocus      =   "TableLayout.frx":2A96
         colorDisabled   =   "TableLayout.frx":2AC0
         colorPressed    =   "TableLayout.frx":2AEA
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh ShowAssign 
         Height          =   855
         Left            =   2820
         TabIndex        =   42
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Show Assigned"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":2B14
         textLT          =   "TableLayout.frx":2B8E
         textCT          =   "TableLayout.frx":2BA6
         textRT          =   "TableLayout.frx":2BBE
         textLM          =   "TableLayout.frx":2BD6
         textRM          =   "TableLayout.frx":2BEE
         textLB          =   "TableLayout.frx":2C06
         textCB          =   "TableLayout.frx":2C1E
         textRB          =   "TableLayout.frx":2C36
         colorBack       =   "TableLayout.frx":2C4E
         colorIntern     =   "TableLayout.frx":2C78
         colorMO         =   "TableLayout.frx":2CA2
         colorFocus      =   "TableLayout.frx":2CCC
         colorDisabled   =   "TableLayout.frx":2CF6
         colorPressed    =   "TableLayout.frx":2D20
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
      Begin BTNENHLib4.BtnEnh BtnEnh12 
         Height          =   855
         Left            =   12270
         TabIndex        =   72
         Top             =   600
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Move Table"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":2D4A
         textLT          =   "TableLayout.frx":2DBE
         textCT          =   "TableLayout.frx":2DD6
         textRT          =   "TableLayout.frx":2DEE
         textLM          =   "TableLayout.frx":2E06
         textRM          =   "TableLayout.frx":2E1E
         textLB          =   "TableLayout.frx":2E36
         textCB          =   "TableLayout.frx":2E4E
         textRB          =   "TableLayout.frx":2E66
         colorBack       =   "TableLayout.frx":2E7E
         colorIntern     =   "TableLayout.frx":2EA8
         colorMO         =   "TableLayout.frx":2ED2
         colorFocus      =   "TableLayout.frx":2EFC
         colorDisabled   =   "TableLayout.frx":2F26
         colorPressed    =   "TableLayout.frx":2F50
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
         Office2003ColorBorderMouseOver=   0
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   0
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   11640
         TabIndex        =   73
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label GenInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label LoggedOn 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   7440
         TabIndex        =   9
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Presented"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   8760
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Expired "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Seated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6720
         TabIndex        =   6
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   0
         Width           =   495
      End
      Begin VB.Label ReadyTo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox EditBox 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   9660
      TabIndex        =   2
      Top             =   9960
      Visible         =   0   'False
      Width           =   9720
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   255
         Left            =   6240
         TabIndex        =   64
         Top             =   720
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Align Bottom"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":2F7A
         textLT          =   "TableLayout.frx":2FF2
         textCT          =   "TableLayout.frx":300A
         textRT          =   "TableLayout.frx":3022
         textLM          =   "TableLayout.frx":303A
         textRM          =   "TableLayout.frx":3052
         textLB          =   "TableLayout.frx":306A
         textCB          =   "TableLayout.frx":3082
         textRB          =   "TableLayout.frx":309A
         colorBack       =   "TableLayout.frx":30B2
         colorIntern     =   "TableLayout.frx":30DC
         colorMO         =   "TableLayout.frx":3106
         colorFocus      =   "TableLayout.frx":3130
         colorDisabled   =   "TableLayout.frx":315A
         colorPressed    =   "TableLayout.frx":3184
         Style           =   2
      End
      Begin VB.ComboBox LayoutName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin BTNENHLib4.BtnEnh LoadTables 
         Height          =   270
         Left            =   120
         TabIndex        =   44
         Top             =   15
         Width           =   1215
         _Version        =   589828
         _ExtentX        =   2143
         _ExtentY        =   476
         _StockProps     =   66
         Caption         =   "Tables"
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":31AE
         textLT          =   "TableLayout.frx":321A
         textCT          =   "TableLayout.frx":3232
         textRT          =   "TableLayout.frx":324A
         textLM          =   "TableLayout.frx":3262
         textRM          =   "TableLayout.frx":327A
         textLB          =   "TableLayout.frx":3292
         textCB          =   "TableLayout.frx":32AA
         textRB          =   "TableLayout.frx":32C2
         colorBack       =   "TableLayout.frx":32DA
         colorIntern     =   "TableLayout.frx":3304
         colorMO         =   "TableLayout.frx":332E
         colorFocus      =   "TableLayout.frx":3358
         colorDisabled   =   "TableLayout.frx":3382
         colorPressed    =   "TableLayout.frx":33AC
      End
      Begin BTNENHLib4.BtnEnh LoadFloor 
         Height          =   270
         Left            =   1440
         TabIndex        =   45
         Top             =   15
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   476
         _StockProps     =   66
         Caption         =   "Floor"
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":33D6
         textLT          =   "TableLayout.frx":3440
         textCT          =   "TableLayout.frx":3458
         textRT          =   "TableLayout.frx":3470
         textLM          =   "TableLayout.frx":3488
         textRM          =   "TableLayout.frx":34A0
         textLB          =   "TableLayout.frx":34B8
         textCB          =   "TableLayout.frx":34D0
         textRB          =   "TableLayout.frx":34E8
         colorBack       =   "TableLayout.frx":3500
         colorIntern     =   "TableLayout.frx":352A
         colorMO         =   "TableLayout.frx":3554
         colorFocus      =   "TableLayout.frx":357E
         colorDisabled   =   "TableLayout.frx":35A8
         colorPressed    =   "TableLayout.frx":35D2
      End
      Begin BTNENHLib4.BtnEnh Save 
         Height          =   270
         Left            =   2880
         TabIndex        =   46
         Top             =   15
         Width           =   1215
         _Version        =   589828
         _ExtentX        =   2143
         _ExtentY        =   476
         _StockProps     =   66
         Caption         =   "Save"
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":35FC
         textLT          =   "TableLayout.frx":3664
         textCT          =   "TableLayout.frx":367C
         textRT          =   "TableLayout.frx":3694
         textLM          =   "TableLayout.frx":36AC
         textRM          =   "TableLayout.frx":36C4
         textLB          =   "TableLayout.frx":36DC
         textCB          =   "TableLayout.frx":36F4
         textRB          =   "TableLayout.frx":370C
         colorBack       =   "TableLayout.frx":3724
         colorIntern     =   "TableLayout.frx":374E
         colorMO         =   "TableLayout.frx":3778
         colorFocus      =   "TableLayout.frx":37A2
         colorDisabled   =   "TableLayout.frx":37CC
         colorPressed    =   "TableLayout.frx":37F6
      End
      Begin BTNENHLib4.BtnEnh startover 
         Height          =   315
         Left            =   2880
         TabIndex        =   47
         Top             =   360
         Width           =   1215
         _Version        =   589828
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   66
         Caption         =   "Clear"
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":3820
         textLT          =   "TableLayout.frx":388A
         textCT          =   "TableLayout.frx":38A2
         textRT          =   "TableLayout.frx":38BA
         textLM          =   "TableLayout.frx":38D2
         textRM          =   "TableLayout.frx":38EA
         textLB          =   "TableLayout.frx":3902
         textCB          =   "TableLayout.frx":391A
         textRB          =   "TableLayout.frx":3932
         colorBack       =   "TableLayout.frx":394A
         colorIntern     =   "TableLayout.frx":3974
         colorMO         =   "TableLayout.frx":399E
         colorFocus      =   "TableLayout.frx":39C8
         colorDisabled   =   "TableLayout.frx":39F2
         colorPressed    =   "TableLayout.frx":3A1C
      End
      Begin BTNENHLib4.BtnEnh SSCommand9 
         Height          =   315
         Left            =   1440
         TabIndex        =   48
         Top             =   360
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   66
         Caption         =   "Delete Section"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":3A46
         textLT          =   "TableLayout.frx":3AC2
         textCT          =   "TableLayout.frx":3ADA
         textRT          =   "TableLayout.frx":3AF2
         textLM          =   "TableLayout.frx":3B0A
         textRM          =   "TableLayout.frx":3B22
         textLB          =   "TableLayout.frx":3B3A
         textCB          =   "TableLayout.frx":3B52
         textRB          =   "TableLayout.frx":3B6A
         colorBack       =   "TableLayout.frx":3B82
         colorIntern     =   "TableLayout.frx":3BAC
         colorMO         =   "TableLayout.frx":3BD6
         colorFocus      =   "TableLayout.frx":3C00
         colorDisabled   =   "TableLayout.frx":3C2A
         colorPressed    =   "TableLayout.frx":3C54
      End
      Begin BTNENHLib4.BtnEnh SSCommand2 
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1215
         _Version        =   589828
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   66
         Caption         =   "Delete Table"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":3C7E
         textLT          =   "TableLayout.frx":3CF6
         textCT          =   "TableLayout.frx":3D0E
         textRT          =   "TableLayout.frx":3D26
         textLM          =   "TableLayout.frx":3D3E
         textRM          =   "TableLayout.frx":3D56
         textLB          =   "TableLayout.frx":3D6E
         textCB          =   "TableLayout.frx":3D86
         textRB          =   "TableLayout.frx":3D9E
         colorBack       =   "TableLayout.frx":3DB6
         colorIntern     =   "TableLayout.frx":3DE0
         colorMO         =   "TableLayout.frx":3E0A
         colorFocus      =   "TableLayout.frx":3E34
         colorDisabled   =   "TableLayout.frx":3E5E
         colorPressed    =   "TableLayout.frx":3E88
      End
      Begin BTNENHLib4.BtnEnh SSCommand5 
         Height          =   435
         Left            =   7440
         TabIndex        =   50
         Top             =   120
         Width           =   1935
         _Version        =   589828
         _ExtentX        =   3413
         _ExtentY        =   767
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
         ButtonRaiseFactor=   6
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":3EB2
         textLT          =   "TableLayout.frx":3F1A
         textCT          =   "TableLayout.frx":3F32
         textRT          =   "TableLayout.frx":3F4A
         textLM          =   "TableLayout.frx":3F62
         textRM          =   "TableLayout.frx":3F7A
         textLB          =   "TableLayout.frx":3F92
         textCB          =   "TableLayout.frx":3FAA
         textRB          =   "TableLayout.frx":3FC2
         colorBack       =   "TableLayout.frx":3FDA
         colorIntern     =   "TableLayout.frx":4004
         colorMO         =   "TableLayout.frx":402E
         colorFocus      =   "TableLayout.frx":4058
         colorDisabled   =   "TableLayout.frx":4082
         colorPressed    =   "TableLayout.frx":40AC
      End
      Begin BTNENHLib4.BtnEnh BtnEnh6 
         Height          =   255
         Left            =   4560
         TabIndex        =   60
         Top             =   720
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Align Right"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":40D6
         textLT          =   "TableLayout.frx":414C
         textCT          =   "TableLayout.frx":4164
         textRT          =   "TableLayout.frx":417C
         textLM          =   "TableLayout.frx":4194
         textRM          =   "TableLayout.frx":41AC
         textLB          =   "TableLayout.frx":41C4
         textCB          =   "TableLayout.frx":41DC
         textRB          =   "TableLayout.frx":41F4
         colorBack       =   "TableLayout.frx":420C
         colorIntern     =   "TableLayout.frx":4236
         colorMO         =   "TableLayout.frx":4260
         colorFocus      =   "TableLayout.frx":428A
         colorDisabled   =   "TableLayout.frx":42B4
         colorPressed    =   "TableLayout.frx":42DE
         Style           =   2
      End
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   255
         Left            =   0
         TabIndex        =   61
         Top             =   720
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Snap Right"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":4308
         textLT          =   "TableLayout.frx":437C
         textCT          =   "TableLayout.frx":4394
         textRT          =   "TableLayout.frx":43AC
         textLM          =   "TableLayout.frx":43C4
         textRM          =   "TableLayout.frx":43DC
         textLB          =   "TableLayout.frx":43F4
         textCB          =   "TableLayout.frx":440C
         textRB          =   "TableLayout.frx":4424
         colorBack       =   "TableLayout.frx":443C
         colorIntern     =   "TableLayout.frx":4466
         colorMO         =   "TableLayout.frx":4490
         colorFocus      =   "TableLayout.frx":44BA
         colorDisabled   =   "TableLayout.frx":44E4
         colorPressed    =   "TableLayout.frx":450E
         Style           =   2
      End
      Begin BTNENHLib4.BtnEnh BtnEnh8 
         Height          =   255
         Left            =   1440
         TabIndex        =   62
         Top             =   720
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Snap Bottom"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":4538
         textLT          =   "TableLayout.frx":45AE
         textCT          =   "TableLayout.frx":45C6
         textRT          =   "TableLayout.frx":45DE
         textLM          =   "TableLayout.frx":45F6
         textRM          =   "TableLayout.frx":460E
         textLB          =   "TableLayout.frx":4626
         textCB          =   "TableLayout.frx":463E
         textRB          =   "TableLayout.frx":4656
         colorBack       =   "TableLayout.frx":466E
         colorIntern     =   "TableLayout.frx":4698
         colorMO         =   "TableLayout.frx":46C2
         colorFocus      =   "TableLayout.frx":46EC
         colorDisabled   =   "TableLayout.frx":4716
         colorPressed    =   "TableLayout.frx":4740
         Style           =   2
      End
      Begin BTNENHLib4.BtnEnh BtnEnh9 
         Height          =   255
         Left            =   3000
         TabIndex        =   63
         Top             =   720
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Cancel Options"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":476A
         textLT          =   "TableLayout.frx":47E6
         textCT          =   "TableLayout.frx":47FE
         textRT          =   "TableLayout.frx":4816
         textLM          =   "TableLayout.frx":482E
         textRM          =   "TableLayout.frx":4846
         textLB          =   "TableLayout.frx":485E
         textCB          =   "TableLayout.frx":4876
         textRB          =   "TableLayout.frx":488E
         colorBack       =   "TableLayout.frx":48A6
         colorIntern     =   "TableLayout.frx":48D0
         colorMO         =   "TableLayout.frx":48FA
         colorFocus      =   "TableLayout.frx":4924
         colorDisabled   =   "TableLayout.frx":494E
         colorPressed    =   "TableLayout.frx":4978
         Style           =   2
      End
      Begin BTNENHLib4.BtnEnh BtnEnh5 
         Height          =   255
         Left            =   7920
         TabIndex        =   68
         Top             =   720
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Utility"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Surface         =   4
         SmoothEdges     =   2
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "TableLayout.frx":49A2
         textLT          =   "TableLayout.frx":4A10
         textCT          =   "TableLayout.frx":4A28
         textRT          =   "TableLayout.frx":4A40
         textLM          =   "TableLayout.frx":4A58
         textRM          =   "TableLayout.frx":4A70
         textLB          =   "TableLayout.frx":4A88
         textCB          =   "TableLayout.frx":4AA0
         textRB          =   "TableLayout.frx":4AB8
         colorBack       =   "TableLayout.frx":4AD0
         colorIntern     =   "TableLayout.frx":4AFA
         colorMO         =   "TableLayout.frx":4B24
         colorFocus      =   "TableLayout.frx":4B4E
         colorDisabled   =   "TableLayout.frx":4B78
         colorPressed    =   "TableLayout.frx":4BA2
         Style           =   2
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Section Name"
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
         Height          =   615
         Left            =   4200
         TabIndex        =   4
         Top             =   0
         Width           =   3135
      End
   End
   Begin BTNENHLib4.BtnEnh BtnEnh11 
      Height          =   855
      Left            =   240
      TabIndex        =   71
      Top             =   10080
      Width           =   1335
      _Version        =   589828
      _ExtentX        =   2355
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Show Menu"
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
      BackColorContainer=   8388608
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "TableLayout.frx":4BCC
      textLT          =   "TableLayout.frx":4C3E
      textCT          =   "TableLayout.frx":4C56
      textRT          =   "TableLayout.frx":4C6E
      textLM          =   "TableLayout.frx":4C86
      textRM          =   "TableLayout.frx":4C9E
      textLB          =   "TableLayout.frx":4CB6
      textCB          =   "TableLayout.frx":4CCE
      textRB          =   "TableLayout.frx":4CE6
      colorBack       =   "TableLayout.frx":4CFE
      colorIntern     =   "TableLayout.frx":4D28
      colorMO         =   "TableLayout.frx":4D52
      colorFocus      =   "TableLayout.frx":4D7C
      colorDisabled   =   "TableLayout.frx":4DA6
      colorPressed    =   "TableLayout.frx":4DD0
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
      Office2003ColorNorth=   198
      Office2003ColorSouth=   128
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   198
      Office2003ColorSouthMouseOver=   198
      Office2003ColorBorderMouseOver=   1812708
      Office2003ColorNorthPressed=   198
      Office2003ColorSouthPressed=   128
      Office2003ColorBorderPressed=   1812708
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VB.PictureBox Room 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00302821&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   0
      ScaleHeight     =   9735
      ScaleWidth      =   15315
      TabIndex        =   23
      Top             =   0
      Width           =   15310
      Begin VB.Timer IamWaiting 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   0
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox TableData 
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
         Height          =   6255
         Left            =   5640
         ScaleHeight     =   6225
         ScaleWidth      =   2475
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   2500
         Begin VB.TextBox TableName 
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
            Height          =   285
            Left            =   1200
            TabIndex        =   26
            Top             =   4800
            Width           =   975
         End
         Begin VB.FileListBox bmpfile 
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
            Height          =   1920
            Left            =   240
            Pattern         =   "FLR-*.*"
            TabIndex        =   25
            Top             =   2640
            Width           =   1935
         End
         Begin Threed.SSCheck Cover 
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   5160
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   196610
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Covers Required ?"
         End
         Begin BTNENHLib4.BtnEnh SSCommand6 
            Height          =   435
            Left            =   120
            TabIndex        =   51
            Top             =   5655
            Width           =   2175
            _Version        =   589828
            _ExtentX        =   3836
            _ExtentY        =   767
            _StockProps     =   66
            Caption         =   "Done"
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
            LogPixels       =   96
            TextureMode     =   0
            TextureBevelFactor=   4
            UserData        =   0.1
            textCaption     =   "TableLayout.frx":4DFA
            textLT          =   "TableLayout.frx":4E62
            textCT          =   "TableLayout.frx":4E7A
            textRT          =   "TableLayout.frx":4E92
            textLM          =   "TableLayout.frx":4EAA
            textRM          =   "TableLayout.frx":4EC2
            textLB          =   "TableLayout.frx":4EDA
            textCB          =   "TableLayout.frx":4EF2
            textRB          =   "TableLayout.frx":4F0A
            colorBack       =   "TableLayout.frx":4F22
            colorIntern     =   "TableLayout.frx":4F4C
            colorMO         =   "TableLayout.frx":4F76
            colorFocus      =   "TableLayout.frx":4FA0
            colorDisabled   =   "TableLayout.frx":4FCA
            colorPressed    =   "TableLayout.frx":4FF4
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
            Office2003BorderWidth=   1
         End
         Begin Threed.SSCommand Display1 
            Height          =   2055
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3625
            _Version        =   196610
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   4
            ShapeSize       =   1
         End
         Begin VB.Label AddWhat 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Touch To Add ?"
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
            Left            =   240
            TabIndex        =   31
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Table Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   4800
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Table / Floor"
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
            Height          =   3135
            Left            =   120
            TabIndex        =   29
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label10 
            Caption         =   "Table Name"
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
            TabIndex        =   28
            Top             =   4800
            Width           =   975
         End
      End
      Begin Threed.SSCommand MyTable 
         CausesValidation=   0   'False
         Height          =   960
         Index           =   0
         Left            =   960
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1693
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   16777215
         BackColor       =   12632256
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TableLayout.frx":501E
         Caption         =   "Tbl"
         AutoSize        =   2
         ButtonStyle     =   4
         Outline         =   0   'False
         Shape           =   "TableLayout.frx":8070
      End
   End
End
Attribute VB_Name = "TableLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type BitMap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private DoWork  As String
Private SyncButton As Integer
Private SetLeft As Integer
Private SetTop  As Integer

Private DeskBmp As BitMap                        ' Bitmap copy of the desktop
Private Type Join_Table
       MainTable As String
       JoinWith  As String
End Type
Private Type Cleaned_Table
       CleanThisTable As String
End Type
Private AllCleaned()        As Cleaned_Table
Private AllJoined()         As Join_Table
Private MoveTable           As Integer
Private MoveLeft            As Single
Private MoveRight           As Single
Private LT                  As Single
Private RT                  As Single
Private Tops4               As Integer
Private TableSelect         As Integer
Private Sql                 As String
Private ViewTable           As EachTable
Private ViewTables          As AllTables
Private MyFloor             As String
Private HowLongToWait       As Long
Private DisplayAll          As Boolean
Private WhichButton         As Integer
Private JoinTable           As Boolean
Private WhichTables(0 To 1) As String
Private TSelected           As Integer
Private T_ToJoin            As Integer
Private Enum CheckUsage
        MulitCheck
        OkCheck
        InuseCheck
        NewCheckSelect
End Enum
Private MoveTo As Boolean
Private MoveFrm As String
Private MoveNewTbl As String
'//var for moving tables in hostess screen
Private LeaveScreen As Boolean
Private SIU As Boolean
Private ServerTableAssigned As Boolean
Private Db                 As ADODB.Connection
Private CKSDB              As ADODB.Connection
Private MoveToANewTable    As Boolean

Private Function ChkUsage(InUsage As CheckUsage) As String
        Select Case InUsage
           Case Is = 0
              ChkUsage = "Multi Check"
           Case Is = 1
              ChkUsage = "Ok"
           Case Is = 2
              ChkUsage = "InUse"
           Case Is = 3
              ChkUsage = "New"
              STChk.ChkStartTime = ""
       End Select
End Function
Public Sub LoadHeader()
With TData
    .cols = 7
    .Clear
    .TextMatrix(0, 0) = "Table"
    .TextMatrix(0, 1) = "Time Open"
    .TextMatrix(0, 2) = "Party No"
    .TextMatrix(0, 3) = "Check No"
    .TextMatrix(0, 4) = "Server"
    .TextMatrix(0, 5) = "Chk Total"
    .TextMatrix(0, 6) = "Cust. Avg"
    .ColWidth(0) = 1550 * WidScale1024
    .ColWidth(1) = 1550 * WidScale1024
    .ColWidth(2) = 1200 * WidScale1024
    .ColWidth(3) = 1300 * WidScale1024
    .ColWidth(4) = 2500 * WidScale1024
    .ColWidth(5) = 1500 * WidScale1024
    .ColWidth(6) = 1500 * WidScale1024
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignCenterCenter
    .ColAlignment(3) = flexAlignCenterCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignLeftCenter
End With
With TData
.FixedRows = 1
End With
End Sub
Public Sub Close_Check()
On Error GoTo ErrorHandler
RemoveOpenCheck False
OrderIs = GetOrderType(HereSale)
Reset_Ck_Totals
Erase How_Paid
Check_Status = MyOrderStatus(EditCheck)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Close Check Tables ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function IsNotMyTable(TBLNo As String, WhoAMI As String) As Boolean
On Error GoTo ERH
Dim Trs As ADODB.Recordset
IsNotMyTable = False
Sql = "Select * from [Open Checks] where [Close Out Day] = 0 and [Table Number] = '" & TBLNo & "'"
If CKSDB Is Nothing Then
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, CKSDB
End If
ConnectADORS Sql, CKSDB, Trs
  If Not Trs.EOF Then
  Do While Not Trs.EOF
    If WhoAMI <> Trs![Server Name] Then
     IsNotMyTable = True
    ElseIf WhoAMI = Trs.fields("Server Name").Value Then
      IsNotMyTable = False
      Exit Do
    End If
   Trs.MoveNext
  Loop
  End If
SetNothing Trs
Exit Function
ERH:
Err.Clear
End Function
Public Function IsOpen(TBLNo As String, LoggedIn As String) As String
On Error GoTo ErrorHandler
Dim Trs As ADODB.Recordset
Dim NR As Integer
Sql = "Select * from [Open Checks] where [Close Out Day] = 0 and [Table Number] = '" & TBLNo & "'"   'left([Table Number]," & Len(TBLNo) & ") = '" & TBLNo & "'"
If CKSDB Is Nothing Then
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, CKSDB
End If
ConnectADORS Sql, CKSDB, Trs

If Not Trs.EOF Then
 If RecordsCount(Trs) > 1 Then
    Label11.Caption = "Split Check Touch Function Then Check"
    LoadHeader
      NR = 0
     Do While Not Trs.EOF
       NR = NR + 1
       TData.TextMatrix(NR, 0) = Trs![Table Number] & " " & Trs![Seat Number] & ""
       TData.TextMatrix(NR, 1) = Format(Trs![check Opened], "HH:MM ampm")
       TData.TextMatrix(NR, 2) = Trs![Number in Party]
       TData.TextMatrix(NR, 3) = Trs![Check Number]
       TData.TextMatrix(NR, 4) = Trs![Server Name]
       TData.TextMatrix(NR, 5) = Trs![Check Total]
       Trs.MoveNext
     Loop
     SetNothing Trs
     CenterObject Me, ShowChecks
     With ShowChecks
        .ZOrder 0
        .Visible = True
     End With
     IsOpen = ChkUsage(MulitCheck)
     Exit Function
 End If
    If LoggedIn = "" Then  'added for check on joined table
      IsOpen = ChkUsage(OkCheck)
    ElseIf Trs![Server Name] = LoggedIn Then
      IsOpen = ChkUsage(OkCheck)
      Check_no = Trs![Check Number]
      Number_In_Party = Trs![Number in Party]
    Else
      IsOpen = ChkUsage(InuseCheck)
      See_Msg "Currently In Use", 5, "By " & Trs![Server Name]
        Check_no = Trs![Check Number]
        Number_In_Party = Trs![Number in Party]
        LoggedIn = Trs![Server Name]
       If Can_Continue("Managers Menu") Then
         IsOpen = ChkUsage(OkCheck)
         LoggedOnAs(0).ManagerOk = ""
         LoggedOnAs(0).EMPName = LoggedIn
         UserVoid.Manager = ""
         '// you have to reset all back to normal or the ownership will change
         '// for the check
         LeaveScreen = True
       Else
         ReleaseRegisterLock
         DbClose Nothing, Trs
         Exit Function
        End If
    End If
Else
   IsOpen = ChkUsage(NewCheckSelect)
   Check_Status = MyOrderStatus(NewCheck)
End If
SetNothing Trs
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Is Open Table Layout ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetOpenChecks(Tables() As Variant)
'Must be changed to server path upon add to program
Dim R As Integer
Dim RS As ADODB.Recordset
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Function
 End If
End If
If CKSDB Is Nothing Then
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, CKSDB
End If
If DisplayAll = True Then
    Sql = "Select * from [Open Checks] where [Close Out Day] = 0"
Else
  If SvrPath.GroupAudit Then
    Sql = "Select * from [Open Checks] Where [Close Out Day] = 0 and [Server Name] = '" & LoggedOnAs(0).EmpAssignedTo & "'"
  Else
    Sql = "Select * from [Open Checks] Where [Close Out Day] = 0 and [Server Name] = '" & LoggedOnAs(0).EMPName & "'"
  End If
End If
ConnectADORS Sql, CKSDB, RS
ReDim Tables(8, RecordsCount(RS))
  R = -1
  Do While Not RS.EOF
      R = R + 1
      Tables(0, R) = RS![Table Number]
      Tables(1, R) = Format(ElapsedMinutes(CDate(RS![Order Date]) & " " & CDate(RS![check Opened]), Now), "00") & " Mins"
      Tables(4, R) = RS![Server Name]
      Tables(3, R) = RS![Check Number]
      Tables(2, R) = RS![Number in Party]
      Tables(5, R) = ElapsedMinutes(CDate(RS![Order Date]) & " " & CDate(RS![check Opened]), Now)
      Tables(6, R) = Val(RS![Presented] & "")
      Tables(7, R) = Format(Val(RS![Check Total] & ""), "0.00")
      If Val(Tables(2, R)) > 0 And Val(RS![Check Total] & "") > 0 Then
        Tables(8, R) = Format(Val(RS![Check Total] & "") / Val(Tables(2, R)), "0.00")
      Else
        Tables(8, R) = "0.00"
      End If
    RS.MoveNext
  Loop
SetNothing RS
End Function
Private Sub bmpfile_Click()
If bmpfile.ListIndex > -1 Then
  Display1.Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & bmpfile.List(bmpfile.ListIndex))
End If
End Sub
Private Function CheckTableName(NewName As String) As Boolean
Dim TT As Integer
'Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
'If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Function
For TT = 0 To MyTable.count - 1
    If NewName = MyTable(TT).Caption Then
      CheckTableName = True
      Exit Function
    Else
      CheckTableName = False
    End If
Next TT
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Function
 End If
End If

'ConnectADODB SvrPath.TerminalPath & "MSaleData", DB
Sql = "Select * from [Layout] where [Table Name] = '" & NewName & "'"
Sql = Sql & " and [Layout Name] = '" & LayoutName.Text & "'"
ConnectADORS Sql, Db, Tb
If Not Tb.EOF Then
  CheckTableName = True
Else
  CheckTableName = False
End If
SetNothing Tb
End Function
Private Function TotalTables() As Long
Dim Tb As ADODB.Recordset
Sql = "Select * from [Layout]"
ConnectADORS Sql, Db, Tb
TotalTables = RecordsCount(Tb)
SetNothing Tb
End Function
Private Sub BtnEnh1_Click()
If Can_Continue("Split Check") = False Then Exit Sub
Check_Status = MyOrderStatus(SplitCheck)
End Sub
Private Sub BtnEnh10_Click()
TBLUTIL.Visible = False
TBLUTIL.Left = -5000
End Sub
Private Sub BtnEnh11_Click()
ServerMenu.Visible = True
End Sub

Private Sub BtnEnh12_Click()
 If MoveToANewTable Then
   MoveToANewTable = False
   lblStatus.Caption = ""
   See_Msg "Table Move Canceled", 5, ""
   BtnEnh12.Caption = "Move Table"
 Else
  MoveToANewTable = True
  lblStatus.Caption = "Select table to move FROM"
  BtnEnh12.Caption = "Cancel Move Table"
 End If
End Sub

Private Sub BtnEnh2_Click()
SyncButton = 0
DoWork = "AlignLeft"
End Sub
Private Sub BtnEnh3_Click()
If See_Msg("Are you Sure?", 4, "") <> 6 Then Exit Sub
Call ClearData(SvrPath.TerminalPath, "Table Layout.MDB")
If BlendData(SvrPath.TerminalPath, "MSaleData", SvrPath.TerminalPath, "Table Layout.MDB", False, "") Then
  See_Msg "Export Finished....", 5, ""
Else
  See_Msg "Export Failed....", 5, ""
End If
End Sub
Private Sub BtnEnh4_Click()
If See_Msg("Are you Sure?", 4, "") <> 6 Then Exit Sub
Call ClearData(SvrPath.TerminalPath, "MSaleData")
If BlendData(SvrPath.TerminalPath, "Table Layout.MDB", SvrPath.TerminalPath, "MSaleData", False, "") Then
  See_Msg "Import Finished....", 5, ""
Else
  See_Msg "Import Failed....", 5, ""
End If
End Sub
Private Sub BtnEnh5_Click()
CenterObject Me, TBLUTIL
TBLUTIL.Visible = True
TBLUTIL.ZOrder 0
End Sub
Private Sub BtnEnh6_Click()
SyncButton = 0
DoWork = "AlignTop"
End Sub
Private Sub BtnEnh7_Click()
SyncButton = 0
DoWork = "SnapTop"
End Sub
Private Sub BtnEnh8_Click()
SyncButton = 0
DoWork = "SnapLeft"
End Sub
Private Sub BtnEnh9_Click()
SyncButton = 0
DoWork = "Nothing"
End Sub

Private Sub Display1_Click()
On Error GoTo ErrorHandler
Dim NewTableName As String
If Left(bmpfile.List(bmpfile.ListIndex), 4) = "TBL-" Then
    If Trim(TableName.Text) = "" Then
      Call See_Msg("Table Name Required", 5, "")
      Exit Sub
    Else
      NewTableName = TableName.Text
      If IsNumeric(TableName.Text) Then
         TableName.Text = Val(TableName.Text) + 1
      End If
    End If
    If CheckTableName(NewTableName) = True Then
      Call See_Msg("Duplicate Name Found", 0, "Cannot Have A Duplicate")
      Exit Sub
    End If
NxtTbl:
    Tops4 = Tops4 + 1
    On Error Resume Next
    Load MyTable(Tops4)
    On Error GoTo ErrorHandler
    MyTable(Tops4).Caption = NewTableName
    MyTable(Tops4).CaptionStyle = ssWrapped
    MyTable(Tops4).Left = 120 '* WidScale
    MyTable(Tops4).Top = 120 '* TopScale
    MyTable(Tops4).Tag = CStr(Cover.Value)
    If Display1.Picture <> 0 Then
        MyTable(Tops4).Picture = Display1.Picture
        MyTable(Tops4).Shape = Display1.Picture
        MyTable(Tops4).TagVariant = bmpfile.List(bmpfile.ListIndex)
        MyTable(Tops4).ToolTipText = bmpfile.List(bmpfile.ListIndex)
    End If
    MyTable(Tops4).Visible = True
    MyTable(Tops4).ZOrder 0
    Sleep 0
End If
Exit Sub
ErrorHandler:
GoTo NxtTbl
End Sub
Private Sub Form_Activate()
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Sub
 End If
End If
If CKSDB Is Nothing Then
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, CKSDB
End If
HowLongToWait = 0
LeaveScreen = False
LoadScreen
If LoggedOnAs(0).EMPName = "Hostess Station" Then
  OpenTableOnly = True
End If
If LoggedOnAs(0).EmpAccessLevel >= Access_Required("Managers Menu") Then
 Status.Visible = True
 Status.ZOrder 0
Else
 BtnEnh1.Visible = True
 BtnEnh1.ZOrder
End If
LoadJoined
LoadCleaned
TimersOn
End Sub
Private Sub MoveNewTable(NewCheckNo As Long, NewTable As String, SvrName As String)
    Dim CHTN As GeneralLib.Utilities
    Set CHTN = New GeneralLib.Utilities
    GetSelected = CHTN.ChangeTabName(CStr(NewCheckNo), NewTable, SvrName)
    If GetSelected = "Success" Then
     See_Msg "Update Complete", 5, ""
    Else
     See_Msg "Problem", 5, GetSelected
    End If
    GetSelected = ""
    Set CHTN = Nothing
End Sub
Private Sub TimersOn()
Timer1.Enabled = True
IamWaiting.Enabled = True
End Sub
Private Sub TimersOFF()
Timer1.Enabled = False
IamWaiting.Enabled = False
End Sub
Public Sub AddToServerList(TBLNo As String)
Dim lf As Integer
For lf = 0 To TblList.ListCount - 1
    If TblList.List(lf) = TBLNo Then Exit Sub
Next lf
TblList.AddItem TBLNo
End Sub
Public Function AssignedTableList(SvrName As String)
On Error GoTo ErrorHandler
'Dim DB As adodb.Connection
Dim RS As ADODB.Recordset
TblList.Clear
If SvrName = "" Then Exit Function
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Function
 End If
End If

'If ConnectADODB(SvrPath.TerminalPath & "MSaleData", DB) = False Then Exit Function
Sql = "Select *  from [Layout] Where [Assigned To] = '" & SvrName & "'"
Sql = Sql & " Order by [Table Name] asc"
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
Do While Not RS.EOF
   TblList.AddItem RS![Table Name] & ""
 RS.MoveNext
Loop '
'DbClose DB, RS
SetNothing RS
'Set GRS = Nothing
Exit Function
ErrorHandler:
Err.Clear
End Function
Public Sub UnLoadTables()
On Error GoTo ErrorHandler
Room.Visible = False
If MyTable.count >= 1 Then
    For Ul = 0 To MyTable.count - 1
        MyTable(Ul).Picture = Nothing
        MyTable(Ul).Shape = Nothing
        MyTable(Ul).TagVariant = ""
        MyTable(Ul).Tag = ""
        MyTable(Ul).Visible = False
NextUl:
    Next Ul
    Tops4 = 0
End If
Room.Visible = True
'Set ViewTables = New AllTables
Exit Sub
ErrorHandler:
Error_Log Err.Number, Err.Description, "UnloadTables", ""
Err.Clear
Room.Visible = True
GoTo NextUl
End Sub
Public Function SectionInUse(SvrOn As String) As Boolean
Dim tTb As ADODB.Recordset
'If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Function
'ConnectADODB SvrPath.TerminalPath & "MSaleData", TDb
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Function
 End If
End If
 
 Sql = "Select * from [Layout] where [Assigned To] = '" & SvrOn & "'"
ConnectADORS Sql, Db, tTb
If tTb.EOF Then
  SectionInUse = False
Else
  SectionInUse = True
End If
DbClose Nothing, tTb
'DbClose TDb, tTb
End Function
Public Sub LoadSections(ShowAll As Boolean)
' Layout editor mode — POS runtime removed
On Error GoTo ErrorHandler
Dim Tb As ADODB.Recordset
UnLoadTables
Room.Visible = False
If Db Is Nothing Then
  If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
    See_Msg "Cannot Connect To Table Database", 0, ""
    Exit Sub
  End If
End If
Sql = "Select * from [Layout] where [Layout Name] = '" & GV.WorkArea & "'"
'Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & "MSaleData"), "Layout")
'If Not Tb Is Nothing Then
'  If Not Tb.BOF Then Tb.MoveFirst
'End If

ConnectADORS Sql, Db, Tb
If Tb.EOF Then
  SetNothing Tb
  Exit Sub
End If
' Apply background colour from stored style ID
SetBgColor Tb![background] & ""
Tops4 = RecordsCount(Tb)
Dim LT As Integer
LT = 0
Do While Not Tb.EOF
  If Tb![Picture] & "" > "" Then
    LT = LT + 1
    On Error Resume Next
    Load MyTable(LT)
    SetShapeAppearance MyTable(LT), Tb![Picture] & ""
    MyTable(LT).TagVariant = Tb![Picture] & ""
    MyTable(LT).width  = CLng(Tb![width])  * 15
    MyTable(LT).Height = CLng(Tb![Height]) * 15
    MyTable(LT).Top    = CLng(Tb![Top])    * 15
    MyTable(LT).Left   = CLng(Tb![Left])   * 15
    MyTable(LT).Caption = Tb![Table Name] & ""
    MyTable(LT).Tag = Check_Null(Tb![Party Number])
    MyTable(LT).Visible = True
  End If
  Tb.MoveNext
Loop
Room.Visible = True
SetNothing Tb
Exit Sub
ErrorHandler:
Error_Log Err.Number, Err.Description, "Load Sections TBL", ""
Resume Next
End Sub
Private Sub Form_Deactivate()
  ' Removed - POS runtime only
End Sub

Public Sub Form_Load()
  ' Layout editor only — POS runtime removed
  EditLayout = True
  LoadTableScreen
End Sub
Public Sub LoadTableScreen()
' Layout editor only — POS runtime removed
On Error GoTo ErrorHandler
Dim Tb As ADODB.Recordset
If Db Is Nothing Then
  If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
    See_Msg "Cannot Connect To Table Database", 0, ""
    Exit Sub
  End If
End If
GV.WorkArea = ""
With TableData
   .Left = 1440 * WidScale1024
   .Top = 0
   If bmpfile.ListCount > -1 Then bmpfile.ListIndex = 1
   .Visible = True
End With
EditBox.Left = 2520: EditBox.Top = 9960 * TopScale1024
EditBox.Visible = True
Sql = "Select Distinct [Layout Name] from [Layout]"
ConnectADORS Sql, Db, Tb
If Not Tb.BOF Then Tb.MoveFirst
Do While Not Tb.EOF
  LayoutName.AddItem Tb![Layout Name] & ""
  Tb.MoveNext
Loop
SetNothing Tb
If Trim(GV.WorkArea) > "" Then
   LayoutName.Text = GV.WorkArea
   LayoutName_Click
End If
SSCommand6_Click
Exit Sub
ErrorHandler:
Error_Log Err.Number, Err.Description, "Load Layout", ""
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetTablesToNothing
End Sub

Private Sub IamWaiting_Timer()
  ' Removed - POS runtime only
End Sub

Private Sub LayoutName_Click()
Dim TTLTable As Integer
On Error GoTo ErrorHandler
Dim Ul As Integer
Dim Tb As ADODB.Recordset
If See_Msg("Change Section?", 4, "") <> 6 Then Exit Sub
UnLoadTables
TableLayout.Refresh
If LayoutName.ListIndex = -1 Then
  Sql = "Select * from [Layout] where [Layout Name] = '" & GV.WorkArea & "'"
Else
  Sql = "Select * from [Layout] where [Layout Name] = '" & LayoutName.List(LayoutName.ListIndex) & "'"
End If
ConnectADORS Sql, Db, Tb
TTLTable = RecordsCount(Tb)
Tops4 = TTLTable
Ul = 0
If Not Tb.EOF Then
    If Tb![background] & "" > "" Then
      Room.Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & Tb![background])
      MyFloor = Tb![background] & ""
    End If
End If
Ul = -1
Do While Not Tb.EOF
    If Tb![Picture] & "" > "" Then
        Ul = Ul + 1
        Load MyTable(Ul)
        MyTable(Ul).Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & Tb![Picture])
        MyTable(Ul).ToolTipText = Tb![Picture]
        'MyTable(Ul).Picture = SvrPath.LocalTerminalPath & Tb![Picture]
        MyTable(Ul).Shape = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & Tb![Picture])
        MyTable(Ul).TagVariant = Tb![Picture] & ""
        MyTable(Ul).width = Tb![width] '* WidScale
        MyTable(Ul).Height = Tb![Height] '* TopScale
        MyTable(Ul).Top = Tb![Top] '* TopScale
        MyTable(Ul).Left = Tb![Left] '* WidScale
        MyTable(Ul).Caption = Tb![Table Name]
        MyTable(Ul).Tag = CStr(Tb![Party Number])
        MyTable(Ul).Visible = True
    End If
 Tb.MoveNext
Loop
TableLayout.ZOrder 0
Tops4 = Ul
SetNothing Tb
'Set GRS = Nothing
Exit Sub
ErrorHandler:
Error_Log Err.Number, Err.Description, "Layout Name click", ""
Resume Next
End Sub
Private Sub LoadFloor_Click()
On Error Resume Next
bmpfile.Pattern = "FLR-" & "*.*"
bmpfile.path = SvrPath.LocalTerminalPath & "Picture Files\"
AddWhat.Caption = "Touch to change floor"
Label3.Caption = "Flooring"
With TableData
   .Left = 1440 * WidScale1024
   .Top = 0 * TopScale1024
   .Visible = True
End With
If bmpfile.ListCount > -1 Then
    bmpfile.ListIndex = 0
    Display1.Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & bmpfile.List(0))
End If
End Sub
Private Sub LoadTables_Click()
bmpfile.Pattern = "TBL-" & "*.*"
bmpfile.path = SvrPath.LocalTerminalPath & "Picture Files\"
With TableData
   .Left = 1440 * WidScale1024
   .Top = 0 * TopScale1024
   .Visible = True
End With
'bmpfile.Refresh
If bmpfile.ListCount > 0 Then
    bmpfile.ListIndex = 0
    Display1.Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & bmpfile.List(0))
End If
AddWhat.Caption = "Touch to add table"
Label3.Caption = "Tables"
TableName.Text = ""
Cover.Value = ssCBChecked
End Sub
Private Sub Picture1_Click()
End Sub

Private Sub Room_Click()
On Error Resume Next
HowLongToWait = 0
If EditLayout = False Then Exit Sub
If bmpfile.ListIndex > -1 And Left(bmpfile.List(bmpfile.ListIndex), 3) = "FLR" Then
    Room.Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & bmpfile.List(bmpfile.ListIndex))
    Room.Tag = bmpfile.List(bmpfile.ListIndex)
    MyFloor = bmpfile.List(bmpfile.ListIndex)
End If
End Sub
Public Sub Get_Tables()
If EditLayout = True Then Exit Sub
On Error GoTo ErrorHandler
Dim OpenDate        As Date
Dim MinsOpen        As Long
Dim CRS             As ADODB.Recordset
Dim Tb              As ADODB.Recordset
Dim ChkOnHold       As Boolean
Dim PreAmt          As Currency
Dim MyTables        As String
Dim TBn             As Integer
Dim HoldChecks()    As Long
Dim DataStr         As String
Dim CN              As Long
Dim PRAdb           As ADODB.Connection
Set ViewTable = Nothing
Set ViewTables = Nothing
Set EachTable = Nothing
Set AllTables = Nothing
Set ViewTables = New AllTables
TimersOFF
Dim TableClean As Boolean
DataStr = "[Check Opened],[Check Number],[Presented],[Table Number],[Order Date],[Number in Party],[Gross Sales],[Payments]"
Sql = "Select " & DataStr & " from [Open Checks] Where [Close Out Day] = 0 Group By " & DataStr & " order by [Table Number] asc"
If CKSDB Is Nothing Then
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, CKSDB
End If
ConnectADORS Sql, CKSDB, Tb

If Check_Option(71) = 1 Then '//option 71 is hold menu items
    Sql = "Select [Check Number] from [Check Detail] where [On Hold] > 0 Group By [Check Number]"      'and [Check Number]= " & TryCheck
   ConnectADORS Sql, CKSDB, CRS
   ChkOnHold = False
   If Not CRS.EOF Then
     ReDim HoldChecks(RecordsCount(CRS))
     ChkOnHold = True
   End If
   CN = -1
   Do While Not CRS.EOF
     CN = CN + 1
     HoldChecks(CN) = CRS![Check Number]
     CRS.MoveNext
   Loop
   DbClose Nothing, CRS
End If

Room.Visible = False
Do While Not Tb.EOF
    MyTables = Trim$(Tb.fields("Table Number").Value & "")
    If ViewTables.Item(MyTables) Is Nothing Then
       Set ViewTable = ViewTables.add(MyTables)
    End If
    With ViewTable
      If Not IsDate(Tb.fields("Order Date").Value & "") Then
          OpenDate = Format(Date, "mm/dd/yyyy")
      Else
          OpenDate = Format(Tb.fields("Order Date").Value & "", "mm/dd/yyyy")
      End If
     .T_CheckNo = Val(Tb.fields("Check Number").Value & "")
     .T_Presented = Val(Tb.fields("Presented").Value & "")
     .T_Table = MyTables
     .T_StartTime = Format(ElapsedMinutes(OpenDate & " " & CDate(Tb.fields("Check Opened").Value), Now), "00")
     .T_Color = vbBlue
     If Tb![Gross Sales] = 0 Then
            .T_Color = vbYellow
     End If
     'If ElapsedMinutes(OpenDate & " " & CDate(Tb.fields("Check Opened").Value), Now) >= Val(Input_Option(2)) Then
     '   .T_Color = vbRed
     'End If
     If Tb.fields("Presented").Value = 1 Then
        .T_Color = &HC000C0
     End If
     If Check_Option(120) = 1 Then
      If Tb.fields("Payments").Value > 0 Then
       .T_Color = &H80C0FF
      End If
     End If
    .T_Party = Val(Tb![Number in Party] & "")
    .T_Hold = 0
    If Check_Option(71) = 1 Then   ' option 71 = hold menu items
        If ChkOnHold Then
            For Ck = LBound(HoldChecks) To UBound(HoldChecks)
                If Tb![Check Number] = HoldChecks(Ck) Then
                    .T_Hold = 1
                    Exit For
                End If
            Next Ck
        End If
    End If
    End With
  Tb.MoveNext
Loop
Room.Visible = True
TimersOn
SetNothing Tb
Exit Sub
ErrorHandler:
  Error_Log Err.Number, Err.Description, "Layout get tables", ""
  'See_Msg "Error Loading Tables", 5, ""
  'DbClose Db, Tb
  Room.Visible = True
  Err.Clear
  Exit Sub
End Sub
Private Sub LoadCleaned()
If Check_Option(110) = 0 Then Exit Sub
'Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Counter As Integer
'ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Sub
 End If
End If
Sql = "Select * from [Cleared] "
ConnectADORS Sql, Db, RS
ReDim AllCleaned(RecordsCount(RS))
If RS.EOF Then
  DbClose Nothing, RS
  Exit Sub
End If
Do While Not RS.EOF
 AllCleaned(Counter).CleanThisTable = RS![Table Number] & ""
 Counter = Counter + 1
 RS.MoveNext
Loop
DbClose Nothing, RS
End Sub
Private Function Cleaned(tableN As String) As Boolean
If Check_Option(110) = 0 Then Cleaned = True: Exit Function
Dim Counter As Integer
For Counter = 0 To UBound(AllCleaned, 1)
 If AllCleaned(Counter).CleanThisTable = tableN Then
    Cleaned = False
    Exit Function
 End If
Next Counter
Cleaned = True
End Function
Private Sub Cleared(tableN As String)
'dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
'If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Sub

'ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Sub
 End If
End If

Sql = "Delete from [Cleared] Where [Table Number] = '" & tableN & "'"
Db.Execute Sql
DbClose Nothing, RS
Dim Counter As Integer
For Counter = 0 To UBound(AllCleaned, 1)
 If AllCleaned(Counter).CleanThisTable = tableN Then
    AllCleaned(Counter).CleanThisTable = ""
    Exit For
 End If
Next Counter
End Sub
Private Sub Room_DragDrop(Source As Control, x As Single, Y As Single)
If EditLayout = False Then Exit Sub
If MoveTable <> -1 Then
With MyTable(MoveTable)
   .Move x - LT, Y - RT, .width, .Height
End With
Else
  With TableData
  .Move x - LT, Y - RT, .width, .Height
  End With
End If
End Sub
Private Sub Room_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler
If EditLayout = False Then Exit Sub
MoveLeft = x: MoveRight = Y
If MoveTable > -1 Then
MyTable(MoveTable).Drag vbEndDrag
Else
TableData.Drag vbEndDrag
End If
Exit Sub
ErrorHandler:
Err.Clear
End Sub

Private Sub ServerMenu_Click()

End Sub

Private Sub ShowAssign_Click()
HostessMenu.Top = 0
HostessMenu.Left = 13200 * WidScale1024
HostessMenu.ZOrder 0
HostessMenu.Visible = True
End Sub

Private Sub SSCommand1_Click()
On Error Resume Next
If LoggedOnAs(0).EMPName = "Reservations" Then Exit Sub
HowLongToWait = 0
'ChangeSection.Show 1
Load_What = "Change Sections"
Find_What = "Choose a Section"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
GV.WorkArea = GetSelected: GetSelected = ""
If GV.WorkArea > "" Then
  TableLayout.LoadSections False
End If
End Sub
Private Sub OpenTable_Click()
On Error GoTo ERH
If LoggedOnAs(0).EMPName = "Reservations" Then Exit Sub
If Can_Continue(OpenTable.Caption) = False Then Exit Sub
HowLongToWait = 0
ReadyTo.Caption = "Add New Check"
See_Msg "Touch Table", 5, ""
Exit Sub
ERH:
Err.Clear
End Sub
Private Sub SelectNo_Click()
End Sub
Private Sub SSCommand10_Click()
If LoggedOnAs(0).EMPName = "Reservations" Then Exit Sub
If LoggedOnAs(0).EMPName = "Hostess Station" Then
 See_Msg "No Access", 5, ""
Exit Sub
End If
If Input_Option(10) = "Open Only" Then
   See_Msg "Open Only Station", 5, "Cannot close check from this station"
   Exit Sub
End If
HowLongToWait = 0
ReadyTo.Caption = "Close Check"
If ReadyTo.Caption > "" Then
  SSCommand8.Caption = "Cancel"
End If
Check_Status = MyOrderStatus(CloseCheck)
End Sub
Private Sub SSCommand11_Click()
  ' Removed - POS runtime only
End Sub
Private Sub SSCommand16_Click()
HostessMenu.Top = 0
HostessMenu.Left = 13200 * WidScale1024
HostessMenu.ZOrder 0
HostessMenu.Visible = True
End Sub
Private Sub SSCommand12_Click()
Dim SR As Integer
With SSCommand12
 Select Case .Caption
     Case Is = "Save Assignment"
         .Caption = "Assign Tables"
         If TblList.ListCount > -1 Then
           If SvrsName.Caption = "" Then
             See_Msg "Select A Server", 5, ""
             Exit Sub
           End If
           If See_Msg("Save Table Assignment ?", 4, "") = 6 Then
               ResetAssigned SvrsName.Caption
               For SR = 0 To TblList.ListCount - 1
                  AssignTableTo TblList.List(SR), SvrsName.Caption
               Next SR
           End If
         See_Msg "Assignment Complete", 5, ""
         SvrsName.Caption = ""
         TblList.Clear
         End If
         Exit Sub
     Case Else
        See_Msg "Select a Server", 5, ""
        Exit Sub
 End Select
End With
End Sub
Private Sub SSCommand13_Click()
                Load_What = "AllTiped"
                Find_What = "Select a Server"
                RequestListForm.Show 1
                ReleaseRegisterLock
                If GetSelected = "" Then Exit Sub
                SvrsName.Caption = GetSelected
                AssignedTableList SvrsName.Caption
                SSCommand12.Caption = "Save Assignment"
End Sub
Private Sub SSCommand14_Click()
With HostessMenu
  .Visible = False
End With
End Sub
Private Sub SSCommand15_Click()
If Check_Option(150) = 1 Then
  WaitList.Show 1
Else
  See_Msg "No List Set", 5, ""
End If
End Sub
Private Sub SSCommand17_Click()
  ' Removed - POS runtime only
End Sub
Private Sub SSCommand2_Click()
If See_Msg("Delete Table", 4, "Are you sure ??") = 6 Then
If TableSelect > 0 Then
 Unload MyTable(TableSelect)
 TableSelect = -1
End If
End If
End Sub
Private Sub save_Click()
On Error Resume Next
If Trim(LayoutName.Text) = "" Then
  Call See_Msg("No Name", 0, "Must Have a Name")
  Exit Sub
End If
If See_Msg("Save Layout", 4, "Save Layout") <> 6 Then Exit Sub
Dim Tb As ADODB.Recordset
Dim TTLTable As Integer
Dim St As Integer
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Sub
 End If
End If

'ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
 Sql = "delete from [Layout] where [Layout Name] = '" & LayoutName.Text & "'"
Db.Execute Sql
 Sql = "Select * from [Layout] where [Layout Name] = '" & LayoutName.Text & "'"
ConnectADORS Sql, Db, Tb
For St = 0 To Tops4
 If MyTable(St).width > 0 Then
  Tb.AddNew
    Tb![Picture] = NormalizePicture(CStr(MyTable(St).TagVariant))
    Tb![Job Code] = LayoutName.Text
    Tb![Layout Name] = LayoutName.Text
    Tb![Index Number] = St
    Tb![width] = MyTable(St).width \ 15
    Tb![Height] = MyTable(St).Height \ 15
    Tb![Top] = MyTable(St).Top \ 15
    Tb![Left] = MyTable(St).Left \ 15
    Tb![Table Name] = MyTable(St).Caption
    Tb![Party Number] = Val(MyTable(St).Tag)
    Tb![background] = NormalizePicture(MyFloor)   'Room.Tag
 End If
Tb.Update
Next St
SetNothing Tb
'Set Db = Nothing
See_Msg "Saved", 5, "Saved..."
End Sub
Private Sub SSCommand22_Click()
  ' Removed - POS runtime only
End Sub
Private Sub SSCommand3_Click()
  ' Removed - POS runtime only
End Sub
Private Sub SSCommand4_Click()
With HostessMenu
  .Left = 0
  .Top = 0
  .ZOrder 0
  .Visible = True
End With
End Sub
Private Sub SSCommand5_Click()
Unload TableLayout
EditLayout = False
End Sub
Private Sub SSCommand6_Click()
TableData.Visible = False
TableData.Left = -4000
TableData.Refresh
bmpfile.ListIndex = -1
JoinTable = False
TSelected = -1
End Sub
Private Sub SSCommand7_Click()
  ' Removed - POS runtime only
End Sub
Private Sub SSCommand8_Click()
HowLongToWait = 0
Check_Status = MyOrderStatus(CancelCheck)
If SSCommand8.Caption = "Cancel" Then
  ReadyTo.Caption = ""
  SSCommand8.Caption = "Exit"
  Exit Sub
Else
  'Call MemoryCheck(TableLayout)
  Unload Me
End If
End Sub
Private Sub StartOver_Click()
If See_Msg("Clear Layout?", 4, "") <> 6 Then Exit Sub
UnLoadTables
TableLayout.Refresh
End Sub
Private Sub SSCommand9_Click()
If See_Msg("Delete Room", 4, "Are you sure ??") = 6 Then
    Dim Tb As ADODB.Recordset
    If Db Is Nothing Then
      If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
       See_Msg "Cannot Connect To Table Database", 0, ""
       Exit Sub
      End If
    End If
    Sql = "Delete from [Layout] where [Layout Name] = '" & LayoutName.Text & "'"
    Db.Execute (Sql)
    StartOver_Click
End If
End Sub
Private Sub Status_Click()
HowLongToWait = 0
ReadyTo.Caption = "Table Status"
If ReadyTo.Caption > "" Then
  SSCommand8.Caption = "Cancel"
End If
Dim allcheck() As Variant
On Error Resume Next
ReadyTo.Caption = "Table Review"
If ReadyTo.Caption > "" Then
  SSCommand8.Caption = "Cancel"
End If
Dim CT As Currency
Dim CuT As Currency
CT = 0: CuT = 0
Dim La As Integer
Dim PresentedChk As Integer
Dim SeatedChk As Integer
Dim OpenChk   As Integer
If OpenTableOnly Then DisplayAll = True
GetOpenChecks allcheck()
Label11.Caption = "Open Check Status"
LoadHeader
PresentedChk = 0
SeatedChk = 0
OpenChk = 0
With TData
    For La = 0 To UBound(allcheck, 2) - 1
        If allcheck(6, La) = 1 Then
           .Cell(flexcpForeColor, La + 1, 0, La + 1, 6) = &HC000C0
           PresentedChk = PresentedChk + 1
        ElseIf allcheck(8, La) = 0 Then
           .Cell(flexcpForeColor, La + 1, 0, La + 1, 6) = &HC0C0&
           SeatedChk = SeatedChk + 1
        Else
           .Cell(flexcpForeColor, La + 1, 0, La + 1, 6) = &HFF0000
           OpenChk = OpenChk + 1
        End If
        .TextMatrix(La + 1, 0) = allcheck(0, La)
        .TextMatrix(La + 1, 1) = allcheck(1, La)
        .TextMatrix(La + 1, 2) = allcheck(2, La)
        .TextMatrix(La + 1, 3) = allcheck(3, La)
        .TextMatrix(La + 1, 4) = allcheck(4, La)
        .TextMatrix(La + 1, 5) = allcheck(7, La)
        .TextMatrix(La + 1, 6) = allcheck(8, La)
        CT = CT + Val(allcheck(7, La))
        CuT = CuT + Val(allcheck(8, La))
    Next La
    La = La + 1
        .TextMatrix(La + 1, 4) = "Totals"
        .TextMatrix(La + 1, 5) = Format(CT, "0.00")
        .TextMatrix(La + 1, 6) = Format(CuT, "0.00")
        PCHK.Caption = "Check Down  " & Format(PresentedChk, "0")
        SCHK.Caption = "Seated  " & Format(SeatedChk, "0")
        FCHK.Caption = "W / Food  " & Format(OpenChk, "0")
        OCHK.Caption = "Ready  " & Format(TotalTables - (PresentedChk + SeatedChk + OpenChk), "0")
    .ZOrder 0
End With
If UBound(allcheck, 2) = 0 Then
 See_Msg "No Checks Found", 0, "All Tables Are Available"
 Exit Sub
Else
CenterObject Me, ShowChecks
With ShowChecks
   .ZOrder 0
   ShowChecks.Visible = True
End With
ReadyTo.Caption = ""
SSCommand8.Caption = "Exit"
ShowChecks.ZOrder 0
End If
End Sub
Private Function JoinedTableSave() As Boolean
'Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
If WhichTables(0) = WhichTables(1) Then
  JoinedTableSave = False
  TSelected = -1
  JoinTable = False
  Exit Function
End If
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Function
 End If
End If

Sql = "Select * from [Joined Tables] Where [Joined To] = '" & WhichTables(1) & "'"
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
If Not RS.EOF Then
  JoinedTableSave = False
  See_Msg "Main Table Already In Use", 5, ""
  JoinTable = False
  TSelected = -1
  DbClose Nothing, RS
  Exit Function
End If
SetNothing RS

Sql = "Select * from [Joined Tables] Where [Main Table] = '" & WhichTables(0) & "'"
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
If Not RS.EOF Then
  JoinedTableSave = False
  See_Msg "Table Already Joined", 5, ""
  JoinTable = False
  TSelected = -1
  DbClose Nothing, RS
  Exit Function
End If
SetNothing RS
JoinedTableSave = True
Sql = "Select * from [Joined Tables] Where [Main Table] = '" & WhichTables(1) & "'"
Sql = Sql & " and [Joined To] = '" & WhichTables(0) & "'"
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
   If RS.EOF Then RS.AddNew
   RS![Main Table] = WhichTables(1)
   RS![Joined To] = WhichTables(0)
   RS.Update
DbClose Nothing, RS
End Function
Private Function GetTableName(InData As Variant) As String
    If InStr(InData, vbCrLf) > 0 Then
     GetTableName = Left(InData, InStr(InData, vbCrLf) - 1)
    Else
     GetTableName = InData
    End If
End Function
Private Function JoinedTableGet(TableName As String) As String
Dim Counter As Integer
For Counter = 0 To UBound(AllJoined, 1)
   If AllJoined(Counter).JoinWith = TableName Then
      JoinedTableGet = AllJoined(Counter).MainTable
      Exit Function
   End If
Next Counter
JoinedTableGet = ""
End Function
Private Sub LoadJoined()
Dim Counter As Integer
'Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Sql = "Select * from [Joined Tables] "
'ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
If Db Is Nothing Then
 If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then
   See_Msg "Cannot Connect To Table Database", 0, ""
   Exit Sub
 End If
End If

ConnectADORS Sql, Db, RS
ReDim AllJoined(RecordsCount(RS))
If RS.EOF Then
  DbClose Nothing, RS
  Exit Sub
End If
Counter = 0
 Do While Not RS.EOF
   AllJoined(Counter).MainTable = RS![Main Table] & ""
   AllJoined(Counter).JoinWith = RS![Joined To] & ""
   Counter = Counter + 1
  RS.MoveNext
 Loop
 DbClose Nothing, RS
End Sub
Private Sub SetTablesToNothing()
Set ViewTable = Nothing
Set ViewTables = Nothing
Set EachTable = Nothing
Set AllTables = Nothing
Set GRS = Nothing
DbClose Db, Nothing
DbClose CKSDB, Nothing
End Sub
Private Sub MoveToANewSpot(CurrentTable As String, NewTable As String)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
If ConnectADODB(SvrPath.TerminalPath & CheckStatDb, Db) = False Then Exit Sub
Sql = "Update [Open Checks] set [Table Number] = '" & NewTable & "' Where [close out day] = 0 and [Table Number] = '" & CurrentTable & "'"
Db.Execute Sql
Exit Sub
ErrorHandler:
See_Msg Err.Description, 0, "Error"
End Sub

Private Sub MyTable_Click(Index As Integer)
' Layout editor only — POS runtime removed
If EditLayout = True Then EditTables Index
Exit Sub
ERH_DISABLED:
Dim TBLName As String
TBLName = GetTableName(MyTable(Index).Caption)
 If OpenTableOnly = True And SSCommand12.Caption = "Save Assignment" Then
    AddToServerList TBLName
    ReleaseRegisterLock
    TimersOn
    Exit Sub
 End If
 GenInfo.Caption = ""
    


 If MoveToANewTable = True Then
   '// this section for moving tables onces seated.
   If MoveTo Then
     If MoveNewTbl = "" And ViewTables.Item(TBLName) Is Nothing And MoveFrm > "" Then
      MoveNewTbl = TBLName
      Call See_Msg("Moving To", 5, "Table " & MoveNewTbl)
      MoveToANewSpot MoveFrm, MoveNewTbl
      MoveToANewTable = False
      TableLayout.LoadTableScreen
      TableLayout.Show
      MoveTo = False
      lblStatus.Caption = ""
      BtnEnh12.Caption = "Move Table"
      Exit Sub
     ElseIf MoveFrm > "" Then
      Call See_Msg("Cancel Move", 5, "Cannot move to a occupied table")
      MoveTo = False
      MoveFrm = ""
      MoveNewTbl = ""
      BtnEnh12.Caption = "Move Table"
      MoveToANewTable = False
     Else
       MoveFrm = TBLName
     End If
     Exit Sub
     '// end moving table section
   End If
   SvrsName.Caption = AssignedToWho
   AssignedTableList AssignedToWho
   See_Msg "Select Table To Move To", 5, ""
   lblStatus.Caption = "Select table to move To"
   MoveFrm = TBLName
   MoveTo = True
   Exit Sub
 End If
 Dim ExitSub As Boolean
  If MyTable(Index).BackColor = &H80C0FF Then
    If IsCredit(Check_no) = False Then Exit Sub
    If See_Msg("Apply Tip and Close", 4, "Apply Tip To Check") = 6 Then
      If ViewTables.Item(TBLName) Is Nothing Then Exit Sub
      Check_no = ViewTables.Item(TBLName).T_CheckNo
      Set allcards = Nothing
      Set allcards = New AllCharges
      If allcards.Item(CStr(Check_no)) Is Nothing Then
       allcards.add (CStr(Check_no))
      End If
      allcards.Item(CStr(Check_no)).IncommingCharge = 0
      ListCards.LoadEMVChecks Check_no
      ListCards.Show 1
    End If
  Close_Open_Check Check_no, LoggedOnAs(0).EMPName
  Close_Check    ' Check Title, 0 for no void
  ReleaseRegisterLock
  TimersOn
  If MyTable(Index).TagVariant > "" Then
   MyTable(Index).Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & MyTable(Index).TagVariant)
  End If
  MyTable(Index).BackColor = vbWhite
  Sleep 0
  Exit Sub
 End If
    If InStr(TBLName, "Joined") > 0 Then
      See_Msg "Table Joined", 5, ""
      JoinTable = False
      TSelected = -1
      ReleaseRegisterLock
      TimersOn
      Exit Sub
    End If
    If JoinTable Then
       If MyTable(Index).Picture = 0 And TSelected = -1 Then
          See_Msg "Table In Use", 5, ""
          JoinTable = False
          TSelected = -1
          ReleaseRegisterLock
          TimersOn
          Exit Sub
       End If
       GenInfo.Caption = "Join Table To"
       If TSelected = -1 Then T_ToJoin = Index
       TSelected = TSelected + 1
       WhichTables(TSelected) = TBLName
       ReleaseRegisterLock
       TimersOn
       If TSelected < 1 Then Exit Sub
       If TSelected = 1 Then
         JoinTable = False
         TSelected = -1
         If JoinedTableSave = True Then
            MyTable(T_ToJoin).BackColor = vbBlack
            MyTable(T_ToJoin).Caption = "Joined " & TBLName
         Else
           Exit Sub
         End If
       End If
   End If

  If Check_Option(110) = 1 Then
    If MyTable(Index).BackColor = &H8000& Then
       
       ReleaseRegisterLock
       If See_Msg("Table Clean", 4, "Ready To Seat Next Party") = 6 Then
         Cleared TBLName
         JoinedTableClose TBLName
         EraseJoinedTable TBLName
         MyTable(Index).BackColor = vbWhite
         LoadSections False
         ReleaseRegisterLock
         TimersOn
         Exit Sub
       Else
         TimersOn
         Exit Sub
       End If
     End If
  End If
  If Val(LoggedOnAs(0).EmpAccessLevel) = 0 Then ReleaseRegisterLock: Exit Sub
     If LoggedOnAs(0).EMPName = "Reservations" Then
        Table_Number = TBLName
        ReleaseRegisterLock
        Unload Me
        Exit Sub
     End If
     If ReadyTo.Caption = "Add New Check" Then
       ReadyTo.Caption = ""
       GoTo AddCheck
     Else
      ReadyTo.Caption = ""
     End If

     SSCommand8.Caption = "Exit"
     CurrentServer = LoggedOnAs(0).EMPName
     Select Case IsOpen(TBLName, CurrentServer)
         Case Is = ChkUsage(NewCheckSelect)
            If Check_Status <> MyOrderStatus(NewCheck) Then
               Call See_Msg("Check Not Open", 5, "")
               Check_Status = MyOrderStatus(EditCheck)
            End If
AddCheck:
            If OpenTableOnly = False Then
             AssignmentName = TableAssignedTo(TBLName)
             If CurrentServer <> AssignmentName And AssignmentName > "" Then
               If See_Msg("Table Assiged To Someone Else Yes To Reassign Table ", 4, "Table Assigned To " & AssignmentName) <> 6 Then ReleaseRegisterLock: Exit Sub
               
               If Can_Continue("Managers Menu") = False Then
                 ReleaseRegisterLock
                 TimersOn
                 Exit Sub
               End If
             End If
            End If
            
            If Number_In_Party = 0 Then Number_In_Party = 1
            If ServerTableAssigned Then AssignTableTo MyTable(Index).Caption, CurrentServer
            Table_Number = GetTableName(MyTable(Index).Caption)
            Check_Status = MyOrderStatus(NewCheck)
            'MyTable(Index).ToolTipText = CurrentServer
            If Check_Option(29) = 0 Then
                If Abs(Val(MyTable(Index).Tag)) = 1 Then
                   Number_In_Party = GetCovers
                Else
                   Number_In_Party = 1
                End If
            End If
            If Number_In_Party = 0 Then Check_Status = MyOrderStatus(EditCheck)
            If Number_In_Party = 0 Then
               ReleaseRegisterLock
               JoinedTableClose TBLName
               EraseJoinedTable TBLName
               LoadSections False
               TimersOn
               Exit Sub
            End If
             If OpenTableOnly Then
                AssignedToWho = TableAssignedTo(TBLName)
                If AssignedToWho = "" Then
                  Load_What = "Drivers"
                  Find_What = "Select a Server"
                  RequestListForm.Show 1
                  ReleaseRegisterLock
                  If GetSelected = "" Then Exit Sub
                  CurrentServer = GetSelected
                Else
                  CurrentServer = AssignedToWho
                End If
             End If
             dosomething ExitSub, CurrentServer
             If ExitSub = True Then Exit Sub
             TimersOn
            
            LoadSections False
            ReleaseRegisterLock
         Case Is = ChkUsage(OkCheck)
            If Number_In_Party = 0 Then Exit Sub
            Table_Number = MyTable(Index).Caption
            dosomething ExitSub, CurrentServer
            TimersOn
            If LeaveScreen Then
              LeaveScreen = False
              Unload Me
              Exit Sub
            End If
            If ExitSub = True Then Exit Sub
            LoadSections False
            Exit Sub
         Case Is = ChkUsage(InuseCheck)
            ReleaseRegisterLock
            TimersOn
            Exit Sub
     End Select
Exit Sub
ERH:
Err.Clear
End Sub
Private Sub EraseJoinedTable(tableN As String)
Dim Counter As Integer
For Counter = 0 To UBound(AllJoined, 1)
 If AllJoined(Counter).MainTable = tableN Then
    AllJoined(Counter).MainTable = ""
    AllJoined(Counter).JoinWith = ""
    Exit For
 End If
Next Counter
End Sub
Private Sub dosomething(ExitSub As Boolean, TableBelongsTo As String)
  ' Removed - POS runtime only
End Sub
Private Sub dosomething_REMOVED(ExitSub As Boolean, TableBelongsTo As String)
Sleep 0
Dim TBLName  As String
Dim Response As String
'Allow_Discount = 0 ': Timing = 0
If Number_In_Party = 0 Then Number_In_Party = 1
Type_Message = "": Order_Qty = 0
Selection_Number = 0: ModIdNo = 0
Dim HostName As String
Dim OkClose As Boolean
OkClose = False
Select Case Check_Status
        Case Is = MyOrderStatus(SplitCheck)
           If Val(Input_Option(40)) > 0 Then
            If CheckGratuity(Check_no) = False Then
              If Number_In_Party >= Val(Input_Option(40)) Then
                     If See_Msg("No Gratuity Has Been Applied", 4, "Apply Gratuity Now") = 6 Then
                        Check_Status = MyOrderStatus(EditCheck)
                        Exit Sub
                     End If
              End If
            End If
          End If
          If LoadCheck(Check_no, OkClose, TableBelongsTo, False, False, False) = False Then Exit Sub
          If Check_no <= 0 Then Exit Sub
          SplitCheckNo = Check_no
          ReviewCheck = False: Review_Check = False
          Split_Check.Show 1
          Check_Status = MyOrderStatus(EditCheck)
          Exit Sub
        Case Is = MyOrderStatus(NewCheck)
                If Check_Option(92) = 1 Then
                   If GetMember = "" Then ExitSub = True: Exit Sub
                End If
                Check_no = Assign_Check_No(TableBelongsTo, True)
                If Check_no = 0 Then Exit Sub
                ReDim_Array
                OrderIs = GetOrderType(HereSale)
                TimersOFF
                'If OpenTableOnly Then
                '  HostName = LoggedOnAs(0).EMPName
                '  LoggedOnAs(0).EMPName = TableBelongsTo
                '  Order.LoadOrderScreen
                '  Order.Done.DoClick
                '  LoggedOnAs(0).EMPName = HostName
                '  Exit Sub
                'End If
                If OpenTableOnly Then
                  HostName = LoggedOnAs(0).EMPName
                  LoggedOnAs(0).EMPName = TableBelongsTo
                  Order.LoadOrderScreen "Table Screen dosomthing"
                  If Input_Option(45) = "No" Then
                   Order.Done.DoClick
                  Else
                   Order.Show 1
                  End If
                  '// allows hostess to take a drink order
                  'Order.Done.DoClick
                  LoggedOnAs(0).EMPName = HostName
                  Exit Sub
                End If
                '*****************
                If IsServerOn(LoggedOnAs(0).EMPName) = True Then Exit Sub
                If Check_Option(156) = 1 Then
                 If See_Msg("Attach Credit Card?", 4, "") = 6 Then
                   Dim TempTable As String
                   TempTable = Table_Number
                   Chk_Sales_data(5) = Val(GetConfigInfo("Credit Cards", "PreAuth Amount", "10.00", "FUNCTIONS.INI"))
                   Response = PreAuthPax(Chk_Sales_data(5))
                   If Response <> "APPROVED" Then
                     See_Msg "Transaction Failed", 0, Response
                     Exit Sub
                   End If
                   Table_Number = TempTable
                 End If
                End If
                Order.LoadOrderScreen "Table Screen dosomthing"
                Order.Show
                If Check_Option(38) = 0 Then
                   If Check_Option(39) = 0 Or Val(LoggedOnAs(0).EmpAccessLevel) < Access_Required("Managers Menu") Then
                      Me.Hide
                      Call MemoryCheck(Startup)
                      ExitSub = True
                      Exit Sub
                   End If
                Else
                    Do While Order.Visible = True
                      DoEvents
                      Sleep 100
                    Loop
                    Me.ZOrder 0
                    Order.UnLoadOrderScreen
                    UnlockedCheckNumber Check_no
                    Timer1_Timer
                    TimersOn
                    Sleep 0
                End If
                
        Case Is = MyOrderStatus(EditCheck)
                If Check_no <= 0 Then Exit Sub
                If LoadCheck(Check_no, OkClose, TableBelongsTo, False, False, False) = False Then Exit Sub
                If Check_no <= 0 Then Exit Sub
                MultiClose = False
                OrderIs = GetOrderType(HereSale)
                TimersOFF
                Order.LoadOrderScreen "Table Screen dosomthing"
                Order.Show
                Order.ZOrder 0
                Sleep 0
                ExitSub = True
                If Check_Option(38) = 0 Then
                    If Check_Option(39) = 0 Or Val(LoggedOnAs(0).EmpAccessLevel) < Access_Required("Managers Menu") Then
                       Call MemoryCheck(Startup)
                       TimersOFF
                       'TableLayout.Hide
                       Unload Me
                       Exit Sub
                    End If
                Else
                    Do
                      If IsWindowVisible(Order.hwnd) = False Then
                       Exit Do
                      End If
                      DoEvents
                    Loop
                    TableLayout.Visible = True
                    Order.Hide
                    Order.UnLoadOrderScreen
                    UnlockedCheckNumber Check_no
                    If Check_Option(38) = 1 Then
                      TableLayout.LoadSections True
                    End If
                    Timer1_Timer
                    'Me.Refresh
                    TimersOn
                    Sleep 0
                End If
       Case Is = MyOrderStatus(PrintCheck)
               If Check_Status = MyOrderStatus(PrintCheck) Then
                    If LoadCheck(Check_no, OkClose, TableBelongsTo, False, False, False) = False Then Exit Sub
                    If Check_no < 0 Then Exit Sub
                    Check_Status = MyOrderStatus(PresentCheck)
                    TimersOFF
                    Order.LoadOrderScreen "Table Screen dosomthing"
                    Order.Done_Click
                    Order.Hide
                    Order.UnLoadOrderScreen
                   ' UnlockedCheckNumber Check_no
                    If PresentAllSplit = True Then
                       PrintAllSplits ReturnTableNo(Val(Check_no)), Check_no
                       PresentAllSplit = False
                    End If
                    TimersOn
                Exit Sub
               End If
     Case Is = MyOrderStatus(PresentCheck)
                If LoadCheck(Check_no, OkClose, TableBelongsTo, False, False, False) = False Then Exit Sub
                MultiClose = False
                OrderIs = GetOrderType(HereSale)
                TableLayout.Hide
                Sleep 0
                Order.LoadOrderScreen "Table Screen dosomthing"
                Order.Show  '1
                Order.ZOrder 0
                TimersOFF
                TableLayout.Visible = True
                If Check_Option(38) = 0 Then
                  If Check_Option(39) = 0 Or Val(LoggedOnAs(0).EmpAccessLevel) < Access_Required("Managers Menu") Then
                    'TableLayout.Hide: ExitSub = True: Exit Sub
                    Unload Me
                    Exit Sub
                  End If
                Else
                 Do
                  If Check_Status = MyOrderStatus(EditCheck) Then
                  Exit Do
                  End If
                  DoEvents
                 Loop
                End If
                'Timer1.Enabled = True
                UnlockedCheckNumber Check_no
                TimersOn
                Call MemoryCheck(Startup)
     Case Is = MyOrderStatus(CloseCheck)
                'CheckIsOpen = True
                If LoadCheck(Check_no, OkClose, TableBelongsTo, False, False, False) = False Then Exit Sub
                If OkClose = False Then Exit Sub
                    If InStr(MyTable(WhichButton).Caption, vbCrLf) > 0 Then
                     TBLName = Left(MyTable(WhichButton).Caption, InStr(MyTable(WhichButton).Caption, vbCrLf) - 1)  '-1
                    Else
                     TBLName = MyTable(WhichButton).Caption
                    End If
                'If Check_Option(63) = 1 Then
                If Check_Status = MyOrderStatus(CloseCheck) And MyTable(WhichButton).BackColor = vbYellow Then
                    DeleteNoBalanceOpenCheck Check_no, TableBelongsTo
                    Close_Open_Check Check_no, TableBelongsTo
                    'ReleaseRegisterLock CStr(Check_no)
                    UnlockedCheckNumber Check_no
                    JoinedTableClose TBLName
                    EraseJoinedTable TBLName
                    'MyTable(WhichButton).ToolTipText = ""
                    Exit Sub
                End If
                'End If

                OrderIs = GetOrderType(HereSale)
                TimersOFF
                Order.LoadOrderScreen "Table Screen dosomthing"
                Check_Status = MyOrderStatus(CloseCheck)
                Order.Done_Click
                Order.Hide
                Order.UnLoadOrderScreen
                Sleep 0
                JoinedTableClose TBLName
                EraseJoinedTable TBLName
                'MyTable(WhichButton).ToolTipText = ""
                'ReleaseRegisterLock CStr(Check_no)
                UnlockedCheckNumber Check_no
                If OkClose = False Then Exit Sub
                Close_Check     ' Check Title, 0 for no void
                If Check_Option(39) = 1 Then
                 If Val(LoggedOnAs(0).EmpAccessLevel) >= Access_Required("Managers Menu") Then
                  
                  Else
                   If IsServerOn(LoggedOnAs(0).EMPName) = True Then Exit Sub
                   Order.LoadOrderScreen "Table Screen dosomthing"
                   Order.SetBar
                   Order.Show
                   Call MemoryCheck(Startup)
                   Exit Sub
                  End If
                End If
                If Check_Option(38) = 0 Then
                  
                Else
                  TableLayout.Visible = True
                  Timer1_Timer
                  TimersOn
                  Me.ZOrder 0
                End If
     Case Else
               Check_Status = MyOrderStatus(EditCheck)
               TimersOFF
               Startup.Show
               'TableLayout.Hide
               Unload Me
    End Select
End Sub
Private Sub EditTables(Index As Integer)
    If TableData.Visible = True Then
       If MyTable(Index).Picture = "" Then
        If bmpfile.ListIndex > -1 Then
         MyTable(Index).Picture = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & bmpfile.List(bmpfile.ListIndex))
         MyTable(Index).Shape = LoadPicture(SvrPath.LocalTerminalPath & "Picture Files\" & bmpfile.List(bmpfile.ListIndex))
         MyTable(Index).TagVariant = bmpfile.List(bmpfile.ListIndex)
        End If
       End If
        If Trim(TableName.Text) > "" Then
            MyTable(Index).Tag = CStr(Cover.Value)
            MyTable(Index).Caption = TableName.Text
        End If
      Exit Sub
    End If
    If TableData.Visible = False Then
       Cover.Value = MyTable(Index).Tag
       TableName.Text = MyTable(Index).Caption
       TableData.Visible = True
       TableData.Left = 1440 * WidScale1024
       TableData.Top = 0
       Exit Sub
    End If
End Sub
Private Function GetCovers() As Integer
    On Error Resume Next
    Dim GC As CallCovers
    Set GC = New CallCovers
      GetCovers = GC.ShowCovers
      'FixedCoverCount = GetCovers
    Set GC = Nothing
End Function
Private Sub MyTable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

If DoWork = "AlignLeft" Then
    If SyncButton = 0 Then
       SyncButton = MyTable(Index).Left
    Else
       MyTable(Index).Left = SyncButton
    End If
    Exit Sub
End If
If DoWork = "AlignTop" Then
    If SyncButton = 0 Then
       SyncButton = MyTable(Index).Top
    Else
       MyTable(Index).Top = SyncButton
    End If
    Exit Sub
End If

If DoWork = "SnapLeft" Then
    If SyncButton = 0 Then
       SyncButton = MyTable(Index).Left
       SetTop = MyTable(Index).Height + (MyTable(Index).Top + MyTable(Index).Height / 3)
    Else
       MyTable(Index).Left = SyncButton
       MyTable(Index).Top = SetTop
       SetTop = SetTop + (MyTable(Index).Height + (MyTable(Index).Height / 3))
    End If
    Exit Sub
End If
If DoWork = "SnapTop" Then
    If SyncButton = 0 Then
       SyncButton = MyTable(Index).Top
       SetLeft = MyTable(Index).Left + (MyTable(Index).width + (MyTable(Index).width / 2))
    Else
       MyTable(Index).Top = SyncButton
       MyTable(Index).Left = SetLeft
       SetLeft = SetLeft + (MyTable(Index).width + (MyTable(Index).width / 2))
    End If
    Exit Sub
End If

    If EditLayout = True Then
        If Button = 1 Then
            MyTable(Index).Drag vbBeginDrag
            MoveTable = Index
            TableSelect = Index
        End If
    End If
End Sub
Private Sub MyTable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Error Resume Next
  If EditLayout = True Then
   If Button = 2 Then
      MyTable(Index).width = x
      MyTable(Index).Height = Y
   End If
 End If
 LT = x: RT = Y
End Sub

Private Sub TableData_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If EditLayout = True Then
    If Button = 1 Then
        TableData.Drag vbBeginDrag
        MoveTable = -1
    End If
   End If
End Sub
Private Sub TableData_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
LT = x: RT = Y
End Sub

Private Sub TableName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
If KeyAscii = 32 Then KeyAscii = 0
End Sub
Private Sub TBLlist_Click()
TblList.RemoveItem TblList.ListIndex
Sleep 0
End Sub
Private Sub TData_Click()
  ' Removed - POS runtime only
End Sub
Private Sub LoadScreen()
  ' Removed - POS runtime only
End Sub
Private Sub Timer1_Timer()
  ' Removed - POS runtime only
End Sub

' ─── Web Editor Helpers ──────────────────────────────────────────────────────

' Sets BackColor of a table button based on shape ID stored by the web editor
Public Sub SetShapeAppearance(ctl As Object, shapeName As String)
  Dim s As String
  s = LCase(shapeName)
  If InStr(s, "round") > 0 Then
    ctl.BackColor = &H5078A0
  ElseIf InStr(s, "bar") > 0 Or InStr(s, "stool") > 0 Or InStr(s, "highbar") > 0 Then
    ctl.BackColor = &H6A7A5A
  ElseIf InStr(s, "booth") > 0 Then
    ctl.BackColor = &H9A6B5C
  Else
    ctl.BackColor = &H4A6B8C   ' square / rect
  End If
  ctl.ForeColor = vbWhite
End Sub

' Strips "TBL-"/"FLR-" prefix and image extension so web shape IDs are stored cleanly
Public Function NormalizePicture(pic As String) As String
  Dim s As String
  s = pic
  If Left(LCase(s), 4) = "tbl-" Then s = Mid(s, 5)
  If Left(LCase(s), 4) = "flr-" Then s = Mid(s, 5)
  Dim ext As String
  For Each ext In Array(".bmp", ".jpg", ".jpeg", ".gif", ".png")
    If LCase(Right(s, Len(ext))) = ext Then
      s = Left(s, Len(s) - Len(ext))
      Exit For
    End If
  Next ext
  NormalizePicture = s
End Function

' Sets form background colour from a style ID (e.g. "wood", "tile")
Public Sub SetBgColor(styleID As String)
  Select Case LCase(Trim(styleID))
    Case "wood"     : Room.BackColor = &H2060A0
    Case "hardwood" : Room.BackColor = &H1A4A80
    Case "tile"     : Room.BackColor = &H808080
    Case "carpet"   : Room.BackColor = &H4A5A30
    Case "concrete" : Room.BackColor = &H606060
    Case "dark"     : Room.BackColor = &H202020
    Case Else       : Room.BackColor = &H404040
  End Select
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 