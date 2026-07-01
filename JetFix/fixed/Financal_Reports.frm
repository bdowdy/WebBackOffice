VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Financal_Reports 
   Appearance      =   0  'Flat
   BackColor       =   &H001F1B18&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   10635
   ClientLeft      =   75
   ClientTop       =   525
   ClientWidth     =   15270
   FillColor       =   &H001F1B18&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H001F1B18&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   10635
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ShowCust 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   9855
      Left            =   18120
      ScaleHeight     =   9825
      ScaleWidth      =   11145
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   11175
      Begin BTNENHLib4.BtnEnh BtnEnh30 
         Height          =   735
         Left            =   1800
         TabIndex        =   53
         Top             =   8880
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Address Search"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":0000
         textLT          =   "Financal_Reports.frx":007C
         textCT          =   "Financal_Reports.frx":0094
         textRT          =   "Financal_Reports.frx":00AC
         textLM          =   "Financal_Reports.frx":00C4
         textRM          =   "Financal_Reports.frx":00DC
         textLB          =   "Financal_Reports.frx":00F4
         textCB          =   "Financal_Reports.frx":010C
         textRB          =   "Financal_Reports.frx":0124
         colorBack       =   "Financal_Reports.frx":013C
         colorIntern     =   "Financal_Reports.frx":0166
         colorMO         =   "Financal_Reports.frx":0190
         colorFocus      =   "Financal_Reports.frx":01BA
         colorDisabled   =   "Financal_Reports.frx":01E4
         colorPressed    =   "Financal_Reports.frx":020E
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
      Begin BTNENHLib4.BtnEnh BtnEnh29 
         Height          =   735
         Left            =   3480
         TabIndex        =   52
         Top             =   8880
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Name Search"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":0238
         textLT          =   "Financal_Reports.frx":02AE
         textCT          =   "Financal_Reports.frx":02C6
         textRT          =   "Financal_Reports.frx":02DE
         textLM          =   "Financal_Reports.frx":02F6
         textRM          =   "Financal_Reports.frx":030E
         textLB          =   "Financal_Reports.frx":0326
         textCB          =   "Financal_Reports.frx":033E
         textRB          =   "Financal_Reports.frx":0356
         colorBack       =   "Financal_Reports.frx":036E
         colorIntern     =   "Financal_Reports.frx":0398
         colorMO         =   "Financal_Reports.frx":03C2
         colorFocus      =   "Financal_Reports.frx":03EC
         colorDisabled   =   "Financal_Reports.frx":0416
         colorPressed    =   "Financal_Reports.frx":0440
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
      Begin VSFlex8DAOCtl.VSFlexGrid CustList 
         Height          =   8535
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   10935
         _cx             =   1998211928
         _cy             =   1998207695
         Appearance      =   0
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14079702
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
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
      Begin BTNENHLib4.BtnEnh BtnEnh19 
         Height          =   780
         Left            =   9000
         TabIndex        =   42
         Top             =   8880
         Width           =   2055
         _Version        =   589828
         _ExtentX        =   3625
         _ExtentY        =   1376
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":046A
         textLT          =   "Financal_Reports.frx":04D2
         textCT          =   "Financal_Reports.frx":04EA
         textRT          =   "Financal_Reports.frx":0502
         textLM          =   "Financal_Reports.frx":051A
         textRM          =   "Financal_Reports.frx":0532
         textLB          =   "Financal_Reports.frx":054A
         textCB          =   "Financal_Reports.frx":0562
         textRB          =   "Financal_Reports.frx":057A
         colorBack       =   "Financal_Reports.frx":0592
         colorIntern     =   "Financal_Reports.frx":05BC
         colorMO         =   "Financal_Reports.frx":05E6
         colorFocus      =   "Financal_Reports.frx":0610
         colorDisabled   =   "Financal_Reports.frx":063A
         colorPressed    =   "Financal_Reports.frx":0664
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
      Begin BTNENHLib4.BtnEnh BtnEnh20 
         Height          =   780
         Left            =   6840
         TabIndex        =   43
         Top             =   8880
         Width           =   2055
         _Version        =   589828
         _ExtentX        =   3625
         _ExtentY        =   1376
         _StockProps     =   66
         Caption         =   "Print List"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":068E
         textLT          =   "Financal_Reports.frx":0702
         textCT          =   "Financal_Reports.frx":071A
         textRT          =   "Financal_Reports.frx":0732
         textLM          =   "Financal_Reports.frx":074A
         textRM          =   "Financal_Reports.frx":0762
         textLB          =   "Financal_Reports.frx":077A
         textCB          =   "Financal_Reports.frx":0792
         textRB          =   "Financal_Reports.frx":07AA
         colorBack       =   "Financal_Reports.frx":07C2
         colorIntern     =   "Financal_Reports.frx":07EC
         colorMO         =   "Financal_Reports.frx":0816
         colorFocus      =   "Financal_Reports.frx":0840
         colorDisabled   =   "Financal_Reports.frx":086A
         colorPressed    =   "Financal_Reports.frx":0894
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
      Begin BTNENHLib4.BtnEnh BtnEnh28 
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   8880
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Phone Search"
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":08BE
         textLT          =   "Financal_Reports.frx":0936
         textCT          =   "Financal_Reports.frx":094E
         textRT          =   "Financal_Reports.frx":0966
         textLM          =   "Financal_Reports.frx":097E
         textRM          =   "Financal_Reports.frx":0996
         textLB          =   "Financal_Reports.frx":09AE
         textCB          =   "Financal_Reports.frx":09C6
         textRB          =   "Financal_Reports.frx":09DE
         colorBack       =   "Financal_Reports.frx":09F6
         colorIntern     =   "Financal_Reports.frx":0A20
         colorMO         =   "Financal_Reports.frx":0A4A
         colorFocus      =   "Financal_Reports.frx":0A74
         colorDisabled   =   "Financal_Reports.frx":0A9E
         colorPressed    =   "Financal_Reports.frx":0AC8
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
   End
   Begin VB.PictureBox CustReportMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   6240
      ScaleHeight     =   5505
      ScaleWidth      =   6945
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   6975
      Begin VB.ListBox PendingList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3870
         Left            =   -4800
         TabIndex        =   45
         Top             =   3000
         Visible         =   0   'False
         Width           =   5535
      End
      Begin BTNENHLib4.BtnEnh BtnEnh6 
         Height          =   735
         Left            =   2040
         TabIndex        =   8
         Top             =   4560
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
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
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":0AF2
         textLT          =   "Financal_Reports.frx":0B5A
         textCT          =   "Financal_Reports.frx":0B72
         textRT          =   "Financal_Reports.frx":0B8A
         textLM          =   "Financal_Reports.frx":0BA2
         textRM          =   "Financal_Reports.frx":0BBA
         textLB          =   "Financal_Reports.frx":0BD2
         textCB          =   "Financal_Reports.frx":0BEA
         textRB          =   "Financal_Reports.frx":0C02
         colorBack       =   "Financal_Reports.frx":0C1A
         colorIntern     =   "Financal_Reports.frx":0C44
         colorMO         =   "Financal_Reports.frx":0C6E
         colorFocus      =   "Financal_Reports.frx":0C98
         colorDisabled   =   "Financal_Reports.frx":0CC2
         colorPressed    =   "Financal_Reports.frx":0CEC
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
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   855
         Left            =   450
         TabIndex        =   9
         Top             =   600
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Customer List"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":0D16
         textLT          =   "Financal_Reports.frx":0D90
         textCT          =   "Financal_Reports.frx":0DA8
         textRT          =   "Financal_Reports.frx":0DC0
         textLM          =   "Financal_Reports.frx":0DD8
         textRM          =   "Financal_Reports.frx":0DF0
         textLB          =   "Financal_Reports.frx":0E08
         textCB          =   "Financal_Reports.frx":0E20
         textRB          =   "Financal_Reports.frx":0E38
         colorBack       =   "Financal_Reports.frx":0E50
         colorIntern     =   "Financal_Reports.frx":0E7A
         colorMO         =   "Financal_Reports.frx":0EA4
         colorFocus      =   "Financal_Reports.frx":0ECE
         colorDisabled   =   "Financal_Reports.frx":0EF8
         colorPressed    =   "Financal_Reports.frx":0F22
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh8 
         Height          =   855
         Left            =   450
         TabIndex        =   10
         Top             =   2280
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Edit Customer"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":0F4C
         textLT          =   "Financal_Reports.frx":0FC6
         textCT          =   "Financal_Reports.frx":0FDE
         textRT          =   "Financal_Reports.frx":0FF6
         textLM          =   "Financal_Reports.frx":100E
         textRM          =   "Financal_Reports.frx":1026
         textLB          =   "Financal_Reports.frx":103E
         textCB          =   "Financal_Reports.frx":1056
         textRB          =   "Financal_Reports.frx":106E
         colorBack       =   "Financal_Reports.frx":1086
         colorIntern     =   "Financal_Reports.frx":10B0
         colorMO         =   "Financal_Reports.frx":10DA
         colorFocus      =   "Financal_Reports.frx":1104
         colorDisabled   =   "Financal_Reports.frx":112E
         colorPressed    =   "Financal_Reports.frx":1158
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh9 
         Height          =   855
         Left            =   3450
         TabIndex        =   11
         Top             =   2280
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Frequency Report"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":1182
         textLT          =   "Financal_Reports.frx":1202
         textCT          =   "Financal_Reports.frx":121A
         textRT          =   "Financal_Reports.frx":1232
         textLM          =   "Financal_Reports.frx":124A
         textRM          =   "Financal_Reports.frx":1262
         textLB          =   "Financal_Reports.frx":127A
         textCB          =   "Financal_Reports.frx":1292
         textRB          =   "Financal_Reports.frx":12AA
         colorBack       =   "Financal_Reports.frx":12C2
         colorIntern     =   "Financal_Reports.frx":12EC
         colorMO         =   "Financal_Reports.frx":1316
         colorFocus      =   "Financal_Reports.frx":1340
         colorDisabled   =   "Financal_Reports.frx":136A
         colorPressed    =   "Financal_Reports.frx":1394
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh11 
         Height          =   855
         Left            =   3450
         TabIndex        =   12
         Top             =   600
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Pending Orders"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":13BE
         textLT          =   "Financal_Reports.frx":143A
         textCT          =   "Financal_Reports.frx":1452
         textRT          =   "Financal_Reports.frx":146A
         textLM          =   "Financal_Reports.frx":1482
         textRM          =   "Financal_Reports.frx":149A
         textLB          =   "Financal_Reports.frx":14B2
         textCB          =   "Financal_Reports.frx":14CA
         textRB          =   "Financal_Reports.frx":14E2
         colorBack       =   "Financal_Reports.frx":14FA
         colorIntern     =   "Financal_Reports.frx":1524
         colorMO         =   "Financal_Reports.frx":154E
         colorFocus      =   "Financal_Reports.frx":1578
         colorDisabled   =   "Financal_Reports.frx":15A2
         colorPressed    =   "Financal_Reports.frx":15CC
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh15 
         Height          =   855
         Left            =   450
         TabIndex        =   37
         Top             =   3120
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Show All Customers"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":15F6
         textLT          =   "Financal_Reports.frx":167A
         textCT          =   "Financal_Reports.frx":1692
         textRT          =   "Financal_Reports.frx":16AA
         textLM          =   "Financal_Reports.frx":16C2
         textRM          =   "Financal_Reports.frx":16DA
         textLB          =   "Financal_Reports.frx":16F2
         textCB          =   "Financal_Reports.frx":170A
         textRB          =   "Financal_Reports.frx":1722
         colorBack       =   "Financal_Reports.frx":173A
         colorIntern     =   "Financal_Reports.frx":1764
         colorMO         =   "Financal_Reports.frx":178E
         colorFocus      =   "Financal_Reports.frx":17B8
         colorDisabled   =   "Financal_Reports.frx":17E2
         colorPressed    =   "Financal_Reports.frx":180C
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh18 
         Height          =   855
         Left            =   450
         TabIndex        =   38
         Top             =   1440
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Add Customer"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":1836
         textLT          =   "Financal_Reports.frx":18AE
         textCT          =   "Financal_Reports.frx":18C6
         textRT          =   "Financal_Reports.frx":18DE
         textLM          =   "Financal_Reports.frx":18F6
         textRM          =   "Financal_Reports.frx":190E
         textLB          =   "Financal_Reports.frx":1926
         textCB          =   "Financal_Reports.frx":193E
         textRB          =   "Financal_Reports.frx":1956
         colorBack       =   "Financal_Reports.frx":196E
         colorIntern     =   "Financal_Reports.frx":1998
         colorMO         =   "Financal_Reports.frx":19C2
         colorFocus      =   "Financal_Reports.frx":19EC
         colorDisabled   =   "Financal_Reports.frx":1A16
         colorPressed    =   "Financal_Reports.frx":1A40
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh21 
         Height          =   855
         Left            =   3450
         TabIndex        =   44
         Top             =   1440
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Pending Order Summary"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":1A6A
         textLT          =   "Financal_Reports.frx":1AF4
         textCT          =   "Financal_Reports.frx":1B0C
         textRT          =   "Financal_Reports.frx":1B24
         textLM          =   "Financal_Reports.frx":1B3C
         textRM          =   "Financal_Reports.frx":1B54
         textLB          =   "Financal_Reports.frx":1B6C
         textCB          =   "Financal_Reports.frx":1B84
         textRB          =   "Financal_Reports.frx":1B9C
         colorBack       =   "Financal_Reports.frx":1BB4
         colorIntern     =   "Financal_Reports.frx":1BDE
         colorMO         =   "Financal_Reports.frx":1C08
         colorFocus      =   "Financal_Reports.frx":1C32
         colorDisabled   =   "Financal_Reports.frx":1C5C
         colorPressed    =   "Financal_Reports.frx":1C86
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh16 
         Height          =   855
         Left            =   3450
         TabIndex        =   92
         Top             =   3120
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Email Customers"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   8421504
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":1CB0
         textLT          =   "Financal_Reports.frx":1D2E
         textCT          =   "Financal_Reports.frx":1D46
         textRT          =   "Financal_Reports.frx":1D5E
         textLM          =   "Financal_Reports.frx":1D76
         textRM          =   "Financal_Reports.frx":1D8E
         textLB          =   "Financal_Reports.frx":1DA6
         textCB          =   "Financal_Reports.frx":1DBE
         textRB          =   "Financal_Reports.frx":1DD6
         colorBack       =   "Financal_Reports.frx":1DEE
         colorIntern     =   "Financal_Reports.frx":1E18
         colorMO         =   "Financal_Reports.frx":1E42
         colorFocus      =   "Financal_Reports.frx":1E6C
         colorDisabled   =   "Financal_Reports.frx":1E96
         colorPressed    =   "Financal_Reports.frx":1EC0
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   3855
         Left            =   240
         Top             =   360
         Width           =   6495
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   4095
         Left            =   120
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.PictureBox Report_Type 
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
      Height          =   4455
      Left            =   13920
      ScaleHeight     =   4425
      ScaleWidth      =   4350
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   4380
      Begin VB.ComboBox TaxedSales 
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
         Height          =   390
         ItemData        =   "Financal_Reports.frx":1EEA
         Left            =   600
         List            =   "Financal_Reports.frx":1EF4
         TabIndex        =   39
         Text            =   "All Sales"
         Top             =   2880
         Width           =   3255
      End
      Begin BTNENHLib4.BtnEnh SSCommand9 
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   3855
         _Version        =   589828
         _ExtentX        =   6800
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Week to Date Report"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   12632256
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":1F13
         textLT          =   "Financal_Reports.frx":1F99
         textCT          =   "Financal_Reports.frx":1FB1
         textRT          =   "Financal_Reports.frx":1FC9
         textLM          =   "Financal_Reports.frx":1FE1
         textRM          =   "Financal_Reports.frx":1FF9
         textLB          =   "Financal_Reports.frx":2011
         textCB          =   "Financal_Reports.frx":2029
         textRB          =   "Financal_Reports.frx":2041
         colorBack       =   "Financal_Reports.frx":2059
         colorIntern     =   "Financal_Reports.frx":2083
         colorMO         =   "Financal_Reports.frx":20AD
         colorFocus      =   "Financal_Reports.frx":20D7
         colorDisabled   =   "Financal_Reports.frx":2101
         colorPressed    =   "Financal_Reports.frx":212B
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8421504
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   5395026
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand10 
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3855
         _Version        =   589828
         _ExtentX        =   6800
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Daily Report"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   12632256
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":2155
         textLT          =   "Financal_Reports.frx":21CD
         textCT          =   "Financal_Reports.frx":21E5
         textRT          =   "Financal_Reports.frx":21FD
         textLM          =   "Financal_Reports.frx":2215
         textRM          =   "Financal_Reports.frx":222D
         textLB          =   "Financal_Reports.frx":2245
         textCB          =   "Financal_Reports.frx":225D
         textRB          =   "Financal_Reports.frx":2275
         colorBack       =   "Financal_Reports.frx":228D
         colorIntern     =   "Financal_Reports.frx":22B7
         colorMO         =   "Financal_Reports.frx":22E1
         colorFocus      =   "Financal_Reports.frx":230B
         colorDisabled   =   "Financal_Reports.frx":2335
         colorPressed    =   "Financal_Reports.frx":235F
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8421504
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   5395026
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand7 
         Height          =   735
         Left            =   1080
         TabIndex        =   6
         Top             =   3480
         Width           =   2295
         _Version        =   589828
         _ExtentX        =   4048
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
         textCaption     =   "Financal_Reports.frx":2389
         textLT          =   "Financal_Reports.frx":23F1
         textCT          =   "Financal_Reports.frx":2409
         textRT          =   "Financal_Reports.frx":2421
         textLM          =   "Financal_Reports.frx":2439
         textRM          =   "Financal_Reports.frx":2451
         textLB          =   "Financal_Reports.frx":2469
         textCB          =   "Financal_Reports.frx":2481
         textRB          =   "Financal_Reports.frx":2499
         colorBack       =   "Financal_Reports.frx":24B1
         colorIntern     =   "Financal_Reports.frx":24DB
         colorMO         =   "Financal_Reports.frx":2505
         colorFocus      =   "Financal_Reports.frx":252F
         colorDisabled   =   "Financal_Reports.frx":2559
         colorPressed    =   "Financal_Reports.frx":2583
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
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2535
         Left            =   120
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox CustomerEdit 
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
      Height          =   6375
      Left            =   14040
      ScaleHeight     =   6345
      ScaleWidth      =   8385
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox CData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   2040
         TabIndex        =   48
         Top             =   4200
         Width           =   6135
      End
      Begin BTNENHLib4.BtnEnh BtnEnh14 
         Height          =   615
         Left            =   6240
         TabIndex        =   36
         Top             =   3240
         Width           =   1935
         _Version        =   589828
         _ExtentX        =   3413
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Show Last Order"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":25AD
         textLT          =   "Financal_Reports.frx":262B
         textCT          =   "Financal_Reports.frx":2643
         textRT          =   "Financal_Reports.frx":265B
         textLM          =   "Financal_Reports.frx":2673
         textRM          =   "Financal_Reports.frx":268B
         textLB          =   "Financal_Reports.frx":26A3
         textCB          =   "Financal_Reports.frx":26BB
         textRB          =   "Financal_Reports.frx":26D3
         colorBack       =   "Financal_Reports.frx":26EB
         colorIntern     =   "Financal_Reports.frx":2715
         colorMO         =   "Financal_Reports.frx":273F
         colorFocus      =   "Financal_Reports.frx":2769
         colorDisabled   =   "Financal_Reports.frx":2793
         colorPressed    =   "Financal_Reports.frx":27BD
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
         Office2003ColorNorth=   5395026
         Office2003ColorSouth=   2302755
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   5395026
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   16777215
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin VB.TextBox CData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   6240
         TabIndex        =   35
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox CData 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   3840
         TabIndex        =   33
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox CData 
         Appearance      =   0  'Flat
         Height          =   1245
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox CData 
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
         Height          =   390
         Index           =   5
         Left            =   6720
         TabIndex        =   28
         Top             =   1455
         Width           =   1455
      End
      Begin VB.TextBox CData 
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
         Height          =   405
         Index           =   4
         Left            =   4200
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox CData 
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
         Height          =   390
         Index           =   3
         Left            =   1080
         TabIndex        =   26
         Top             =   1455
         Width           =   2175
      End
      Begin VB.TextBox CData 
         Appearance      =   0  'Flat
         Height          =   405
         Index           =   2
         Left            =   1080
         TabIndex        =   25
         Top             =   960
         Width           =   7095
      End
      Begin VB.TextBox CData 
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
         Height          =   405
         Index           =   1
         Left            =   5280
         TabIndex        =   24
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox CData 
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
         Height          =   405
         Index           =   0
         Left            =   1080
         TabIndex        =   23
         Top             =   480
         Width           =   2775
      End
      Begin BTNENHLib4.BtnEnh BtnEnh10 
         Height          =   975
         Left            =   6720
         TabIndex        =   14
         Top             =   5160
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1720
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
         BackColorContainer=   16384
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":27E7
         textLT          =   "Financal_Reports.frx":284F
         textCT          =   "Financal_Reports.frx":2867
         textRT          =   "Financal_Reports.frx":287F
         textLM          =   "Financal_Reports.frx":2897
         textRM          =   "Financal_Reports.frx":28AF
         textLB          =   "Financal_Reports.frx":28C7
         textCB          =   "Financal_Reports.frx":28DF
         textRB          =   "Financal_Reports.frx":28F7
         colorBack       =   "Financal_Reports.frx":290F
         colorIntern     =   "Financal_Reports.frx":2939
         colorMO         =   "Financal_Reports.frx":2963
         colorFocus      =   "Financal_Reports.frx":298D
         colorDisabled   =   "Financal_Reports.frx":29B7
         colorPressed    =   "Financal_Reports.frx":29E1
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
      Begin BTNENHLib4.BtnEnh BtnEnh12 
         Height          =   975
         Left            =   1560
         TabIndex        =   30
         Top             =   5160
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Save"
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
         BackColorContainer=   4210752
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":2A0B
         textLT          =   "Financal_Reports.frx":2A73
         textCT          =   "Financal_Reports.frx":2A8B
         textRT          =   "Financal_Reports.frx":2AA3
         textLM          =   "Financal_Reports.frx":2ABB
         textRM          =   "Financal_Reports.frx":2AD3
         textLB          =   "Financal_Reports.frx":2AEB
         textCB          =   "Financal_Reports.frx":2B03
         textRB          =   "Financal_Reports.frx":2B1B
         colorBack       =   "Financal_Reports.frx":2B33
         colorIntern     =   "Financal_Reports.frx":2B5D
         colorMO         =   "Financal_Reports.frx":2B87
         colorFocus      =   "Financal_Reports.frx":2BB1
         colorDisabled   =   "Financal_Reports.frx":2BDB
         colorPressed    =   "Financal_Reports.frx":2C05
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
      Begin BTNENHLib4.BtnEnh BtnEnh13 
         Height          =   975
         Left            =   120
         TabIndex        =   31
         Top             =   5160
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Delete Customer"
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
         BackColorContainer=   4210752
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":2C2F
         textLT          =   "Financal_Reports.frx":2CAD
         textCT          =   "Financal_Reports.frx":2CC5
         textRT          =   "Financal_Reports.frx":2CDD
         textLM          =   "Financal_Reports.frx":2CF5
         textRM          =   "Financal_Reports.frx":2D0D
         textLB          =   "Financal_Reports.frx":2D25
         textCB          =   "Financal_Reports.frx":2D3D
         textRB          =   "Financal_Reports.frx":2D55
         colorBack       =   "Financal_Reports.frx":2D6D
         colorIntern     =   "Financal_Reports.frx":2D97
         colorMO         =   "Financal_Reports.frx":2DC1
         colorFocus      =   "Financal_Reports.frx":2DEB
         colorDisabled   =   "Financal_Reports.frx":2E15
         colorPressed    =   "Financal_Reports.frx":2E3F
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
         Office2003ColorNorth=   183
         Office2003ColorSouth=   128
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   183
         Office2003ColorSouthMouseOver=   183
         Office2003ColorBorderMouseOver=   1812708
         Office2003ColorNorthPressed=   183
         Office2003ColorSouthPressed=   128
         Office2003ColorBorderPressed=   1812708
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh23 
         Height          =   975
         Left            =   3120
         TabIndex        =   46
         Top             =   5160
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Attach Coupon"
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
         BackColorContainer=   4210752
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":2E69
         textLT          =   "Financal_Reports.frx":2EE3
         textCT          =   "Financal_Reports.frx":2EFB
         textRT          =   "Financal_Reports.frx":2F13
         textLM          =   "Financal_Reports.frx":2F2B
         textRM          =   "Financal_Reports.frx":2F43
         textLB          =   "Financal_Reports.frx":2F5B
         textCB          =   "Financal_Reports.frx":2F73
         textRB          =   "Financal_Reports.frx":2F8B
         colorBack       =   "Financal_Reports.frx":2FA3
         colorIntern     =   "Financal_Reports.frx":2FCD
         colorMO         =   "Financal_Reports.frx":2FF7
         colorFocus      =   "Financal_Reports.frx":3021
         colorDisabled   =   "Financal_Reports.frx":304B
         colorPressed    =   "Financal_Reports.frx":3075
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
         Office2003ColorNorth=   5395026
         Office2003ColorSouth=   2302755
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   5395026
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   16777215
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh24 
         Height          =   615
         Left            =   240
         TabIndex        =   47
         Top             =   1920
         Visible         =   0   'False
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Send Email to Customer"
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
         BackColorContainer=   14737632
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":309F
         textLT          =   "Financal_Reports.frx":312B
         textCT          =   "Financal_Reports.frx":3143
         textRT          =   "Financal_Reports.frx":315B
         textLM          =   "Financal_Reports.frx":3173
         textRM          =   "Financal_Reports.frx":318B
         textLB          =   "Financal_Reports.frx":31A3
         textCB          =   "Financal_Reports.frx":31BB
         textRB          =   "Financal_Reports.frx":31D3
         colorBack       =   "Financal_Reports.frx":31EB
         colorIntern     =   "Financal_Reports.frx":3215
         colorMO         =   "Financal_Reports.frx":323F
         colorFocus      =   "Financal_Reports.frx":3269
         colorDisabled   =   "Financal_Reports.frx":3293
         colorPressed    =   "Financal_Reports.frx":32BD
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
         Office2003ColorNorth=   5395026
         Office2003ColorSouth=   2302755
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   5395026
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   16777215
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Coupon Attached"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   49
         Top             =   4300
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Seen On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   6240
         TabIndex        =   34
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   32
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   21
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   20
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Information"
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
         Index           =   0
         Left            =   3000
         TabIndex        =   15
         Top             =   120
         Width           =   2535
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   4935
         Left            =   120
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.PictureBox PickList 
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
      Height          =   7215
      Left            =   -8040
      ScaleHeight     =   7185
      ScaleWidth      =   8985
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CheckBox PaymentbyHour 
         BackColor       =   &H00000000&
         Caption         =   "Payments By Hour"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   2880
         TabIndex        =   50
         Top             =   5730
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5145
         ItemData        =   "Financal_Reports.frx":32E7
         Left            =   240
         List            =   "Financal_Reports.frx":32E9
         TabIndex        =   2
         Top             =   480
         Width           =   8415
      End
      Begin BTNENHLib4.BtnEnh sscommand38 
         Height          =   855
         Left            =   3120
         TabIndex        =   105
         Top             =   6240
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
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
         textCaption     =   "Financal_Reports.frx":32EB
         textLT          =   "Financal_Reports.frx":3353
         textCT          =   "Financal_Reports.frx":336B
         textRT          =   "Financal_Reports.frx":3383
         textLM          =   "Financal_Reports.frx":339B
         textRM          =   "Financal_Reports.frx":33B3
         textLB          =   "Financal_Reports.frx":33CB
         textCB          =   "Financal_Reports.frx":33E3
         textRB          =   "Financal_Reports.frx":33FB
         colorBack       =   "Financal_Reports.frx":3413
         colorIntern     =   "Financal_Reports.frx":343D
         colorMO         =   "Financal_Reports.frx":3467
         colorFocus      =   "Financal_Reports.frx":3491
         colorDisabled   =   "Financal_Reports.frx":34BB
         colorPressed    =   "Financal_Reports.frx":34E5
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Make a Selection"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   6015
         Left            =   120
         Top             =   120
         Width           =   8775
      End
   End
   Begin VB.Frame CList 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Email Customer List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6975
      Left            =   -2160
      TabIndex        =   88
      Top             =   -5640
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CommandButton Command6 
         Caption         =   "Load Employee Email Address"
         Height          =   375
         Left            =   3600
         TabIndex        =   104
         Top             =   1560
         Width           =   4455
      End
      Begin VB.ComboBox AcctType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Financal_Reports.frx":350F
         Left            =   5640
         List            =   "Financal_Reports.frx":351F
         TabIndex        =   103
         Text            =   "House Account"
         Top             =   1080
         Width           =   2370
      End
      Begin VB.CommandButton Command5 
         Caption         =   "POS Accounts"
         Height          =   375
         Left            =   3600
         TabIndex        =   102
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Message 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   3600
         MultiLine       =   -1  'True
         TabIndex        =   101
         Top             =   3000
         Width           =   4455
      End
      Begin VB.TextBox Subject 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3600
         TabIndex        =   98
         Top             =   2280
         Width           =   4455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Mail File"
         Height          =   375
         Left            =   3600
         TabIndex        =   97
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6360
         TabIndex        =   96
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton SALL 
         Caption         =   "Select All"
         Height          =   375
         Left            =   240
         TabIndex        =   95
         Top             =   6480
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load By Date Range from phone orders"
         Height          =   375
         Left            =   3600
         TabIndex        =   93
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox FileToSend 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3600
         TabIndex        =   91
         Top             =   5280
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Select File to Send"
         Height          =   375
         Left            =   3600
         TabIndex        =   90
         Top             =   5760
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog Lpath 
         Left            =   7440
         Top             =   6480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ListBox MyList 
         Appearance      =   0  'Flat
         Height          =   5790
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   89
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Message To Customer"
         Height          =   255
         Left            =   3600
         TabIndex        =   100
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   255
         Left            =   3600
         TabIndex        =   99
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Only Loads Customers with Email Address"
         Height          =   255
         Left            =   3960
         TabIndex        =   94
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Caption         =   "Maintenance"
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
      Height          =   1695
      Left            =   720
      TabIndex        =   85
      Top             =   5520
      Width           =   5295
      Begin BTNENHLib4.BtnEnh SSCommand1 
         Height          =   1215
         Left            =   120
         TabIndex        =   86
         Top             =   360
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   2143
         _StockProps     =   66
         Caption         =   "End of Day Close"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   14120960
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":3559
         textLT          =   "Financal_Reports.frx":35D9
         textCT          =   "Financal_Reports.frx":35F1
         textRT          =   "Financal_Reports.frx":3609
         textLM          =   "Financal_Reports.frx":3621
         textRM          =   "Financal_Reports.frx":3639
         textLB          =   "Financal_Reports.frx":3651
         textCB          =   "Financal_Reports.frx":3669
         textRB          =   "Financal_Reports.frx":3681
         colorBack       =   "Financal_Reports.frx":3699
         colorIntern     =   "Financal_Reports.frx":36C3
         colorMO         =   "Financal_Reports.frx":36ED
         colorFocus      =   "Financal_Reports.frx":3717
         colorDisabled   =   "Financal_Reports.frx":3741
         colorPressed    =   "Financal_Reports.frx":376B
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   32768
         Office2003ColorSouth=   32768
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   32768
         Office2003ColorSouthMouseOver=   32768
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand5 
         Height          =   1215
         Left            =   2640
         TabIndex        =   87
         Top             =   360
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   2143
         _StockProps     =   66
         Caption         =   "End of Week Close"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   14120960
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":3795
         textLT          =   "Financal_Reports.frx":3817
         textCT          =   "Financal_Reports.frx":382F
         textRT          =   "Financal_Reports.frx":3847
         textLM          =   "Financal_Reports.frx":385F
         textRM          =   "Financal_Reports.frx":3877
         textLB          =   "Financal_Reports.frx":388F
         textCB          =   "Financal_Reports.frx":38A7
         textRB          =   "Financal_Reports.frx":38BF
         colorBack       =   "Financal_Reports.frx":38D7
         colorIntern     =   "Financal_Reports.frx":3901
         colorMO         =   "Financal_Reports.frx":392B
         colorFocus      =   "Financal_Reports.frx":3955
         colorDisabled   =   "Financal_Reports.frx":397F
         colorPressed    =   "Financal_Reports.frx":39A9
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   32768
         Office2003ColorSouth=   32768
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   32768
         Office2003ColorSouthMouseOver=   32768
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Keeping"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   82
      Top             =   7320
      Width           =   13455
      Begin BTNENHLib4.BtnEnh exit 
         Height          =   855
         Left            =   10680
         TabIndex        =   83
         Top             =   350
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
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
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":39D3
         textLT          =   "Financal_Reports.frx":3A3B
         textCT          =   "Financal_Reports.frx":3A53
         textRT          =   "Financal_Reports.frx":3A6B
         textLM          =   "Financal_Reports.frx":3A83
         textRM          =   "Financal_Reports.frx":3A9B
         textLB          =   "Financal_Reports.frx":3AB3
         textCB          =   "Financal_Reports.frx":3ACB
         textRB          =   "Financal_Reports.frx":3AE3
         colorBack       =   "Financal_Reports.frx":3AFB
         colorIntern     =   "Financal_Reports.frx":3B25
         colorMO         =   "Financal_Reports.frx":3B4F
         colorFocus      =   "Financal_Reports.frx":3B79
         colorDisabled   =   "Financal_Reports.frx":3BA3
         colorPressed    =   "Financal_Reports.frx":3BCD
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
      Begin BTNENHLib4.BtnEnh SSCommand3 
         Height          =   855
         Left            =   240
         TabIndex        =   84
         Top             =   360
         Width           =   5055
         _Version        =   589828
         _ExtentX        =   8916
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Time Keeping Reports"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":3BF7
         textLT          =   "Financal_Reports.frx":3C7F
         textCT          =   "Financal_Reports.frx":3C97
         textRT          =   "Financal_Reports.frx":3CAF
         textLM          =   "Financal_Reports.frx":3CC7
         textRM          =   "Financal_Reports.frx":3CDF
         textLB          =   "Financal_Reports.frx":3CF7
         textCB          =   "Financal_Reports.frx":3D0F
         textRB          =   "Financal_Reports.frx":3D27
         colorBack       =   "Financal_Reports.frx":3D3F
         colorIntern     =   "Financal_Reports.frx":3D69
         colorMO         =   "Financal_Reports.frx":3D93
         colorFocus      =   "Financal_Reports.frx":3DBD
         colorDisabled   =   "Financal_Reports.frx":3DE7
         colorPressed    =   "Financal_Reports.frx":3E11
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H001F1B18&
      Caption         =   "Sales History"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   6120
      TabIndex        =   61
      Top             =   1560
      Width           =   8055
      Begin BTNENHLib4.BtnEnh SSCommand34 
         Height          =   855
         Left            =   2760
         TabIndex        =   62
         Top             =   2160
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Register Group Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":3E3B
         textLT          =   "Financal_Reports.frx":3EC3
         textCT          =   "Financal_Reports.frx":3EDB
         textRT          =   "Financal_Reports.frx":3EF3
         textLM          =   "Financal_Reports.frx":3F0B
         textRM          =   "Financal_Reports.frx":3F23
         textLB          =   "Financal_Reports.frx":3F3B
         textCB          =   "Financal_Reports.frx":3F53
         textRB          =   "Financal_Reports.frx":3F6B
         colorBack       =   "Financal_Reports.frx":3F83
         colorIntern     =   "Financal_Reports.frx":3FAD
         colorMO         =   "Financal_Reports.frx":3FD7
         colorFocus      =   "Financal_Reports.frx":4001
         colorDisabled   =   "Financal_Reports.frx":402B
         colorPressed    =   "Financal_Reports.frx":4055
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   855
         Left            =   2760
         TabIndex        =   64
         Top             =   3000
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Register Group Meal Period Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":407F
         textLT          =   "Financal_Reports.frx":411F
         textCT          =   "Financal_Reports.frx":4137
         textRT          =   "Financal_Reports.frx":414F
         textLM          =   "Financal_Reports.frx":4167
         textRM          =   "Financal_Reports.frx":417F
         textLB          =   "Financal_Reports.frx":4197
         textCB          =   "Financal_Reports.frx":41AF
         textRB          =   "Financal_Reports.frx":41C7
         colorBack       =   "Financal_Reports.frx":41DF
         colorIntern     =   "Financal_Reports.frx":4209
         colorMO         =   "Financal_Reports.frx":4233
         colorFocus      =   "Financal_Reports.frx":425D
         colorDisabled   =   "Financal_Reports.frx":4287
         colorPressed    =   "Financal_Reports.frx":42B1
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand16 
         Height          =   855
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Sales Dollars"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":42DB
         textLT          =   "Financal_Reports.frx":4355
         textCT          =   "Financal_Reports.frx":436D
         textRT          =   "Financal_Reports.frx":4385
         textLM          =   "Financal_Reports.frx":439D
         textRM          =   "Financal_Reports.frx":43B5
         textLB          =   "Financal_Reports.frx":43CD
         textCB          =   "Financal_Reports.frx":43E5
         textRB          =   "Financal_Reports.frx":43FD
         colorBack       =   "Financal_Reports.frx":4415
         colorIntern     =   "Financal_Reports.frx":443F
         colorMO         =   "Financal_Reports.frx":4469
         colorFocus      =   "Financal_Reports.frx":4493
         colorDisabled   =   "Financal_Reports.frx":44BD
         colorPressed    =   "Financal_Reports.frx":44E7
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand27 
         Height          =   855
         Left            =   2760
         TabIndex        =   67
         Top             =   1320
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Sales Summary"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":4511
         textLT          =   "Financal_Reports.frx":458B
         textCT          =   "Financal_Reports.frx":45A3
         textRT          =   "Financal_Reports.frx":45BB
         textLM          =   "Financal_Reports.frx":45D3
         textRM          =   "Financal_Reports.frx":45EB
         textLB          =   "Financal_Reports.frx":4603
         textCB          =   "Financal_Reports.frx":461B
         textRB          =   "Financal_Reports.frx":4633
         colorBack       =   "Financal_Reports.frx":464B
         colorIntern     =   "Financal_Reports.frx":4675
         colorMO         =   "Financal_Reports.frx":469F
         colorFocus      =   "Financal_Reports.frx":46C9
         colorDisabled   =   "Financal_Reports.frx":46F3
         colorPressed    =   "Financal_Reports.frx":471D
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh26 
         Height          =   855
         Left            =   240
         TabIndex        =   68
         Top             =   1320
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Sales Tax"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":4747
         textLT          =   "Financal_Reports.frx":47B9
         textCT          =   "Financal_Reports.frx":47D1
         textRT          =   "Financal_Reports.frx":47E9
         textLM          =   "Financal_Reports.frx":4801
         textRM          =   "Financal_Reports.frx":4819
         textLB          =   "Financal_Reports.frx":4831
         textCB          =   "Financal_Reports.frx":4849
         textRB          =   "Financal_Reports.frx":4861
         colorBack       =   "Financal_Reports.frx":4879
         colorIntern     =   "Financal_Reports.frx":48A3
         colorMO         =   "Financal_Reports.frx":48CD
         colorFocus      =   "Financal_Reports.frx":48F7
         colorDisabled   =   "Financal_Reports.frx":4921
         colorPressed    =   "Financal_Reports.frx":494B
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand18 
         Height          =   855
         Left            =   240
         TabIndex        =   69
         Top             =   2160
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Category Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":4975
         textLT          =   "Financal_Reports.frx":49F1
         textCT          =   "Financal_Reports.frx":4A09
         textRT          =   "Financal_Reports.frx":4A21
         textLM          =   "Financal_Reports.frx":4A39
         textRM          =   "Financal_Reports.frx":4A51
         textLB          =   "Financal_Reports.frx":4A69
         textCB          =   "Financal_Reports.frx":4A81
         textRB          =   "Financal_Reports.frx":4A99
         colorBack       =   "Financal_Reports.frx":4AB1
         colorIntern     =   "Financal_Reports.frx":4ADB
         colorMO         =   "Financal_Reports.frx":4B05
         colorFocus      =   "Financal_Reports.frx":4B2F
         colorDisabled   =   "Financal_Reports.frx":4B59
         colorPressed    =   "Financal_Reports.frx":4B83
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand19 
         Height          =   855
         Left            =   240
         TabIndex        =   70
         Top             =   3000
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   $"Financal_Reports.frx":4BAD
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":4BC3
         textLT          =   "Financal_Reports.frx":4C47
         textCT          =   "Financal_Reports.frx":4C5F
         textRT          =   "Financal_Reports.frx":4C77
         textLM          =   "Financal_Reports.frx":4C8F
         textRM          =   "Financal_Reports.frx":4CA7
         textLB          =   "Financal_Reports.frx":4CBF
         textCB          =   "Financal_Reports.frx":4CD7
         textRB          =   "Financal_Reports.frx":4CEF
         colorBack       =   "Financal_Reports.frx":4D07
         colorIntern     =   "Financal_Reports.frx":4D31
         colorMO         =   "Financal_Reports.frx":4D5B
         colorFocus      =   "Financal_Reports.frx":4D85
         colorDisabled   =   "Financal_Reports.frx":4DAF
         colorPressed    =   "Financal_Reports.frx":4DD9
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   855
         Left            =   240
         TabIndex        =   71
         Top             =   3840
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Voided Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":4E03
         textLT          =   "Financal_Reports.frx":4E7B
         textCT          =   "Financal_Reports.frx":4E93
         textRT          =   "Financal_Reports.frx":4EAB
         textLM          =   "Financal_Reports.frx":4EC3
         textRM          =   "Financal_Reports.frx":4EDB
         textLB          =   "Financal_Reports.frx":4EF3
         textCB          =   "Financal_Reports.frx":4F0B
         textRB          =   "Financal_Reports.frx":4F23
         colorBack       =   "Financal_Reports.frx":4F3B
         colorIntern     =   "Financal_Reports.frx":4F65
         colorMO         =   "Financal_Reports.frx":4F8F
         colorFocus      =   "Financal_Reports.frx":4FB9
         colorDisabled   =   "Financal_Reports.frx":4FE3
         colorPressed    =   "Financal_Reports.frx":500D
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh34 
         Height          =   855
         Left            =   240
         TabIndex        =   72
         Top             =   4680
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Product Waste"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":5037
         textLT          =   "Financal_Reports.frx":50B1
         textCT          =   "Financal_Reports.frx":50C9
         textRT          =   "Financal_Reports.frx":50E1
         textLM          =   "Financal_Reports.frx":50F9
         textRM          =   "Financal_Reports.frx":5111
         textLB          =   "Financal_Reports.frx":5129
         textCB          =   "Financal_Reports.frx":5141
         textRB          =   "Financal_Reports.frx":5159
         colorBack       =   "Financal_Reports.frx":5171
         colorIntern     =   "Financal_Reports.frx":519B
         colorMO         =   "Financal_Reports.frx":51C5
         colorFocus      =   "Financal_Reports.frx":51EF
         colorDisabled   =   "Financal_Reports.frx":5219
         colorPressed    =   "Financal_Reports.frx":5243
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh27 
         Height          =   855
         Left            =   2760
         TabIndex        =   73
         Top             =   480
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Audit History"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":526D
         textLT          =   "Financal_Reports.frx":52E7
         textCT          =   "Financal_Reports.frx":52FF
         textRT          =   "Financal_Reports.frx":5317
         textLM          =   "Financal_Reports.frx":532F
         textRM          =   "Financal_Reports.frx":5347
         textLB          =   "Financal_Reports.frx":535F
         textCB          =   "Financal_Reports.frx":5377
         textRB          =   "Financal_Reports.frx":538F
         colorBack       =   "Financal_Reports.frx":53A7
         colorIntern     =   "Financal_Reports.frx":53D1
         colorMO         =   "Financal_Reports.frx":53FB
         colorFocus      =   "Financal_Reports.frx":5425
         colorDisabled   =   "Financal_Reports.frx":544F
         colorPressed    =   "Financal_Reports.frx":5479
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand42 
         Height          =   855
         Left            =   2760
         TabIndex        =   74
         Top             =   3840
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Meal Period Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":54A3
         textLT          =   "Financal_Reports.frx":5525
         textCT          =   "Financal_Reports.frx":553D
         textRT          =   "Financal_Reports.frx":5555
         textLM          =   "Financal_Reports.frx":556D
         textRM          =   "Financal_Reports.frx":5585
         textLB          =   "Financal_Reports.frx":559D
         textCB          =   "Financal_Reports.frx":55B5
         textRB          =   "Financal_Reports.frx":55CD
         colorBack       =   "Financal_Reports.frx":55E5
         colorIntern     =   "Financal_Reports.frx":560F
         colorMO         =   "Financal_Reports.frx":5639
         colorFocus      =   "Financal_Reports.frx":5663
         colorDisabled   =   "Financal_Reports.frx":568D
         colorPressed    =   "Financal_Reports.frx":56B7
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   855
         Left            =   2760
         TabIndex        =   75
         Top             =   4680
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Discounts"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":56E1
         textLT          =   "Financal_Reports.frx":5753
         textCT          =   "Financal_Reports.frx":576B
         textRT          =   "Financal_Reports.frx":5783
         textLM          =   "Financal_Reports.frx":579B
         textRM          =   "Financal_Reports.frx":57B3
         textLB          =   "Financal_Reports.frx":57CB
         textCB          =   "Financal_Reports.frx":57E3
         textRB          =   "Financal_Reports.frx":57FB
         colorBack       =   "Financal_Reports.frx":5813
         colorIntern     =   "Financal_Reports.frx":583D
         colorMO         =   "Financal_Reports.frx":5867
         colorFocus      =   "Financal_Reports.frx":5891
         colorDisabled   =   "Financal_Reports.frx":58BB
         colorPressed    =   "Financal_Reports.frx":58E5
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   855
         Left            =   5280
         TabIndex        =   76
         Top             =   3840
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Customer Tracking"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":590F
         textLT          =   "Financal_Reports.frx":5991
         textCT          =   "Financal_Reports.frx":59A9
         textRT          =   "Financal_Reports.frx":59C1
         textLM          =   "Financal_Reports.frx":59D9
         textRM          =   "Financal_Reports.frx":59F1
         textLB          =   "Financal_Reports.frx":5A09
         textCB          =   "Financal_Reports.frx":5A21
         textRB          =   "Financal_Reports.frx":5A39
         colorBack       =   "Financal_Reports.frx":5A51
         colorIntern     =   "Financal_Reports.frx":5A7B
         colorMO         =   "Financal_Reports.frx":5AA5
         colorFocus      =   "Financal_Reports.frx":5ACF
         colorDisabled   =   "Financal_Reports.frx":5AF9
         colorPressed    =   "Financal_Reports.frx":5B23
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand14 
         Height          =   855
         Left            =   5280
         TabIndex        =   77
         Top             =   2160
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Order Destination"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":5B4D
         textLT          =   "Financal_Reports.frx":5BCF
         textCT          =   "Financal_Reports.frx":5BE7
         textRT          =   "Financal_Reports.frx":5BFF
         textLM          =   "Financal_Reports.frx":5C17
         textRM          =   "Financal_Reports.frx":5C2F
         textLB          =   "Financal_Reports.frx":5C47
         textCB          =   "Financal_Reports.frx":5C5F
         textRB          =   "Financal_Reports.frx":5C77
         colorBack       =   "Financal_Reports.frx":5C8F
         colorIntern     =   "Financal_Reports.frx":5CB9
         colorMO         =   "Financal_Reports.frx":5CE3
         colorFocus      =   "Financal_Reports.frx":5D0D
         colorDisabled   =   "Financal_Reports.frx":5D37
         colorPressed    =   "Financal_Reports.frx":5D61
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh22 
         Height          =   855
         Left            =   5280
         TabIndex        =   78
         Top             =   480
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
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
         Shape           =   1
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":5D8B
         textLT          =   "Financal_Reports.frx":5E03
         textCT          =   "Financal_Reports.frx":5E1B
         textRT          =   "Financal_Reports.frx":5E33
         textLM          =   "Financal_Reports.frx":5E4B
         textRM          =   "Financal_Reports.frx":5E63
         textLB          =   "Financal_Reports.frx":5E7B
         textCB          =   "Financal_Reports.frx":5E93
         textRB          =   "Financal_Reports.frx":5EAB
         colorBack       =   "Financal_Reports.frx":5EC3
         colorIntern     =   "Financal_Reports.frx":5EED
         colorMO         =   "Financal_Reports.frx":5F17
         colorFocus      =   "Financal_Reports.frx":5F41
         colorDisabled   =   "Financal_Reports.frx":5F6B
         colorPressed    =   "Financal_Reports.frx":5F95
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh31 
         Height          =   855
         Left            =   5280
         TabIndex        =   79
         Top             =   3000
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Sales By Terminal"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":5FBF
         textLT          =   "Financal_Reports.frx":6041
         textCT          =   "Financal_Reports.frx":6059
         textRT          =   "Financal_Reports.frx":6071
         textLM          =   "Financal_Reports.frx":6089
         textRM          =   "Financal_Reports.frx":60A1
         textLB          =   "Financal_Reports.frx":60B9
         textCB          =   "Financal_Reports.frx":60D1
         textRB          =   "Financal_Reports.frx":60E9
         colorBack       =   "Financal_Reports.frx":6101
         colorIntern     =   "Financal_Reports.frx":612B
         colorMO         =   "Financal_Reports.frx":6155
         colorFocus      =   "Financal_Reports.frx":617F
         colorDisabled   =   "Financal_Reports.frx":61A9
         colorPressed    =   "Financal_Reports.frx":61D3
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh32 
         Height          =   855
         Left            =   5280
         TabIndex        =   80
         Top             =   1320
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Hourly Sales By Department"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":61FD
         textLT          =   "Financal_Reports.frx":6291
         textCT          =   "Financal_Reports.frx":62A9
         textRT          =   "Financal_Reports.frx":62C1
         textLM          =   "Financal_Reports.frx":62D9
         textRM          =   "Financal_Reports.frx":62F1
         textLB          =   "Financal_Reports.frx":6309
         textCB          =   "Financal_Reports.frx":6321
         textRB          =   "Financal_Reports.frx":6339
         colorBack       =   "Financal_Reports.frx":6351
         colorIntern     =   "Financal_Reports.frx":637B
         colorMO         =   "Financal_Reports.frx":63A5
         colorFocus      =   "Financal_Reports.frx":63CF
         colorDisabled   =   "Financal_Reports.frx":63F9
         colorPressed    =   "Financal_Reports.frx":6423
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand2 
         Height          =   855
         Left            =   5280
         TabIndex        =   81
         Top             =   4680
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Menu and Menu Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":644D
         textLT          =   "Financal_Reports.frx":64D3
         textCT          =   "Financal_Reports.frx":64EB
         textRT          =   "Financal_Reports.frx":6503
         textLM          =   "Financal_Reports.frx":651B
         textRM          =   "Financal_Reports.frx":6533
         textLB          =   "Financal_Reports.frx":654B
         textCB          =   "Financal_Reports.frx":6563
         textRB          =   "Financal_Reports.frx":657B
         colorBack       =   "Financal_Reports.frx":6593
         colorIntern     =   "Financal_Reports.frx":65BD
         colorMO         =   "Financal_Reports.frx":65E7
         colorFocus      =   "Financal_Reports.frx":6611
         colorDisabled   =   "Financal_Reports.frx":663B
         colorPressed    =   "Financal_Reports.frx":6665
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorSouth=   6118749
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Live Daily and Week to Date Sales"
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
      Height          =   3855
      Left            =   720
      TabIndex        =   54
      Top             =   1560
      Width           =   5295
      Begin BTNENHLib4.BtnEnh SSCommand8 
         Height          =   855
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Sales Reports"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":668F
         textLT          =   "Financal_Reports.frx":6709
         textCT          =   "Financal_Reports.frx":6721
         textRT          =   "Financal_Reports.frx":6739
         textLM          =   "Financal_Reports.frx":6751
         textRM          =   "Financal_Reports.frx":6769
         textLB          =   "Financal_Reports.frx":6781
         textCB          =   "Financal_Reports.frx":6799
         textRB          =   "Financal_Reports.frx":67B1
         colorBack       =   "Financal_Reports.frx":67C9
         colorIntern     =   "Financal_Reports.frx":67F3
         colorMO         =   "Financal_Reports.frx":681D
         colorFocus      =   "Financal_Reports.frx":6847
         colorDisabled   =   "Financal_Reports.frx":6871
         colorPressed    =   "Financal_Reports.frx":689B
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand35 
         Height          =   855
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Today's  Sales Recap"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":68C5
         textLT          =   "Financal_Reports.frx":694D
         textCT          =   "Financal_Reports.frx":6965
         textRT          =   "Financal_Reports.frx":697D
         textLM          =   "Financal_Reports.frx":6995
         textRM          =   "Financal_Reports.frx":69AD
         textLB          =   "Financal_Reports.frx":69C5
         textCB          =   "Financal_Reports.frx":69DD
         textRB          =   "Financal_Reports.frx":69F5
         colorBack       =   "Financal_Reports.frx":6A0D
         colorIntern     =   "Financal_Reports.frx":6A37
         colorMO         =   "Financal_Reports.frx":6A61
         colorFocus      =   "Financal_Reports.frx":6A8B
         colorDisabled   =   "Financal_Reports.frx":6AB5
         colorPressed    =   "Financal_Reports.frx":6ADF
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand12 
         Height          =   855
         Left            =   2640
         TabIndex        =   57
         Top             =   360
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Current Week Voids"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":6B09
         textLT          =   "Financal_Reports.frx":6B8D
         textCT          =   "Financal_Reports.frx":6BA5
         textRT          =   "Financal_Reports.frx":6BBD
         textLM          =   "Financal_Reports.frx":6BD5
         textRM          =   "Financal_Reports.frx":6BED
         textLB          =   "Financal_Reports.frx":6C05
         textCB          =   "Financal_Reports.frx":6C1D
         textRB          =   "Financal_Reports.frx":6C35
         colorBack       =   "Financal_Reports.frx":6C4D
         colorIntern     =   "Financal_Reports.frx":6C77
         colorMO         =   "Financal_Reports.frx":6CA1
         colorFocus      =   "Financal_Reports.frx":6CCB
         colorDisabled   =   "Financal_Reports.frx":6CF5
         colorPressed    =   "Financal_Reports.frx":6D1F
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand11 
         Height          =   855
         Left            =   120
         TabIndex        =   58
         Top             =   2040
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Category and  Department Sales"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":6D49
         textLT          =   "Financal_Reports.frx":6DE5
         textCT          =   "Financal_Reports.frx":6DFD
         textRT          =   "Financal_Reports.frx":6E15
         textLM          =   "Financal_Reports.frx":6E2D
         textRM          =   "Financal_Reports.frx":6E45
         textLB          =   "Financal_Reports.frx":6E5D
         textCB          =   "Financal_Reports.frx":6E75
         textRB          =   "Financal_Reports.frx":6E8D
         colorBack       =   "Financal_Reports.frx":6EA5
         colorIntern     =   "Financal_Reports.frx":6ECF
         colorMO         =   "Financal_Reports.frx":6EF9
         colorFocus      =   "Financal_Reports.frx":6F23
         colorDisabled   =   "Financal_Reports.frx":6F4D
         colorPressed    =   "Financal_Reports.frx":6F77
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand6 
         Height          =   855
         Left            =   2640
         TabIndex        =   59
         Top             =   1200
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Current Week Discount Reports"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":6FA1
         textLT          =   "Financal_Reports.frx":703B
         textCT          =   "Financal_Reports.frx":7053
         textRT          =   "Financal_Reports.frx":706B
         textLM          =   "Financal_Reports.frx":7083
         textRM          =   "Financal_Reports.frx":709B
         textLB          =   "Financal_Reports.frx":70B3
         textCB          =   "Financal_Reports.frx":70CB
         textRB          =   "Financal_Reports.frx":70E3
         colorBack       =   "Financal_Reports.frx":70FB
         colorIntern     =   "Financal_Reports.frx":7125
         colorMO         =   "Financal_Reports.frx":714F
         colorFocus      =   "Financal_Reports.frx":7179
         colorDisabled   =   "Financal_Reports.frx":71A3
         colorPressed    =   "Financal_Reports.frx":71CD
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand37 
         Height          =   855
         Left            =   120
         TabIndex        =   60
         Top             =   2880
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   $"Financal_Reports.frx":71F7
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":720B
         textLT          =   "Financal_Reports.frx":728B
         textCT          =   "Financal_Reports.frx":72A3
         textRT          =   "Financal_Reports.frx":72BB
         textLM          =   "Financal_Reports.frx":72D3
         textRM          =   "Financal_Reports.frx":72EB
         textLB          =   "Financal_Reports.frx":7303
         textCB          =   "Financal_Reports.frx":731B
         textRB          =   "Financal_Reports.frx":7333
         colorBack       =   "Financal_Reports.frx":734B
         colorIntern     =   "Financal_Reports.frx":7375
         colorMO         =   "Financal_Reports.frx":739F
         colorFocus      =   "Financal_Reports.frx":73C9
         colorDisabled   =   "Financal_Reports.frx":73F3
         colorPressed    =   "Financal_Reports.frx":741D
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand33 
         Height          =   855
         Left            =   2640
         TabIndex        =   63
         Top             =   2040
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Revenue Centers"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":7447
         textLT          =   "Financal_Reports.frx":74C5
         textCT          =   "Financal_Reports.frx":74DD
         textRT          =   "Financal_Reports.frx":74F5
         textLM          =   "Financal_Reports.frx":750D
         textRM          =   "Financal_Reports.frx":7525
         textLB          =   "Financal_Reports.frx":753D
         textCB          =   "Financal_Reports.frx":7555
         textRB          =   "Financal_Reports.frx":756D
         colorBack       =   "Financal_Reports.frx":7585
         colorIntern     =   "Financal_Reports.frx":75AF
         colorMO         =   "Financal_Reports.frx":75D9
         colorFocus      =   "Financal_Reports.frx":7603
         colorDisabled   =   "Financal_Reports.frx":762D
         colorPressed    =   "Financal_Reports.frx":7657
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand25 
         Height          =   855
         Left            =   2640
         TabIndex        =   65
         Top             =   2880
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Daily Sales Summary"
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   4194304
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Financal_Reports.frx":7681
         textLT          =   "Financal_Reports.frx":7707
         textCT          =   "Financal_Reports.frx":771F
         textRT          =   "Financal_Reports.frx":7737
         textLM          =   "Financal_Reports.frx":774F
         textRM          =   "Financal_Reports.frx":7767
         textLB          =   "Financal_Reports.frx":777F
         textCB          =   "Financal_Reports.frx":7797
         textRB          =   "Financal_Reports.frx":77AF
         colorBack       =   "Financal_Reports.frx":77C7
         colorIntern     =   "Financal_Reports.frx":77F1
         colorMO         =   "Financal_Reports.frx":781B
         colorFocus      =   "Financal_Reports.frx":7845
         colorDisabled   =   "Financal_Reports.frx":786F
         colorPressed    =   "Financal_Reports.frx":7899
         HollowFrame     =   -1  'True
         LightDirection  =   1
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
         Office2003ColorNorth=   8404992
         Office2003ColorSouth=   8404992
         Office2003ColorBorder=   12632256
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
   End
End
Attribute VB_Name = "Financal_Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Report_Name As String
Private Type DayData
     DayofWeek As String
     NoOfWeeks  As Long
     BkSLS As Currency
     BkCUST As Long
     LuSLS As Currency
     LuCUST As Long
     ThSLS As Currency
     ThCUST As Long
     DnSLS As Currency
     DnCUST As Long
     LtSLS As Currency
     LtCUST As Long
     DaySale As Currency
     DayCount As Long
End Type
Private Type FinanicalData
     GSales   As Currency
     NSales   As Currency
     NotTaxed As Currency
     NSTax    As Currency
     NSPTax   As Currency
     NLqTax   As Currency
     NSRev    As Currency
     TxName   As String
     Taxed    As Currency
End Type
Private Enum SalesByHow
     ByHour
     ByDepartment
     ByItem
     ByCategory
     ByInd
End Enum
Dim DayInfo() As DayData
Dim PathforSendFile As String
Public Sub MealPeriodReport(ByVal StartDate As String, ByVal EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort As String
Dim Heading     As String
Dim NST         As Currency
Dim BS          As Currency
Dim BC          As Integer
Dim LpTh        As Integer
Erase DayInfo
ReDim DayInfo(7)
DayInfo(1).DayofWeek = "Monday"
DayInfo(2).DayofWeek = "Tuesday"
DayInfo(3).DayofWeek = "Wednesday"
DayInfo(4).DayofWeek = "Thursday"
DayInfo(5).DayofWeek = "Friday"
DayInfo(6).DayofWeek = "Saturday"
DayInfo(7).DayofWeek = "Sunday"
Dim LS      As Currency
Dim LC      As Long
Dim Ts      As Currency
Dim TC      As Long
Dim TCust   As Long
Dim dS      As Currency
Dim DC      As Long
Dim LTS     As Currency
Dim LTC     As Long
Dim BigT    As Currency
Dim Sp      As Long
Dim TTax    As Currency
Dim STAX    As Currency
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim STRS    As ADODB.Recordset
ConnectADODB MyRegisters(0).Svr & ClosingHistory, ClosingHistory, Db
    Sql = "Select * From [Meal Periods] "
    Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " order by TRY_CAST([date] AS datetime),[Name] = '" & Input_Option(71) & "',[Name] = '" & Input_Option(72) & "'"
    Sql = Sql & " ,[Name] = '" & Input_Option(73) & "',[Name] = '" & Input_Option(74) & "',[Name] = '" & Input_Option(75) & "' asc"
ConnectADORS (Sql), Db, RS
'//Input_Option(73)  Meal period names
'//Input_Option(74)  ex Breakfast - lunch - dinner etc.
'//Input_Option(75)
'//Input_Option(76)
If RS.EOF Then
   DbClose Db, RS
   See_Msg "No Sales Found", 5, ""
   Exit Sub
End If
    MealPeriodDept StartDate, EndDate
    SetRegular 7
    DrawLine
    Sp = 500
    PrintLine "Date", 1200
    PrintLine "Day", 2000
    PrintLine Input_Option(71), 2750: PrintLine "Count", 3500 '+ Sp
    PrintLine Input_Option(72), 4250: PrintLine "Count", 5000 '+ Sp
    PrintLine Input_Option(73), 5750: PrintLine "Count", 6500 '+ Sp
    PrintLine Input_Option(74), 7250: PrintLine "Count", 8000 '+ Sp
    PrintLine Input_Option(75), 8750: PrintLine "Count", 9500 '+ Sp
    PrintLine "Total", 10250
    PrintLine "Tax", -1
    DrawLine
    BS = 0: BC = 0: LS = 0: LC = 0: Ts = 0: TC = 0: dS = 0: DC = 0: LTS = 0: LTC = 0: NST = 0: BigT = 0: TTax = 0: TCust = 0
    
    Dim PeriodDate As String
    
    Do While Not RS.EOF
        NST = 0
        PrintLine Format(RS![Date], "MM/DD/YY"), 1200
        PrintLine Format(RS![Date], "dddd"), 2000
        PeriodDate = Format(RS![Date], "MM/DD/YYYY")
        If RS![Amount] <> 0 Then
        PrintLine Format(RS![Amount], "0.00"), 2750: PrintLine Format(RS![Quantity], "00"), 3500
        Else
         PrintLine "", 2750: PrintLine "", 3500
        End If
        NST = NST + RS![Amount]
        BS = RS![Amount]
        BC = RS![Quantity]
        TCust = TCust + BC
        
        If Not RS.EOF Then RS.MoveNext
        If RS.EOF Then GoTo StartHere
        If RS![Amount] <> 0 Then
        PrintLine Format(RS![Amount], "0.00"), 4250: PrintLine Format(RS![Quantity], "00"), 5000
        Else
        PrintLine "", 4250: PrintLine "", 5000
        End If
        NST = NST + RS![Amount]
        LS = RS![Amount]
        LC = RS![Quantity]
        TCust = TCust + LC
        If Not RS.EOF Then RS.MoveNext
        If RS.EOF Then GoTo StartHere
        If RS![Amount] <> 0 Then
         PrintLine Format(RS![Amount], "0.00"), 5750: PrintLine Format(RS![Quantity], "00"), 6500
        Else
         PrintLine "", 5750: PrintLine "", 6500
        End If
        NST = NST + RS![Amount]
        Ts = RS![Amount]
        TC = RS![Quantity]
        TCust = TCust + TC
        If Not RS.EOF Then RS.MoveNext
        If RS.EOF Then GoTo StartHere
        If RS![Amount] <> 0 Then
         PrintLine Format(RS![Amount], "0.00"), 7250: PrintLine Format(RS![Quantity], "00"), 8000
        Else
         PrintLine "", 7250: PrintLine "", 8000
        End If
        NST = NST + RS![Amount]
        dS = RS![Amount]
        DC = RS![Quantity]
        TCust = TCust + DC
        If Not RS.EOF Then RS.MoveNext
        If RS.EOF Then GoTo StartHere
        If RS![Amount] <> 0 Then
         PrintLine Format(RS![Amount], "0.00"), 8750: PrintLine Format(RS![Quantity], "00"), 9500
        Else
         PrintLine "", 8750: PrintLine "", 9500
        End If
        NST = NST + RS![Amount]
        LTS = RS![Amount]
        LTC = RS![Quantity]
        TCust = TCust + LTC
StartHere:
        PrintLine Format(NST, "0.00"), 10250
        Sql = "Select * from [Sales] Where [Date] = '" & PeriodDate & "'"
         ConnectADORS Sql, Db, STRS

          If Not STRS.EOF Then
            STAX = Val(STRS![Sales Tax] & "") + Val(STRS![Special Tax] & "") + Val(STRS![Liquor Tax] & "")
          Else
            STAX = 0
          End If
          SetNothing STRS
        PrintLine Format(STAX, "0.00"), -1
        TTax = TTax + STAX
        BigT = BigT + NST
         Select Case Format(PeriodDate, "dddd")
           Case Is = "Monday"
             DayInfo(1).BkCUST = DayInfo(1).BkCUST + BC
             DayInfo(1).BkSLS = DayInfo(1).BkSLS + BS
             DayInfo(1).LuCUST = DayInfo(1).LuCUST + LC
             DayInfo(1).LuSLS = DayInfo(1).LuSLS + LS
             DayInfo(1).ThCUST = DayInfo(1).ThCUST + TC
             DayInfo(1).ThSLS = DayInfo(1).ThSLS + Ts
             DayInfo(1).DnCUST = DayInfo(1).DnCUST + DC
             DayInfo(1).DnSLS = DayInfo(1).DnSLS + dS
             DayInfo(1).LtCUST = DayInfo(1).LtCUST + LTC
             DayInfo(1).LtSLS = DayInfo(1).LtSLS + LTS
             DayInfo(1).DaySale = DayInfo(1).DaySale + NST
             DayInfo(1).NoOfWeeks = DayInfo(1).NoOfWeeks + 1
           Case Is = "Tuesday"
             DayInfo(2).BkCUST = DayInfo(2).BkCUST + BC
             DayInfo(2).BkSLS = DayInfo(2).BkSLS + BS
             DayInfo(2).LuCUST = DayInfo(2).LuCUST + LC
             DayInfo(2).LuSLS = DayInfo(2).LuSLS + LS
             DayInfo(2).ThCUST = DayInfo(2).ThCUST + TC
             DayInfo(2).ThSLS = DayInfo(2).ThSLS + Ts
             DayInfo(2).DnCUST = DayInfo(2).DnCUST + DC
             DayInfo(2).DnSLS = DayInfo(2).DnSLS + dS
             DayInfo(2).LtCUST = DayInfo(2).LtCUST + LTC
             DayInfo(2).LtSLS = DayInfo(2).LtSLS + LTS
             DayInfo(2).DaySale = DayInfo(2).DaySale + NST
             DayInfo(2).NoOfWeeks = DayInfo(2).NoOfWeeks + 1
           Case Is = "Wednesday"
             DayInfo(3).NoOfWeeks = DayInfo(3).NoOfWeeks + 1
             DayInfo(3).BkCUST = DayInfo(3).BkCUST + BC
             DayInfo(3).BkSLS = DayInfo(3).BkSLS + BS
             DayInfo(3).LuCUST = DayInfo(3).LuCUST + LC
             DayInfo(3).LuSLS = DayInfo(3).LuSLS + LS
             DayInfo(3).ThCUST = DayInfo(3).ThCUST + TC
             DayInfo(3).ThSLS = DayInfo(3).ThSLS + Ts
             DayInfo(3).DnCUST = DayInfo(3).DnCUST + DC
             DayInfo(3).DnSLS = DayInfo(3).DnSLS + dS
             DayInfo(3).LtCUST = DayInfo(3).LtCUST + LTC
             DayInfo(3).LtSLS = DayInfo(3).LtSLS + LTS
             DayInfo(3).DaySale = DayInfo(3).DaySale + NST
           Case Is = "Thursday"
             DayInfo(4).NoOfWeeks = DayInfo(4).NoOfWeeks + 1
             DayInfo(4).BkCUST = DayInfo(4).BkCUST + BC
             DayInfo(4).BkSLS = DayInfo(4).BkSLS + BS
             DayInfo(4).LuCUST = DayInfo(4).LuCUST + LC
             DayInfo(4).LuSLS = DayInfo(4).LuSLS + LS
             DayInfo(4).ThCUST = DayInfo(4).ThCUST + TC
             DayInfo(4).ThSLS = DayInfo(4).ThSLS + Ts
             DayInfo(4).DnCUST = DayInfo(4).DnCUST + DC
             DayInfo(4).DnSLS = DayInfo(4).DnSLS + dS
             DayInfo(4).LtCUST = DayInfo(4).LtCUST + LTC
             DayInfo(4).LtSLS = DayInfo(4).LtSLS + LTS
             DayInfo(4).DaySale = DayInfo(4).DaySale + NST
           Case Is = "Friday"
             DayInfo(5).NoOfWeeks = DayInfo(5).NoOfWeeks + 1
             DayInfo(5).BkCUST = DayInfo(5).BkCUST + BC
             DayInfo(5).BkSLS = DayInfo(5).BkSLS + BS
             DayInfo(5).LuCUST = DayInfo(5).LuCUST + LC
             DayInfo(5).LuSLS = DayInfo(5).LuSLS + LS
             DayInfo(5).ThCUST = DayInfo(5).ThCUST + TC
             DayInfo(5).ThSLS = DayInfo(5).ThSLS + Ts
             DayInfo(5).DnCUST = DayInfo(5).DnCUST + DC
             DayInfo(5).DnSLS = DayInfo(5).DnSLS + dS
             DayInfo(5).LtCUST = DayInfo(5).LtCUST + LTC
             DayInfo(5).LtSLS = DayInfo(5).LtSLS + LTS
             DayInfo(5).DaySale = DayInfo(5).DaySale + NST
           Case Is = "Saturday"
             DayInfo(6).NoOfWeeks = DayInfo(6).NoOfWeeks + 1
             DayInfo(6).BkCUST = DayInfo(6).BkCUST + BC
             DayInfo(6).BkSLS = DayInfo(6).BkSLS + BS
             DayInfo(6).LuCUST = DayInfo(6).LuCUST + LC
             DayInfo(6).LuSLS = DayInfo(6).LuSLS + LS
             DayInfo(6).ThCUST = DayInfo(6).ThCUST + TC
             DayInfo(6).ThSLS = DayInfo(6).ThSLS + Ts
             DayInfo(6).DnCUST = DayInfo(6).DnCUST + DC
             DayInfo(6).DnSLS = DayInfo(6).DnSLS + dS
             DayInfo(6).LtCUST = DayInfo(6).LtCUST + LTC
             DayInfo(6).LtSLS = DayInfo(6).LtSLS + LTS
             DayInfo(6).DaySale = DayInfo(6).DaySale + NST
           Case Else
             DayInfo(7).NoOfWeeks = DayInfo(7).NoOfWeeks + 1
             DayInfo(7).BkCUST = DayInfo(7).BkCUST + BC
             DayInfo(7).BkSLS = DayInfo(7).BkSLS + BS
             DayInfo(7).LuCUST = DayInfo(7).LuCUST + LC
             DayInfo(7).LuSLS = DayInfo(7).LuSLS + LS
             DayInfo(7).ThCUST = DayInfo(7).ThCUST + TC
             DayInfo(7).ThSLS = DayInfo(7).ThSLS + Ts
             DayInfo(7).DnCUST = DayInfo(7).DnCUST + DC
             DayInfo(7).DnSLS = DayInfo(7).DnSLS + dS
             DayInfo(7).LtCUST = DayInfo(7).LtCUST + LTC
             DayInfo(7).LtSLS = DayInfo(7).LtSLS + LTS
             DayInfo(7).DaySale = DayInfo(7).DaySale + NST
           End Select
      If Not RS.EOF Then RS.MoveNext
    Loop
    DbClose Db, RS
    Dim SALTL As Currency
    Dim ctl   As Long
    DrawLine
    PrintLine "", 2000
    SALTL = 0: ctl = 0
    For LpTh = 1 To 7
     SALTL = SALTL + DayInfo(LpTh).BkSLS
     ctl = ctl + DayInfo(LpTh).BkCUST
    Next LpTh
    PrintLine Format(SALTL, "0.00"), 2750: PrintLine Format(ctl, "00"), 3500
    SALTL = 0: ctl = 0
    For LpTh = 1 To 7
       SALTL = SALTL + DayInfo(LpTh).LuSLS
       ctl = ctl + DayInfo(LpTh).LuCUST
    Next LpTh
    PrintLine Format(SALTL, "0.00"), 4250: PrintLine Format(ctl, "00"), 5000
    SALTL = 0: ctl = 0
    For LpTh = 1 To 7
       SALTL = SALTL + DayInfo(LpTh).ThSLS
       ctl = ctl + DayInfo(LpTh).ThCUST
    Next LpTh
    PrintLine Format(SALTL, "0.00"), 5750: PrintLine Format(ctl, "00"), 6500
    SALTL = 0: ctl = 0
    For LpTh = 1 To 7
       SALTL = SALTL + DayInfo(LpTh).DnSLS
       ctl = ctl + DayInfo(LpTh).DnCUST
    Next LpTh
    PrintLine Format(SALTL, "0.00"), 7250: PrintLine Format(ctl, "00"), 8000
    SALTL = 0: ctl = 0
    For LpTh = 1 To 7
       SALTL = SALTL + DayInfo(LpTh).LtSLS
       ctl = ctl + DayInfo(LpTh).LtCUST
    Next LpTh
    PrintLine Format(SALTL, "0.00"), 8750: PrintLine Format(ctl, "00"), 9500
    PrintLine Format(BigT, "0.00"), 10250
    PrintLine Format(TTax, "0.00"), -1
    DrawLine
    PrintCLine "Averages", -1, 0
    DrawLine
    PrintLine "Day", 2000
    PrintLine Input_Option(71), 2750: PrintLine "C-Avg", 3500
    PrintLine Input_Option(72), 4250: PrintLine "C-Avg", 5000
    PrintLine Input_Option(73), 5750: PrintLine "C-Avg", 6500
    PrintLine Input_Option(74), 7250: PrintLine "C-Avg", 8000
    PrintLine Input_Option(75), 8750: PrintLine "C-Avg", 9500
    PrintLine "Total", 10250: PrintLine "C-Avg", -1
    DrawLine
    For LpTh = 1 To 7
      If DayInfo(LpTh).DaySale > 0 Then
       PrintLine DayInfo(LpTh).DayofWeek, 2000
       If DayInfo(LpTh).NoOfWeeks > 0 And DayInfo(LpTh).BkSLS > 0 And DayInfo(LpTh).BkCUST > 0 Then
        PrintLine Format(DayInfo(LpTh).BkSLS / DayInfo(LpTh).NoOfWeeks, "0.00"), 2750: PrintLine Format(DayInfo(LpTh).BkSLS / DayInfo(LpTh).BkCUST, "0.00"), 3500
       Else
        PrintLine "0.00", 2750: PrintLine "0.00", 3500
       End If
       If DayInfo(LpTh).NoOfWeeks > 0 And DayInfo(LpTh).LuSLS > 0 And DayInfo(LpTh).LuCUST > 0 Then
        PrintLine Format(DayInfo(LpTh).LuSLS / DayInfo(LpTh).NoOfWeeks, "0.00"), 4250: PrintLine Format(DayInfo(LpTh).LuSLS / DayInfo(LpTh).LuCUST, "0.00"), 5000
       Else
        PrintLine "0.00", 4250: PrintLine "0.00", 5000
       End If
       If DayInfo(LpTh).NoOfWeeks > 0 And DayInfo(LpTh).ThSLS > 0 And DayInfo(LpTh).ThCUST > 0 Then
        PrintLine Format(DayInfo(LpTh).ThSLS / DayInfo(LpTh).NoOfWeeks, "0.00"), 5750: PrintLine Format(DayInfo(LpTh).ThSLS / DayInfo(LpTh).ThCUST, "0.00"), 6500
       Else
        PrintLine "0.00", 5750: PrintLine "0.00", 6500
       End If
       If DayInfo(LpTh).NoOfWeeks > 0 And DayInfo(LpTh).DnSLS > 0 And DayInfo(LpTh).DnCUST > 0 Then
        PrintLine Format(DayInfo(LpTh).DnSLS / DayInfo(LpTh).NoOfWeeks, "0.00"), 7250: PrintLine Format(DayInfo(LpTh).DnSLS / DayInfo(LpTh).DnCUST, "0.00"), 8000
       Else
        PrintLine "0.00", 7250: PrintLine "0.00", 8000
       End If
       If DayInfo(LpTh).NoOfWeeks > 0 And DayInfo(LpTh).LtSLS > 0 And DayInfo(LpTh).LtCUST > 0 Then
        PrintLine Format(DayInfo(LpTh).LtSLS / DayInfo(LpTh).NoOfWeeks, "0.00"), 8750: PrintLine Format(DayInfo(LpTh).LtSLS / DayInfo(LpTh).LtCUST, "0.00"), 9500
       Else
        PrintLine "0.00", 8750: PrintLine "0.00", 9500
       End If
       If DayInfo(LpTh).DaySale > 0 And DayInfo(LpTh).NoOfWeeks > 0 Then
         PrintLine Format(DayInfo(LpTh).DaySale / DayInfo(LpTh).NoOfWeeks, "0.00"), 10250
       Else
        PrintLine "0.00", 9500: PrintLine "0.00", 10250
       End If
       If DayInfo(LpTh).DaySale > 0 And DayInfo(LpTh).NoOfWeeks > 0 Then
         PrintLine Format((DayInfo(LpTh).DaySale / DayInfo(LpTh).NoOfWeeks) / (DayInfo(LpTh).LtCUST + DayInfo(LpTh).DnCUST + DayInfo(LpTh).ThCUST + DayInfo(LpTh).BkCUST + DayInfo(LpTh).LuCUST), "0.00"), -1
       Else
        PrintLine "0.00", 2750: PrintLine "0.00", -1
       End If
      End If
    Next LpTh
    DrawLine
    PrintLine "Total", 9500
    If BigT > 0 Then
      PrintLine Format(BigT, "0.00"), 10250
    Else
      PrintLine "0.00", 10250
    End If
    If BigT > 0 And TCust > 0 Then
     PrintLine Format(BigT / TCust, "0.00"), -1
    Else
     PrintLine "0.00", -1
    End If
    Call UpdateWindow(Financal_Reports.hwnd)
    DrawLine
    FinishJob 1
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Meal Period Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh1_Click()
If Can_Continue(SSCommand6.Caption) = False Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    CouponHistoryReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If

End Sub
Private Sub BtnEnh10_Click()
CustomerEdit.Visible = False
CustomerEdit.Left = -5000
End Sub
Private Sub BtnEnh11_Click()
PrinterPendingOrders
End Sub
Private Sub BtnEnh12_Click()
If Can_Continue(BtnEnh12.Caption) = False Then Exit Sub
On Error GoTo errorhandler
Dim Db  As ADODB.Connection
Dim Tb  As ADODB.Recordset
Dim Clr As Integer
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
Sql = "Select * from [customer data] Where [Customer Name] = '" & CData(0).Text & "' and [Customer Phone] = '" & CData(1).Text & "'"
ConnectADORS Sql, Db, Tb
    If Tb.EOF Then Tb.AddNew
    Tb![Customer Name] = CData(0).Text
    Tb![Customer Phone] = CData(1).Text
    Tb![Customer Address] = CData(2).Text
    Tb![Customer City] = CData(3).Text
    Tb![Customer State] = CData(4).Text
    Tb![Customer Zip] = CData(5).Text
    Tb![Instructions] = CData(6).Text & ""
    Tb![Email] = CData(7).Text & ""
    Tb![Customer Discount] = CData(9).Text & ""
    Tb.Update
DbClose Db, Tb
CustomerEdit.Visible = False
See_Msg "Saved....", 5, ""

For Clr = 0 To CData.count - 1
   CData(Clr).Text = ""
Next Clr

Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Customer ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh13_Click()
If Can_Continue(BtnEnh13.Caption) = False Then Exit Sub
If See_Msg("Remove Customer From Database ??", 4, "Delete Customer") <> 6 Then Exit Sub
On Error GoTo errorhandler
Dim Db  As ADODB.Connection
Dim Tb  As ADODB.Recordset
Dim CID As String
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
Sql = "Select * from [Customer Data] Where [Customer Name] = '" & CData(0).Text & "' and [Customer Phone] = '" & CData(1).Text & "'"
ConnectADORS Sql, Db, Tb
CID = ""
If Not Tb.EOF Then
  CID = Tb.fields("CustomerId").Value & ""
End If
Tb.Close
Sql = "Delete from [customer data] Where [Customer Name] = '" & CData(0).Text & "' and [Customer Phone] = '" & CData(1).Text & "'"
Db.Execute Sql
Sql = "Delete from [Order History] Where [Customer Name] = '" & CData(0).Text & "' and [Customer Phone] = '" & CData(1).Text & "'"
Db.Execute Sql
Sql = "Delete from [Instructions] Where [CustomerId] = '" & CID & "'"
Db.Execute Sql
Sql = "Delete from [Phone Numbers] Where [CustomerId] = '" & CID & "'"
Db.Execute Sql
DbClose Db, Tb
CustomerEdit.Visible = False
BtnEnh8_Click
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Delete Customer ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh14_Click()
Dim Cdb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim PrinterPort     As String
Dim CustName        As String
Dim Heading         As String
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Cdb
  Sql = "Select * from [Order History] Where [Customer Name] = '" & CData(0).Text & "' and [Customer Phone] = '" & CData(1).Text & "'"
  Sql = Sql & " Order by [Selection Number],[Pos] asc"
ConnectADORS Sql, Cdb, CRS
 If CRS.EOF Then
  Call See_Msg("No Orders Found", 5, "")
  DbClose Cdb, CRS
  Exit Sub
 End If
 If Check_Option(96) = 0 Then '// option(96) = 0 is for 40 column printing
    Dim PF As Integer
    PF = FreeFile
    Printer_Drivers "Receipt"
    CenterText = EC(8)
    LeftText = EC(9)
    PrinterPort = "Pending Orders"
OneMoreTime:
    Open PrinterPort For Output Shared As #PF
End If
If Check_Option(96) = 0 Then
    Print #PF, EC(1); EC(4); CenterText; "Last Order for " & CData(0).Text; EC(2); EC(7)
    Print #PF, EC(4); LeftText; EC(7)
    Print #PF, "Current Time " & Format(Time, "HH:MM AmPm"); ; EC(7)
Else
    SetPage ReportView.View, "Last Order for " & CData(0).Text
End If
CustName = ""
Do While Not CRS.EOF
 If CustName <> CRS![Customer Name] & "" Then
    If Check_Option(96) = 0 Then
      Print #PF, String(40, 45); EC(7)
      Print #PF, CRS![Customer Name] & ""; Tab(20); CRS![Customer Phone] & "" & EC(7)
      Print #PF, "Last Order Date " & Format(CRS![Check Date], "MM/DD/YYYY")
      Print #PF, String(40, 45); EC(7)
    Else
      DrawLine
      PrintLine "Name " & CRS![Customer Name] & "", 3500: PrintLine "Phone " & CRS![Customer Phone] & "", -1
      PrintLine "Last Order Date " & Format(CRS![Check Date], "MM/DD/YYYY"), -1
      DrawLine
    End If
    CustName = CRS![Customer Name] & ""
 End If
 If Check_Option(96) = 0 Then
    If Val(CRS![Quantity] & "") > 0 Then
       Print #PF, Format(CRS![Quantity], "0.00"); Tab(10); CRS![Menu Item]; EC(7)
    Else
       Print #PF, Tab(10); CRS![Menu Item]; EC(7)
    End If
 Else
    If Val(CRS![Quantity] & "") > 0 Then
       PrintLine Format(CRS![Quantity], "0.00"), 2500: PrintLine CRS![Menu Item], -1
    Else
       PrintLine "", 2500: PrintLine CRS![Menu Item], -1
    End If
 End If
 CRS.MoveNext
Loop
DbClose Cdb, CRS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Caption = "Customer Order Info"
    ReportView.Show 1
End If
Set SearchDates = Nothing
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Order History")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh15_Click()
Sql = "Select * from [Customer Data] Order by [Customer Name] asc"
CustList.Clear flexClearScrollable
Search_forCustomer Sql, True
End Sub
Private Sub BtnEnh16_Click()
CenterObject Me, CList
CList.Visible = True
CList.ZOrder 0
End Sub
Private Sub BtnEnh18_Click()
LoadCustomer ""
End Sub
Private Sub BtnEnh19_Click()
ShowCust.Visible = False
ShowCust.Left = -5000
End Sub
Private Sub BtnEnh2_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
   VoidHistoryItemReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
   Set SearchDates = Nothing
End If
Call MemoryCheck(Display)
End Sub
Private Sub BtnEnh20_Click()
On Error GoTo errorhandler
Dim Db      As ADODB.Connection
Dim Tb      As ADODB.Recordset
Dim Heading As String
Dim F       As String
Dim H       As String
Dim Data    As String
Heading = "Customer List"

F = "+2500|+1500|+3500|+1000|+1000|+750;"
H = "Name|Phone|Address|City|State|Zip;"
SetPage ReportView.View, Heading
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
Sql = "Select * from [Customer Data] Where [Customer Address] > '' order by [Customer Name] asc"
ConnectADORS Sql, Db, Tb
Data = ""
Do While Not Tb.EOF
    If Trim$(Tb![Customer Name]) > "" Then
        Data = Data & Left$(Tb![Customer Name], 30) & "" & "|"
        Data = Data & Trim$(Tb![Customer Phone] & "") & "|"
        Data = Data & Trim$(Tb![Customer Address] & "") & "|"
        Data = Data & Trim$(Tb![Customer City] & "") & "|"
        Data = Data & Trim$(Tb![Customer State] & "") & "|"
        Data = Data & Trim$(Tb![Customer Zip] & "") & ";"
    End If
Tb.MoveNext
Loop
DbClose Db, Tb
ReportView.View.TableBorder = tbBox
ReportView.View.addTable F, H, Data, vbBlack, vbWhite
ReportView.View.EndDoc
ReportView.Show 1
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Customer Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh21_Click()
PrinterPhoneSummary
End Sub
Private Sub SalesByHour(ByVal StartDate As String, ByVal EndDate As String, ByVal SaleByHow As SalesByHow)
On Error GoTo errorhandler
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim tDollars        As Currency
Dim TQty            As Long
Dim ByTransaction   As Boolean
Dim F$
Dim H$
Dim Data            As String

ReportView.View.Orientation = orPortrait
SetPage ReportView.View, "Hourly Sales " & "From " & Format(StartDate, "MM/DD/YY") & " To " & Format(EndDate, "MM/DD/YY")
H$ = ""
ConnectADODB MyRegisters(0).Svr & ProductHistory, ProductHistory, Db
Select Case SaleByHow
 Case Is = SalesByHow.ByHour
     F = "+<2000|+>1500|<+1500;"
     H = "Hour|Amount|Quantity;"
     If See_Msg("Hourly Sales By Transaction?", 4, "No for hourly sales by item") = 6 Then
        ByTransaction = True
     Else
        ByTransaction = False
     End If
     If ByTransaction Then
      Sql = "Select datepart(hh,[date]) as HR, Sum((([Check Total]-[Tax Total])-[Non Sales Total])) as DollarsSold, count([Check Total]) as QtySold From [Check Header] "
      Sql = Sql & " Where TRY_CAST([date] AS date) Between #" & CDate(StartDate) & " " & GetStartTime & "# and #" & SetEndingDate(EndDate) & " " & SetEndingTime & "#"
      Sql = Sql & " group by datepart(hh,[Date]) order by Hr asc"
    Else
     Sql = "Select datepart(hh,[Time]) as HR, Sum([Item Price]) as DollarsSold, sum([Quantity]) as QtySold From [Detail] "
     Sql = Sql & " Where TRY_CAST([date] AS datetime) Between #" & CDate(StartDate) & " " & GetStartTime & "# and #" & SetEndingDate(EndDate) & " " & SetEndingTime & "#"
     Sql = Sql & " and [Department Name] Not In ('Gift Certificate','Gratuity','Service Charge','Paid In','Gift EPay New Issue','Gift EPay ReLoad','Account Payment','Account Deposit','Non Sales Discount','OnLine Tip','Donation','Online Fee','Online Delivery')"
     Sql = Sql & " group by datepart(hh,[time]) order by Hr asc"
    End If
Case Is = SalesByHow.ByDepartment
     F = "+<1000|+<3500|+>1500|<+1500;"
     H = "Hour|Department Name|Amount|Quantity;"
     Sql = "Select datepart(hh,[time]) as HR, [Department Name], Sum([Item Price]) as DollarsSold, sum([Quantity]) as QtySold From [Detail] "
     Sql = Sql & " Where TRY_CAST([date] AS datetime) Between #" & CDate(StartDate) & " " & GetStartTime & "# and #" & SetEndingDate(EndDate) & " " & SetEndingTime & "#"
     Sql = Sql & " group by datepart(hh,[time]),[Department Name] order by Hr,[Department Name] asc"
End Select
ConnectADORS ReturnDateString(Sql), Db, RS
If RS.EOF Then
   DbClose Db, RS
   See_Msg "No Sales Found", 5, ""
   Exit Sub
End If
Dim HRS As Integer
Dim CH  As Integer
tDollars = 0: TQty = 0
Data = ""
Do While Not RS.EOF
 NST = 0
 HRS = RS![HR] + 1
 If HRS = 24 Then HRS = 0
 Select Case SaleByHow
  Case Is = SalesByHow.ByDepartment
   If CH <> HRS Then
    Data = Data & Format(CDate(HRS & ":00"), "H:MM ampm") & "|"
    Data = Data & "|"
    Data = Data & "|"
    Data = Data & ";"
    CH = HRS
   End If
   Data = Data & "|"
   Data = Data & RS![Department Name] & "|"
   Data = Data & Format(RS![DollarsSold], "$0.00") & "|"
   Data = Data & Format(RS![QtySold], "###") & ";"
  Case Is = SalesByHow.ByHour
   Data = Data & Format(CDate(HRS & ":00"), "H:MM ampm") & "|"
   Data = Data & Format(RS![DollarsSold], "$0.00") & "|"
   Data = Data & Format(RS![QtySold], "###") & ";"
 End Select
 tDollars = tDollars + RS![DollarsSold]
 TQty = TQty + RS![QtySold]
RS.MoveNext
Loop
DbClose Db, RS
ReportView.View.TableBorder = tbColTopBottom
ReportView.View.addTable F, H, Data, vbBlack, vbWhite
If SaleByHow = ByHour Then
F = "+<2000|+>1500|<+1500;"
H = "Total|Amount|Quantity;"
Else
F = "+<4500|+>1500|<+1500;"
H = "Total|Amount|Quantity;"
End If
If ByTransaction Then
 Data = "By Transaction|"
Else
 Data = "By Item|"
End If
Data = Data & Format(tDollars, "$0.00") & "|"
Data = Data & Format(TQty, "####") & ";"
ReportView.View.addTable F, H, Data, vbBlack, vbWhite
FinishJob 1
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Hourly Period Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh22_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    SalesByHour SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, ByHour
    Set SearchDates = Nothing
End If
End Sub

Private Sub BtnEnh23_Click()
On Error GoTo errorhandler
If CData(0).Text = "" Then
  See_Msg "No Customer Name", 5, "Name Required"
  Exit Sub
End If
If CData(1).Text = "" Then
  See_Msg "No Customer Phone", 5, "Phone Required"
  Exit Sub
End If
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
Sql = "Select * from [customer data] Where [Customer Name] = '" & CData(0).Text & "' and [Customer Phone] = '" & CData(1).Text & "'"
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
    Tb.AddNew
    Tb![Customer Name] = CData(0).Text
    Tb![Customer Phone] = CData(1).Text
    Tb![Customer Address] = CData(2).Text
    Tb![Customer City] = CData(3).Text
    Tb![Customer State] = CData(4).Text
    Tb![Customer Zip] = CData(5).Text
    Tb![Instructions] = CData(6).Text & ""
    Tb![Email] = CData(7).Text & ""
    Tb![Customer Discount] = CData(9).Text & ""
    Tb![Card Number] = "Not Used"
    Tb![Card Expires] = "Not Used"
    Tb![Customer Id] = 0
    Dim GRS As GeneralLib.Utilities
    Set GRS = New GeneralLib.Utilities
    Tb.fields("CustomerId").Value = GRS.CreateGuid
    Set GRS = Nothing
    Tb.Update
End If
DbClose Db, Tb
If See_Msg("Assign Coupon To Customer", 4, "Coupon To Be Used At A Later Time.") <> 6 Then Exit Sub
Reset_Button_Form
Load_What = "Coupon Reasons"
Find_What = "Select Reason"
OkToDo = True
RequestListForm.Show 1
OkToDo = False
If GetSelected = "" Then Exit Sub
HaveOneReason = GetSelected
ButtonFunction = BTNFNCT.HaveOne
AbortSelection = False
Button_Display.Show 1
If AbortSelection = True Then
  AbortSelection = False
  Exit Sub
End If
CData(9).Text = HaveOneOnMe
AttachToCustomer CData(1).Text, "Manager"
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Save coupon ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh24_Click()
If CData(7).Text = "" Then
   See_Msg "No EMail Address", 5, "Cannot Email"
   Exit Sub
End If
ReportView.SendEmailTo CData(7).Text, CData(1).Text, "Guest Services"
End Sub
Private Sub BtnEnh25_Click()
End Sub
Private Sub BtnEnh26_Click()
LoadDates "Select Distinct [Date] from [Sales] Where not isnull([date]) order by [Date] asc", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then
Else
ShowSalesTaxCollected SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
End If
Set SearchDates = Nothing
End Sub
Private Sub BtnEnh27_Click()
ShowAllAudits
End Sub
Private Sub BtnEnh28_Click()
Dim SearchData As String
SearchData = ShowKeyBoard("Enter Phone")
If SearchData = "" Then Exit Sub
Sql = "Select * from [Customer Data] where [Customer Phone] like '" & SearchData & "%' Order by [Customer Phone] asc"
CustList.Clear flexClearScrollable
Search_forCustomer Sql, True
End Sub
Private Sub ShowSalesTaxCollected(ByVal StartDate As String, ByVal EndDate As String)
On Error GoTo errorhandler
Printer_Drivers "Receipt"
Dim TotalTip        As Currency
Dim WeekTip         As Currency
Dim PrinterPort     As String
Dim PDB             As ADODB.Connection
Dim Prs             As ADODB.Recordset
Dim TaxData()       As FinanicalData
Dim W               As Integer
Dim D               As Integer
Dim RC              As Integer
Dim R               As Integer
Dim Data            As String
Dim DataDate        As String
Dim Dstr            As String
Dim TTx             As Currency
Dim C               As Integer
Dim TxSalesStr      As String
Dim TxCollect       As String
Dim TxRates         As String

ConnectADODB MyRegisters(0).Svr & ClosingHistory, ClosingHistory, PDB
    Sql = "Select Distinct [Tax Name], Sum([Tax Amount]) as TxSales from [Sales Tax]"
    Sql = Sql & " Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " Group by [Tax Name]"
    ConnectADORS (Sql), PDB, Prs
    TxSalesStr = ""
    If Prs.EOF Then
      See_Msg "No Records Found", 5, ""
      DbClose PDB, Prs
      Exit Sub
    End If
    NetTaxSales = 0
    Do While Not Prs.EOF
      TxSalesStr = TxSalesStr & Prs.fields("Tax Name").Value & "|"
      TxSalesStr = TxSalesStr & Format(Prs.fields("TxSales").Value, "0.00") & ";"
     Prs.MoveNext
    Loop
    SetNothing Prs
    Sql = "Select Distinct sum([Sales Tax]) as TSTax,Sum([Special Tax]) as TTX, Sum([Liquor Tax]) as TLX  from [Sales]"
    Sql = Sql & " Where TRY_CAST([Date] AS datetime) between '" & Format(CDate(StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
    ConnectADORS Sql, PDB, Prs
    If Not Prs.EOF Then
        TxCollect = ""
        TxCollect = GetTextFromINI("Reg Sale Tax") & "|"
        TxCollect = TxCollect & Format(Prs![TSTax], "0.00") & ";"
        TTx = Prs![TSTax]
        TxCollect = TxCollect & GetTextFromINI("Spc Sale Tax") & "|"
        TxCollect = TxCollect & Format(Prs![TTx], "0.00") & ";"
        TTx = TTx + Prs![TTx]
        TxCollect = TxCollect & GetTextFromINI("Surcharge/Tax") & "|"
        TxCollect = TxCollect & Format(Prs![TLX], "0.00") & ";"
        TTx = TTx + Prs![TLX]
    End If
    R = D
    DbClose PDB, Prs
    TxRates = ""
    If ConnectADODB(MyRegisters(0).Svr & IsSqlDataBase(SystemOptionsDB), IsSqlDataBase(SystemOptionsDB), PDB) = False Then Exit Sub
      Sql = "Select * from [Sales Tax] where [Tax Amount] > 0 Order By taxtable asc"
      ConnectADORS Sql, PDB, Prs
      Do While Not Prs.EOF
        TxRates = TxRates & Prs![Tax Name] & "|"
        TxRates = TxRates & Format(Val(Prs![Tax Amount] & ""), "0.0000") & ";"
        Prs.MoveNext
      Loop
      Prs.Close
      Sql = "Select * from [Special Tax Info] where [Special Tax] > 0 Order By taxtable asc"
      ConnectADORS Sql, PDB, Prs
      Do While Not Prs.EOF
        TxRates = TxRates & Prs![Tax Name] & "|"
        TxRates = TxRates & Format(Val(Prs![Special Tax] & ""), "0.0000") & ";"
         Prs.MoveNext
      Loop
     DbClose PDB, Prs
     ReportView.View.Orientation = orPortrait
     SetPage ReportView.View, "Sales Tax Collected " & "From " & Format(StartDate, "MM/DD/YY") & " To " & Format(EndDate, "MM/DD/YY")
     Dim F$
     Dim H$
     H$ = "Sales By Tax Rate " & "|"
     H$ = H$ & "Amount" & ";"
     F$ = "+<2400|+>1200;"
     ReportView.View.TableBorder = tbBoxColumns
     ReportView.View.addTable F, H, TxSalesStr, vbBlack, vbWhite
     H$ = "Tax Collected " & "|"
     H$ = H$ & "Amount" & ";"
     F$ = "+<2400|+>1200;"
     ReportView.View.TableBorder = tbBoxRows
     ReportView.View.addTable F, H, TxCollect, vbBlack, vbWhite
     H$ = "Tax Rates " & "|"
     H$ = H$ & "Amount" & ";"
     F$ = "+<2400|+>1200;"
     ReportView.View.TableBorder = tbBoxRows
     ReportView.View.addTable F, H, TxRates, vbBlack, vbWhite
    
    FinishJob 1
  
Exit Sub
errorhandler:
If Err.Number = 9 Then Resume Next
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Show Tax Collected ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh29_Click()
Dim SearchData As String
SearchData = ShowKeyBoard("Enter Name")
If SearchData = "" Then Exit Sub
Sql = "Select * from [Customer Data] where [Customer Name] like '" & SearchData & "%' Order by [Customer Name] asc"
CustList.Clear flexClearScrollable
Search_forCustomer Sql, True
End Sub
Private Sub BtnEnh3_Click()
CenterObject Me, CustReportMenu
CustReportMenu.ZOrder 0
CustReportMenu.Visible = True
End Sub
Private Sub BtnEnh30_Click()
Dim SearchData As String
SearchData = ShowKeyBoard("Enter Address")
If SearchData = "" Then Exit Sub
Sql = "Select * from [Customer Data] where [Customer Address] like '" & SearchData & "%' Order by [Customer address] asc"
CustList.Clear flexClearScrollable
Search_forCustomer Sql, True
End Sub
Private Sub BtnEnh31_Click()
If Can_Continue(SSCommand34.Caption) = False Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    TerminalRevenue SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End Sub
Private Sub BtnEnh32_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    SalesByHour SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, ByDepartment
    Set SearchDates = Nothing
End If

End Sub
Public Sub BtnEnh34_Click()
DatePicker.Show 1
If SearchDates Is Nothing Then Exit Sub
If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If SearchDates Is Nothing Then
    
    Else
     Get_WasteReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    End If
End If
End Sub
Private Sub BtnEnh4_Click()
If Can_Continue(BtnEnh4.Caption) = False Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    MealPeriodByGroup SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If

End Sub
Private Sub BtnEnh6_Click()
CustReportMenu.Visible = False
CustReportMenu.Left = -5000
End Sub
Private Sub BtnEnh7_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Tb As ADODB.Recordset
Dim ROD As Boolean 'Return order data
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
Sql = "Select Distinct [Customer Name],[Customer Phone] from [Order History] Where "
Sql = Sql & " TRY_CAST([Check Date] AS datetime) between '" & SearchDates.Item(1).StartDate & "'"
Sql = Sql & " and '" & SearchDates.Item(1).EndDate & "'"

ConnectADORS Sql, Db, RS
  If RS.EOF Then
     See_Msg "No Customer Data Found", 5, ""
     DbClose Db, RS
     Exit Sub
  End If
Dim F$
Dim H$
Dim Data As String
Data = ""
SetPage ReportView.View, "Phone Order Customer List"
PrintLine "From " & SearchDates.Item(1).StartDate & " To " & SearchDates.Item(1).EndDate, -1
  Do While Not RS.EOF
    Sql = "Select * from [customer data] Where [Customer Name] = '" & RS![Customer Name] & "' and [Customer Phone] = '" & RS![Customer Phone] & "'"
    ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
        Data = Data & Left$(Tb![Customer Name], 30) & "" & "|"
        Data = Data & Trim$(Tb![Customer Phone] & "") & "|"
        Data = Data & Trim$(Tb![Customer Address] & "") & "|"
        Data = Data & Trim$(Tb![Customer City] & "") & "|"
        Data = Data & Trim$(Tb![Customer State] & "") & "|"
        Data = Data & Trim$(Tb![Customer Zip] & "") & "|"
        Data = Data & Trim$(Tb![Email] & "") & ";" ' 11000
    End If
    SetNothing Tb
   RS.MoveNext
  Loop
  DbClose Db, RS
  F = "+2500|+1500|+2500|+1000|+1000|+750|+2000;"
  H = "Name|Phone|Address|City|State|Zip|Email;"
  
  ReportView.View.TableBorder = tbBox
  ReportView.View.addTable F, H, Data, vbBlack, vbWhite
  ReportView.View.EndDoc
  ReportView.Show 1
  Set SearchDates = Nothing
  Exit Sub
End Sub
Private Sub BtnEnh8_Click()
Dim Cdb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim PrinterPort     As String
Dim CustName        As String
Dim Heading         As String
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Cdb
 Sql = "Select * from [Customer Data] Order by [Customer Name] asc "
ConnectADORS Sql, Cdb, CRS
 If CRS.EOF Then
  Call See_Msg("No Customers Found", 0, "")
  DbClose Cdb, CRS
  Exit Sub
 End If
 List1.Clear
 Do While Not CRS.EOF
   List1.AddItem CRS![Customer Name] & "" & " -- " & CRS![Customer Phone] & ""
   CRS.MoveNext
 Loop
 DbClose Cdb, CRS
With PickList
   CenterObject Me, PickList
   .Visible = True
   .ZOrder 0
End With
Report_Name = "Edit Customer"
End Sub
Private Sub BtnEnh9_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then
  Exit Sub
End If
VisitReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
End Sub
Private Sub CData_DblClick(Index As Integer)
If Index < 9 Then
  CData(Index).Text = ShowKeyBoard("Enter Information")
End If
End Sub
Private Sub CData_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = KeyStroke_Ok(KeyAscii)
End Sub
Private Sub Command1_Click()
On Error GoTo ERH
Dim P             As Integer
  With Lpath
     .ShowOpen
     .Filter = "*.*"
     If .FileName > "" Then
         Dim splt() As String
         splt = Split(.FileName, "\")
         FileToSend.Text = splt(UBound(splt))
         PathforSendFile = Left$(.FileName, InStr(.FileName, FileToSend.Text) - 1)
      End If
  End With
  ChDir MyRegisters(0).RegLoc
  Exit Sub
ERH:
 See_Msg Err.Description, 0, Err.Number
End Sub
Private Sub Command2_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim Tb As ADODB.Recordset
Dim ROD As Boolean 'Return order data
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
Sql = "Select Distinct [Customer Name],[Customer Phone] from [Order History] Where "
Sql = Sql & " TRY_CAST([Check Date] AS datetime) between '" & SearchDates.Item(1).StartDate & "'"
Sql = Sql & " and '" & SearchDates.Item(1).EndDate & "'"
ConnectADORS Sql, Db, RS
  If RS.EOF Then
     See_Msg "No Customer Data Found", 5, ""
     DbClose Db, RS
     Exit Sub
  End If
  MyList.Clear
  Do While Not RS.EOF
    Sql = "Select * from [customer data] Where [Customer Name] = '" & RS![Customer Name] & "' and [Customer Phone] = '" & RS![Customer Phone] & "'"
    Sql = Sql & " and [Email] > ''"
    ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
         MyList.AddItem Tb![Email] & ""
    End If
    SetNothing Tb
   RS.MoveNext
  Loop
  DbClose Db, RS
  Exit Sub
End Sub
Private Sub Command3_Click()
CList.Visible = False
CList.Left = -5000
End Sub
Private Sub Command4_Click()
    
  If FileToSend.Text = "" Then
     MsgBox "Nothing to Send", vbCritical, "Please Select a File"
     Exit Sub
  End If
  If Subject.Text = "" Then
     MsgBox "No Subject", vbCritical, "Please add subject"
     Exit Sub
  End If
  If Message.Text = "" Then
     MsgBox "No Message", vbCritical, "Please add Message"
     Exit Sub
  End If
  If See_Msg("Send Emails ?", 4, "") <> 6 Then Exit Sub
  'FileStr = PathforSendFile & FileToSend.Text
  Dim EMails As String
  Dim Wl As Long
  EMails = ""
  For Wl = 0 To MyList.ListCount - 1
   If MyList.Selected(Wl) Then
     EMails = EMails & MyList.List(Wl) & ";"
   End If
  Next Wl
  EmailFiles PathforSendFile, FileToSend.Text, EMails, Message.Text, Subject

End Sub

Private Sub Command5_Click()
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, "House Accounts", Db) = False Then
  Exit Sub
End If
If ConnectADORS("Select [Email] from [General Data] Where [Email] > '' and [Account Type] = '" & AcctType.Text & "' order by [Account Type] desc,[Email] asc", Db, Tb) = False Then Exit Sub
MyList.Clear
If Tb.EOF Then
   See_Msg "No Accounts Found...", 5, ""
   DbClose Db, Tb
   Exit Sub
End If
Do While Not Tb.EOF
  MyList.AddItem Tb![Email] & ""
  Tb.MoveNext
Loop
DbClose Db, Tb
End Sub

Private Sub Command6_Click()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB MyRegisters(0).RegLoc & "Employee.mdb", "Employee.mdb", Db
Sql = "Select * from [General Data] Where [Phone # 2] > '' "
ConnectADORS Sql, Db, RS
  If RS.EOF Then
     See_Msg "No Emails  Found", 5, ""
     DbClose Db, RS
     Exit Sub
  End If
  MyList.Clear
  Do While Not RS.EOF
   MyList.AddItem RS![Phone # 2] & ""
   RS.MoveNext
  Loop
  DbClose Db, RS
  Exit Sub

End Sub

Private Sub CustList_Click()
If CustList.TextMatrix(CustList.Row, 0) > "" And CustList.TextMatrix(CustList.Row, 1) > "" Then
  LoadCustomer CustList.TextMatrix(CustList.Row, 0) & "" & " -- " & CustList.TextMatrix(CustList.Row, 1)
End If
End Sub

Private Sub Dayexport_Click()
End Sub

Private Sub EMailer_Click()

End Sub

Private Sub Exit_Click()
Call MemoryCheck(Closing_Report)
Call MemoryCheck(Financal_Reports)
End Sub
Public Sub StoreSalesRecap(StartDate As String, EndDate As String)
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim F$
    Dim H$
    Dim Data As String
   ConnectADODB MyRegisters(0).Svr & ClosingHistory, ClosingHistory, Db
    Sql = "Select Distinct Sum([Gross Sales]) as GSales"
    Sql = Sql & " , sum([Discounts]) as TDisc, Sum([Voids]) as TVoids, sum([Exempt Sales]) as NSaleRev"
    Sql = Sql & " , Sum([Sales Tax]) as STax, Sum([Special Tax]) as SPTax, Sum([Liquor Tax]) as SurCharge, Sum([Net Sales]) as NSales"
    Sql = Sql & " , Sum([None Taxed]) as NonTax"
    Sql = Sql & " , Sum([Tender]) as NonCash"
    Sql = Sql & " , Sum([Paid Outs]) as Pout"
    Sql = Sql & " , Sum([Deposits]) as BankDep"
    Sql = Sql & " , Sum([Dining Room]) as DineIn"
    Sql = Sql & " , Sum([Dining Room Qty]) as DineInQty"
    Sql = Sql & " , Sum([Go Sales]) as GoSales"
    Sql = Sql & " , Sum([Go Qty]) as GoSalesQty"
    Sql = Sql & " , Sum([Drive Thru Sales]) as DtSales"
    Sql = Sql & " , Sum([Drive Thru Qty]) as DtQty"
    Sql = Sql & " , Sum([Call In Sales]) as CallINSales"
    Sql = Sql & " , Sum([Call In Qty]) as CallINQty"
    Sql = Sql & " , Sum([Accounted For]) as TotalAcct"
    Sql = Sql & " , Sum([Over Short]) as OvrSht"
    Sql = Sql & " , Sum([Refunds]) as TotalReturn"
    Sql = Sql & " , Sum([Labor Hours]) as LbHrs"
    Sql = Sql & " , Sum([Labor Dollars]) as LbDollars"
    
    Sql = Sql & " from [Sales] "
    Sql = Sql & " Where TRY_CAST([date] AS datetime) between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    'Sql = Sql & " Group By [Order Destination]"
   ConnectADORS (Sql), Db, RS
    If RS.EOF Then
       See_Msg "No Sale Found", 5, ""
       DbClose Db, RS
       Exit Sub
    Else
       SetRegular 0
       Data = Format(RS![GSales], "0.00") & "|"
       Data = Data & Format(RS![TDisc], "0.00") & "|"
       Data = Data & Format(RS![TVoids], "0.00") & "|"
       Data = Data & Format(RS![NSaleRev], "0.00") & "|"
       Data = Data & Format(RS![NSales], "0.00") & "|"
       Data = Data & Format(RS![STAX], "0.00") & "|"
       Data = Data & Format(RS![SPTax], "0.00") & "|"
       Data = Data & Format(RS![Surcharge], "0.00") & ";"
       F = "+>1500|+>1500|+>1500|+>1500|+>1500|+>1500|+>1500|+>1500;"
       H = "Gross Sales|Discounts|Voids|Non Sales|Net Sales|Tax 1|Tax 2|SurCharge;"
       ReportView.View.Orientation = orLandscape
       ReportView.View.StartDoc
       ReportView.View.StartTable
       ReportView.View.TableBorder = tbBoxColumns
       ReportView.View.addTable F, H, Data, vbBlack, vbWhite
       ReportView.View.EndTable
      ' F = "+>1500|+>1500|+>1500|+>1500|+>1500|+>1500|+>1500|+>1500;"
      ' H = "Taxed Sales|Non Taxed|Voids|Non Sales|Net Sales|Tax 1|Tax 2|SurCharge;"
      ' ReportView.View.Orientation = orLandscape
      ' ReportView.View.StartDoc
      ' ReportView.View.StartTable
      ' ReportView.View.TableBorder = tbBoxColumns
      ' ReportView.View.AddTable F, H, Data, vbBlack, vbWhite
      ' ReportView.View.EndTable
    
    End If
    
    
    
    FinishJob 1
DbClose Db, RS
   ' Sql = "Select Distinct [Order Destination], sum([Amount]) as tAmount, Sum([Quantity]) as QtySold from [Order Destination] "
   ' Sql = Sql & " Where not Isnull([Date]) and TRY_CAST([date] AS datetime) between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
   ' Sql = Sql & " Group By [Order Destination]"

End Sub

Private Sub Form_Load()
' Dim buttons As Collection
' Dim b As BtnEnh

' Set buttons = GetBtnEnhControls(Me)
  
' For Each b In buttons
    
    
 '   b.Shape = SHAPE_ROUNDRECT

' Next


ResizeScreen1024 Me
CenterForm Me
'DaysSinceClose
'LoadBackground Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
EndClosing
End Sub
Private Sub Search_forCustomer(InSql As String, ShowMsg As Boolean)
On Error GoTo errorhandler
Dim Cdb As ADODB.Connection, CRS As ADODB.Recordset
Dim Cl As Long
ShowCust.ZOrder 0
'ShowCust.Left = 240 * WidScale800x600: ShowCust.Top = 120 * TopScale800x600
ShowCust.width = 14000 * WidScale1024
CenterObject Me, ShowCust
With CustList
   .ColWidth(0) = 2500 * WidScale1024: .ColWidth(1) = 2000 * WidScale1024: .ColWidth(2) = 3000 * WidScale1024:
   .ColWidth(3) = 1150 * WidScale1024: .ColWidth(4) = 1150 * WidScale1024: .ColWidth(5) = 1150 * WidScale1024: .ColWidth(6) = 2000 * WidScale1024
   .TextMatrix(0, 0) = "Name": .TextMatrix(0, 1) = "Phone": .TextMatrix(0, 2) = "Address": .TextMatrix(0, 3) = "City"
   .TextMatrix(0, 4) = "State": .TextMatrix(0, 5) = "Zip": .TextMatrix(0, 6) = "Instructions"
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Cdb
   'Sql = "Select * from [customer data] order by [customer name] asc"
ConnectADORS InSql, Cdb, CRS
  .cols = 7
     .rows = RecordsCount(CRS) + 1
  Cl = 0
  Do While Not CRS.EOF
    Cl = Cl + 1
    .TextMatrix(Cl, 0) = Check_Null(CRS![Customer Name]) & ""
    .TextMatrix(Cl, 1) = Check_Null(CRS![Customer Phone]) & ""
    .TextMatrix(Cl, 2) = Check_Null(CRS![Customer Address]) & ""
    .TextMatrix(Cl, 3) = Check_Null(CRS![Customer City]) & ""
    .TextMatrix(Cl, 4) = Check_Null(CRS![Customer State]) & ""
    .TextMatrix(Cl, 5) = Check_Null(CRS![Customer Zip]) & ""
    .TextMatrix(Cl, 6) = Check_Null(CRS![Instructions]) & ""
   CRS.MoveNext
  Loop
 DbClose Cdb, CRS
 If Cl = 0 Then
   Call See_Msg("No Records Found", 5, "")
 Else
 ShowCust.Visible = True
 ShowCust.ZOrder 0
 'Picture1.Visible = False
 End If
  CustList.width = 10500 * WidScale800x600
End With
Exit Sub

'
Exit Sub
errorhandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Search For Customer ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub List1_Click()
If Report_Name = "Edit Customer" Then
   LoadCustomer List1.List(List1.ListIndex)
Else
  If PaymentbyHour.Value = 0 Then
    If List1.List(List1.ListIndex) = "Today" Then
      ShowPaymentByTerminal
    
    ElseIf List1.List(List1.ListIndex) = "Week To Date" Then
       Store_Payment False, -1
    Else
       Store_Payment True, List1.ListIndex
    End If
  Else
    If List1.List(List1.ListIndex) = "Week To Date" Then
       HourlyPayments List1.List(List1.ListIndex), -1
    Else
       HourlyPayments List1.List(List1.ListIndex), List1.ListIndex
    End If
  
  End If
End If
End Sub
Public Sub ShowPaymentByTerminal()
    ReportView.View.Orientation = orPortrait
    SetPage ReportView.View, "Terminal Payment Report " & Format(Date, "MM/DD/YYYY")
    Dim FA As Integer
    Dim TotalAmt As Currency
    Dim Data As String
    Data = "": TotalAmt = 0
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
       ConnectADODB MyRegisters(0).Svr & FinancialDb, FinancialDb, Db
        Sql = "Select Distinct [Register Name] ,[Payment Type], Sum([Payment Amount]) as TotalPayment, Sum([Tip Amount]) as TotalTip from [payments]"
        Sql = Sql & " Group by [Register Name], [Payment Type]"
       ConnectADORS Sql, Db, RS
        Do While Not RS.EOF
          Data = Data & RS![Register Name] & "|"
          Data = Data & RS![Payment Type] & "|"
          Data = Data & Format(RS![TotalPayment] + RS![TotalTip], "0.00") & ";"
          TotalAmt = TotalAmt + (RS![TotalPayment] + RS![TotalTip])
        RS.MoveNext
       Loop
      DbClose Db, RS
    
    'If EMVDeviceIsPresent Then
      ConnectADODB MyRegisters(0).Svr & DebitDb, DebitDb, Db
        Sql = "Select Distinct [Register Name] ,[Account Name], Sum([Amount]) as TotalPayment, Sum([Tip]) as TotalTip from [Debit]"
        Sql = Sql & " Where [Tender] = 'EMV' or [Tender] = 'QRPayment'"
        Sql = Sql & " Group by [Register Name], [Account Name]"
       ConnectADORS Sql, Db, RS
      Do While Not RS.EOF
          Data = Data & RS![Register Name] & "|"
          Data = Data & RS![Account Name] & "|"
          Data = Data & Format(RS![TotalPayment] + RS![TotalTip], "0.00") & ";"
          TotalAmt = TotalAmt + (RS![TotalPayment] + RS![TotalTip])
       RS.MoveNext
      Loop
    'End If
    Data = Data & "Total" & "|"
    Data = Data & "" & "|"
    Data = Data & Format(TotalAmt, "0.00") & ";"
  DbClose Db, RS
  F = "+<2500|+<2500|+>1500;"
  H = "Register Name|Media|Amount;"
  ReportView.View.MarginLeft = ".20in"
  ReportView.View.CurrentY = 2700
  ReportView.View.StartTable
  ReportView.View.TableBorder = tbBoxColumns
  ReportView.View.addTable F, H, Data, vbYellow
  ReportView.View.EndTable
  ReportView.View.EndDoc
  ReportView.Show
End Sub
Private Sub LoadCustomer(DataString As String)
On Error GoTo errorhandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim InData() As String
If DataString > "" Then
    InData = Split(DataString, " -- ")
    ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
    Sql = "Select * from [customer data] Where [Customer Name] = '" & InData(0) & "' and [Customer Phone] = '" & InData(1) & "'"
    ConnectADORS Sql, Db, Tb
      If Not Tb.EOF Then
            CData(0).Text = Tb![Customer Name] & ""
            CData(1).Text = Tb![Customer Phone] & ""
            CData(2).Text = Tb![Customer Address] & ""
            CData(3).Text = Tb![Customer City] & ""
            CData(4).Text = Tb![Customer State] & ""
            CData(5).Text = Tb![Customer Zip] & ""
            CData(6).Text = Tb![Instructions] & ""
            CData(7).Text = Tb![Email] & ""
            CData(8).Text = GetLastOrderDate(Tb![Customer Name] & "", Tb![Customer Phone] & "")
            CData(9).Text = Tb![Customer Discount] & ""
      End If
    DbClose Db, Tb
End If
CenterObject Me, CustomerEdit
CustomerEdit.Visible = True
CustomerEdit.ZOrder 0
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Customer ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub MailCoupons_Click()

End Sub

Private Sub MemberBox_Click()

End Sub

Private Sub MReport_Click()


End Sub
Private Sub LoadExport()
Dim Cdb As ADODB.Connection
Dim RS As ADODB.Recordset
Dayexport.Clear
ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Cdb
Sql = "Select Distinct [Close Out Day] from [Check Detail] order by [close out day] asc"
ConnectADORS Sql, Cdb, RS
Do While Not RS.EOF
   Dayexport.AddItem GetDayOfWeek(Val(RS![Close out day] & ""))
  RS.MoveNext
Loop
DbClose Cdb, RS
End Sub

Private Sub PendingList_Click()
On Error GoTo errorhandler
If PendingList.List(PendingList.ListIndex) = "Exit" Then
   PendingList.Visible = False
   Exit Sub
End If
PendingList.Visible = False
Dim Cdb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim PrinterPort     As String
Dim CustName        As String
Dim NeedDate        As String
ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Cdb
 Sql = "Select * from [Check Detail] Where [Remote 1] = '" & PendingList.List(PendingList.ListIndex) & "'"
 Sql = Sql & " or [Remote 2] = '" & PendingList.List(PendingList.ListIndex) & "'"
 Sql = Sql & " or [Remote 3] = '" & PendingList.List(PendingList.ListIndex) & "'"
 Sql = Sql & " or [Remote 4] = '" & PendingList.List(PendingList.ListIndex) & "'"
 Sql = Sql & " or [Remote 5] = '" & PendingList.List(PendingList.ListIndex) & "'"
 Sql = Sql & " or [Remote 6] = '" & PendingList.List(PendingList.ListIndex) & "'"
If SvrPath.sServerIsOn = 1 Then
  Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & SearchDates.Item(1).StartDate & "'"
  Sql = Sql & " and '" & SearchDates.Item(1).EndDate & "'"
  Sql = Sql & "Order By TRY_CAST([Check date] AS datetime),[Check Number],[Selection Number],[Pos] asc"
Else
  Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & Format(CDate(SearchDates.Item(1).StartDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(SearchDates.Item(1).EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & "Order By TRY_CAST([Check date] AS datetime),[Check Number],[Selection Number],[Pos] asc"
End If
ConnectADORS Sql, Cdb, CRS
 If CRS.EOF Then
  Call See_Msg("No Orders Found", 0, "")
  DbClose Cdb, CRS
  Exit Sub
 End If
If Check_Option(96) = 0 Then
    Dim PF As Integer
    PF = FreeFile
    Printer_Drivers "Receipt"
    CenterText = EC(8)
    LeftText = EC(9)
    'RightText = Chr(27) + Chr(97) + Chr(2)
    PrinterPort = "Pending Orders"
OneMoreTime:
    Open PrinterPort For Output Shared As #PF
End If
Heading = "Pending Orders For " & SearchDates.Item(1).StartDate
If Check_Option(96) = 0 Then
    Print #PF, EC(1); EC(4); CenterText; Heading; EC(2); EC(7)
    Print #PF, EC(4); LeftText; EC(7)
    Print #PF, "Current Time " & Format(Time, "HH:MM AmPm"); ; EC(7)
Else
    SetPage ReportView.View, "Pending Orders"
End If
CustName = "": NeedDate = ""
Do While Not CRS.EOF
 If CustName <> CRS![Customer Name] & "" Or NeedDate <> CRS![Check Date] & "" Then
    If Check_Option(96) = 0 Then
      Print #PF, String(40, 45); EC(7)
      Print #PF, CRS![Customer Name] & ""; Tab(20); CRS![Customer Phone] & "" & EC(7)
      Print #PF, "Date Needed " & Format(CRS![Check Date], "MM/DD/YYYY")
      Print #PF, String(40, 45); EC(7)
    Else
      DrawLine
      PrintLine "Name " & CRS![Customer Name] & "", 3500: PrintLine "Phone " & CRS![Customer Phone] & "", -1
      PrintLine "Date Needed " & Format(CRS![Check Date], "MM/DD/YYYY"), -1
      DrawLine
    End If
    CustName = CRS![Customer Name] & ""
    NeedDate = CRS![Check Date] & ""
 End If
 If Check_Option(96) = 0 Then
    If Val(CRS![Quantity] & "") > 0 Then
       Print #PF, Format(CRS![Quantity], "0.00"); Tab(10); CRS![Menu Item]; EC(7)
    Else
       Print #PF, Tab(10); CRS![Menu Item]; EC(7)
    End If
 Else
    If Val(CRS![Quantity] & "") > 0 Then
       PrintLine Format(CRS![Quantity], "0.00"), 2500: PrintLine CRS![Menu Item], -1
    Else
       PrintLine "", 2500: PrintLine CRS![Menu Item], -1
    End If
 End If
 CRS.MoveNext
Loop
DbClose Cdb, CRS
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Caption = "Pending Orders"
    ReportView.Show 1
End If
Set SearchDates = Nothing
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Pending Orders ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub SALL_Click()
Dim SA As Long
If SALL.Caption = "UnSelect" Then
    For SA = 0 To MyList.ListCount - 1
       MyList.Selected(SA) = False
    Next SA
    SALL.Caption = "Select All"
Else
    For SA = 0 To MyList.ListCount - 1
       MyList.Selected(SA) = True
    Next SA
    SALL.Caption = "UnSelect"
End If
End Sub

Public Sub SSCommand1_Click()
If Check_Option(199) = 1 Then   '// option(199) disables daily close from a station.
   See_Msg "Cannot Close From This Station", 5, "Close Out Disabled."
   Exit Sub
End If
If Can_Continue(SSCommand1.Caption) = False Then Exit Sub
On Error GoTo errorhandler
Dim AuditsOpen  As Boolean
Dim SRC         As Integer
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset

SETvariables
IsCommonPresent
FindNoPayments
SuspendAllOpenChecks


   If UCase(GetConfigInfo("Daily Close", "Close Day With Open Phone Orders", "NO", "FUNCTIONS.INI")) = "ENABLED" Then
     If GetConfigInfo("Daily Close", "Close Day With PrePaid Orders", "Yes", "FUNCTIONS.INI") = "DISABLED" Then
      ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
       Sql = "Select * from [Open Checks] where [Payments] > 0"
      ConnectADORS Sql, Db, RS
      If Not RS.EOF Then
         See_Msg "Please Close All PrePaid Phone Orders", 5, "Cannot Close With PrePaid Orders"
         DbClose Db, RS
         Exit Sub
      End If
      DbClose Db, RS
     End If
   Else
      ConnectADODB MyRegisters(0).Svr & CallInOrderDb, CallInOrderDb, Db
        Sql = "Select * from [Open Checks] where [Presented] = 1"
      ConnectADORS Sql, Db, RS
        If Not RS.EOF Then
          See_Msg "Please Close All Presented Phone Orders", 5, "Open Phone Orders"
          DbClose Db, RS
          Exit Sub
        End If
        DbClose Db, RS
   End If
   ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Db
   ConnectADORS "Select * from [Open Checks] where [Close Out Day] = 0", Db, RS
     If Not RS.EOF Then
        See_Msg "There is " & CStr(RecordsCount(RS)) & " Check(s) Open", 5, "Can Not Close Out"
        DbClose Db, RS
        Exit Sub
     End If
     DbClose Db, RS
   AuditsOpen = False
   ConnectADODB MyRegisters(0).Svr & FinancialDb, FinancialDb, Db
   ConnectADORS "Select * from [Audits]", Db, RS
        If Not RS.EOF Then
          If See_Msg("Cashier Audit Open On " & MyRegisters(0).regName, 4, "Close Audit ?") = 6 Then
            Do While Not RS.EOF
              Transaction.Close_Audit RS![pos], False, RS![Register Name], MyRegisters(0).Svr, False, 0
              RS.MoveNext
            Loop
          Else
           AuditsOpen = True
          End If
        End If
        DbClose Db, RS
        ConnectADODB MyRegisters(0).Svr & "MSaleData", "MSaleData", Db
          Sql = "Select * from [Check Numbers] Where [Audit Status] = 0"
        ConnectADORS Sql, Db, RS
         If Not RS.EOF Then
          If See_Msg("Server Audits Open On " & MyRegisters(0).SvrName, 4, "Close All Server Audits ? ") <> 6 Then
           AuditsOpen = True
          Else
           AuditsOpen = False
          End If
         End If
         DbClose Db, RS
   If AuditsOpen = True Then Exit Sub
   CloseAllServerAudits
   If ALLClockedOut = False Then Exit Sub
   Closing_Report.Show
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Financial Report ")
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
Dim ScreenReport As Boolean
If Report_Name = "Dept" Then
  Dept_Report TaxedSales.Text, "Daily " & Format(Date, "M/D/YY"), 0, "Day", "", 0, True
ElseIf Report_Name = "Rev" Then
  TableRev True, True
ElseIf Report_Name = "Term" Then
  TerminalReport True, True
Else
  ScreenReport = True
  If Check_Option(96) = 0 Then '// option(96) on for 40 column printing
     If See_Msg("Print Report", 4, "") = 6 Then
        ScreenReport = False
     End If
  End If
  Void_Report "Day", "", ScreenReport, False
End If
Report_Type.Visible = False
SSCommand7_Click
End Sub
Private Sub SSCommand11_Click()
If Can_Continue(SSCommand11.Caption) = False Then Exit Sub
Report_Name = "Dept"
CenterObject Me, Report_Type
Report_Type.ZOrder 0
Report_Type.Visible = True
End Sub
Private Sub SSCommand12_Click()
If Can_Continue(SSCommand12.Caption) = False Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"

If SearchDates Is Nothing Then

Else
    VoidItemReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End Sub
Private Sub SSCommand14_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
  DestinationReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
  Set SearchDates = Nothing
End If
End Sub
Public Sub DestinationReport(ByVal StartDate As String, ByVal EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort As String
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim TotSales    As Currency
Dim Cntr        As Long
Dim TRCD        As Long
ConnectADODB MyRegisters(0).Svr & ClosingHistory, ClosingHistory, Db
    Sql = "Select Distinct [Order Destination], sum([Amount]) as tAmount, Sum([Quantity]) as QtySold from [Order Destination] "
    Sql = Sql & " Where not Isnull([Date]) and TRY_CAST([date] AS datetime) between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    Sql = Sql & " Group By [Order Destination]"
ConnectADORS (Sql), Db, RS
If RS.EOF Then
   See_Msg "No Sales Found", 5, ""
   DbClose Db, RS
   Exit Sub
End If
TRCD = RecordsCount(RS)
TotSales = 0
Do While Not RS.EOF
  TotSales = TotSales + Val(RS![TAmount] & "")
  RS.MoveNext
Loop
If Not RS.BOF Then RS.MoveFirst
If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    CenterText = EC(8)
    LeftText = EC(9)
    'RightText = Chr(27) + Chr(97) + Chr(2)
    PF = FreeFile
    PrinterPort = "Destination Report"
tryagain:
    Open PrinterPort For Output Shared As #PF
    Print #PF, EC(1); CenterText; Heading; EC(7)
Else
    SetPrinter 0
    PrintCLine CStr(Heading), -1, 0
End If
DoEvents
If Check_Option(96) = 0 Then
    Print #PF, EC(7)
    Print #PF, EC(1); CenterText; "Period Begin "; StartDate
    Print #PF, EC(1); CenterText; "Period Ending  "; EndDate
    Print #PF, LeftText; EC(7)
    Print #PF, ; EC(4); EC(7)
    Print #PF, Format(Time, "HH:MM AmPm");
    Print #PF, Tab(24); "Date  "; Trim(Date); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Order Type      Qty      $         % "
    Print #PF, String(40, 45); EC(7)
Else
    PrintLine "", -1
    PrintCLine "Period Begin " & StartDate, -1, 0
    PrintCLine "Period Ending  " & EndDate, -1, 0
    PrintLine "", -1
    PrintLine Format(Time, "HH:MM ampm"), -1
    PrintLine Format(Date, "MM:DD:YYYY") & " Printed By " & LoggedOnAs(0).EMPName, -1
    DrawLine
    PrintLine "Order Type", 4000
    PrintLine "Quantity", 6000
    PrintLine "Sales", 7500
    PrintLine "Percent", -1
    DrawLine
End If
Dim DSP As Double
  Do While Not RS.EOF
    If Check_Option(96) = 0 Then
       Print #PF, RS![Order Destination] & ""; Tab(17); Format(Val(RS![QtySold] & ""), "####");
       Print #PF, Tab(24); Align(8, Val(RS![TAmount] & ""));
       If TotSales > 0 And Val(RS![TAmount] & "") > 0 Then
          DSP = (RS![TAmount] / TotSales) * 100
       Else
          DSP = 0
       End If
       Print #PF, Tab(35); Format(DSP, "0.00")
    Else
       If TotSales > 0 And Val(RS![TAmount] & "") > 0 Then
         DSP = (RS![TAmount] / TotSales) * 100
       Else
        DSP = 0
       End If
       PrintLine RS![Order Destination] & "", 4000
       PrintLine Format(Val(RS![QtySold] & ""), "####"), 6000
       PrintLine Format(Val(RS![TAmount]) & "", "0.00"), 7500
       PrintLine Format(DSP, "0.00"), -1
    End If
    RS.MoveNext
  Loop
  DbClose Db, RS
    If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, Tab(17); "Total Sales  "; Align(10, TotSales)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
    Else
    DrawLine
    PrintLine "Total Sales  ", 6000: PrintLine Format(TotSales, "0.00"), -1
    FinishJob 1
    End If
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Destination Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub SSCommand16_Click()
SSCommand16.Caption = "Working....."
SSCommand16.Refresh
LoadDates "Select Distinct [Date] from [Sales] order by [Date] asc", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then
Else
Sales_History Report_Name, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
End If
Set SearchDates = Nothing
SSCommand16.Caption = "Sales $ Report"
SSCommand16.Refresh

End Sub
Private Sub SSCommand17_Click()
End Sub
Private Sub SSCommand18_Click()
SSCommand18.Caption = "Working....."
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    Mix_History Report_Name, "Category", SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
SSCommand18.Caption = "Category Report"
SSCommand18.Refresh

End Sub
Private Sub SSCommand19_Click()
SSCommand19.Caption = "Working....."
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    Mix_History Report_Name, "Department", SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
SSCommand19.Caption = "Department Report"
SSCommand19.Refresh
End Sub
Private Sub SSCommand25_Click()
If Can_Continue(SSCommand25.Caption) = False Then Exit Sub
Display.Show
End Sub
Private Sub SSCommand27_Click()
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    SummaryReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If

End Sub
Public Sub SummaryReport(StartDate As String, EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort     As String
Dim Heading         As String
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim GST             As Currency
Dim PAMT            As Single
Dim CashBankDeposit As Currency
Dim F$
Dim H$
Dim Data            As String
ConnectADODB MyRegisters(0).Svr & ClosingHistory, ClosingHistory, Db

Sql = "Select distinct sum([Gross Sales]) as GSales,"
Sql = Sql & " sum([Net Sales]) as NSales,"
Sql = Sql & " sum([Sales Tax]) as NST,"
Sql = Sql & " sum([Special Tax]) as NSPT,"
Sql = Sql & " sum([Liquor Tax]) as NLT,"
Sql = Sql & " sum([Voids]) as NVoid,"
Sql = Sql & " sum([Discounts]) as NDisc,"
Sql = Sql & " sum([Paid Outs]) as PayOut,"
Sql = Sql & " sum([None Taxed Sales]) as NonTaxed,"
Sql = Sql & " sum([Exempt Sales]) as NoneSaleRev,"
Sql = Sql & " sum([Refunds]) as Refund,"
Sql = Sql & " sum([Deposits]) as Deposit"
Sql = Sql & " From [Sales] "
Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), Db, RS
If RS.EOF Then
   DbClose Db, RS
   See_Msg "No Sales Found", 5, ""
   Exit Sub
End If
If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    CenterText = EC(8)
    LeftText = EC(9)
    'RightText = Chr(27) + Chr(97) + Chr(2)
    PF = FreeFile
    PrinterPort = "Summary Report"
tryagain:
    Open PrinterPort For Output Shared As #PF
    Print #PF, EC(1); CenterText; "Sales Summary Report"; EC(7)
Else
    SetPage ReportView.View, "Sales Summary Report" & vbCrLf & "Period " & StartDate & " To " & EndDate
End If
Data = ""
If Check_Option(96) = 0 Then
    Print #PF, EC(7)
    Print #PF, EC(1); CenterText; "Begin Period "; StartDate
    Print #PF, EC(1); CenterText; "End Period   "; EndDate
    Print #PF, LeftText; EC(7)
    Print #PF, ; EC(4); EC(7)
    Print #PF, Format(Time, "HH:MM AmPm");
    Print #PF, Tab(24); "Date  "; Trim(Date); EC(7)
    Print #PF, String(40, 45); EC(7)
End If
CashBankDeposit = Val(RS![Deposit] & "")
If Check_Option(96) = 0 Then
    Print #PF, "Gross Receipts    "; Tab(24); Format(RS![GSales] + ((RS![NST] + RS![NSPT]) + RS![NLT]), "0.00"); EC(7)
    Print #PF, "Gross Sales       "; Tab(24); Format(RS![GSales], "0.00"); EC(7)
    Print #PF, "Refunds           "; Tab(24); Format(RS![Refund], "0.00"); EC(7)
    Print #PF, "Net Sales         "; Tab(24); Format(RS![NSales], "0.00"); EC(7)
    Print #PF, "Non Taxed Sales   "; Tab(24); Format(RS![NonTaxed], "0.00"); EC(7)
    Print #PF, "Non Sales Revenue "; Tab(24); Format(RS![NoneSaleRev], "0.00"); EC(7)
    Print #PF, "Net Sales Tax   "; Tab(24); Format(RS![NST], "0.00"); EC(7)
    Print #PF, "Net Special Tax "; Tab(24); Format(RS![NSPT], "0.00"); EC(7)
    Print #PF, "Net " & SurChargeTxt; Tab(24); Format(RS![NLT], "0.00"); EC(7)
    Print #PF, "Discounts "; Tab(24); Format(RS![NDisc], "0.00"); EC(7)
    Print #PF, "Voided Sales "; Tab(24); Format(RS![NVoid], "0.00"); EC(7)
    Print #PF, "Net Sales w/Tax "; Tab(24); Format(RS![NSales] + RS![NST] + RS![NSPT] + RS![NLT], "0.00"); EC(7)
Else
    Data = Data & "Gross Receipts" & "|" & Format(RS![GSales] + ((RS![NST] + RS![NSPT]) + RS![NLT]), "0.00") & ";"
    Data = Data & "Gross Sales" & "|" & Format(RS![GSales], "0.00") & ";"
    Data = Data & "Refunds" & "|" & Format(RS![Refund], "0.00") & ";"
    Data = Data & "Net Sales" & "|" & Format(RS![NSales], "0.00") & ";"
    Data = Data & "Non Taxed Sales" & "|" & Format(RS![NonTaxed], "0.00") & ";"
    Data = Data & "Non Sales Revenue" & "|" & Format(RS![NoneSaleRev], "0.00") & ";"
    Data = Data & "Net Sales Tax" & "|" & Format(RS![NST], "0.00") & ";"
    Data = Data & "Net Special Tax" & "|" & Format(RS![NSPT], "0.00") & ";"
    Data = Data & "Net " & SurChargeTxt & "|" & Format(RS![NLT], "0.00") & ";"
    Data = Data & "Discounts" & "|" & Format(RS![NDisc], "0.00") & ";"
    Data = Data & "Voided Sales" & "|" & Format(RS![NVoid], "0.00") & ";"
    Data = Data & "Net Sales Including All Tax" & "|" & Format(RS![NSales] + RS![NST] + RS![NSPT] + RS![NLT], "0.00") & ";"
    F = "+5000|+>1500;"
    H = "Sales|Amount;"
    ReportView.View.TableBorder = tbBox
    ReportView.View.addTable F, H, Data, vbBlack, vbWhite
End If
Data = ""
GST = Val(RS![NSales] & "")
SetNothing RS
Sql = "Select distinct [Name],"
Sql = Sql & " sum([Amount]) as TAmt,"
Sql = Sql & " sum([Quantity]) as TQty"

Sql = Sql & " from [Tenders]  "
Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and not isnull([date])"
Sql = Sql & " Group By [Name]"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
 If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Bank Deposits"; Tab(24); Format(CashBankDeposit, "0.00"); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Tenders"; EC(7)
    Print #PF, String(40, 45); EC(7)
 Else
    Data = "Bank Deposits" & "|" & Format(CashBankDeposit, "0.00") & ";"
    F = "+5000|+>1500;"
    H = "Bank|Amount;"
    ReportView.View.TableBorder = tbBox
    ReportView.View.addTable F, H, Data, vbBlack, vbWhite
 End If
 Data = ""
 Do While Not RS.EOF
    If Check_Option(96) = 0 Then
        Print #PF, RS![Name] & ""; Tab(24); Format(Val(RS![TAmt] & ""), "0.00"); Tab(34); Format(Val(RS![TQty] & ""), "##"); EC(7)
    Else
        Data = Data & RS![Name] & "" & "|" & Format(Val(RS![TAmt] & ""), "0.00") & "|" & Format(Val(RS![TQty] & ""), "##") & ";"
    End If
    RS.MoveNext
 Loop
 If Check_Option(96) = 1 Then
    F = "+5000|+>1500|+>1500;"
    H = "Tenders|Amount|Quantity;"
    ReportView.View.TableBorder = tbBox
    ReportView.View.addTable F, H, Data, vbBlack, vbWhite
  End If
End If
SetNothing RS
Sql = "Select distinct [Name],"
Sql = Sql & " sum([Amount]) as POAmt,"
Sql = Sql & " sum([Quantity]) as POQAmt"
Sql = Sql & " from [Paid Outs]  "
Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and not isnull([date])"
Sql = Sql & " Group By [Name]"
ConnectADORS Sql, Db, RS
Data = ""
If Not RS.EOF Then
    If Check_Option(96) = 0 Then
       Print #PF, String(40, 45); EC(7)
       Print #PF, "Paid Outs"; EC(7)
       Print #PF, String(40, 45); EC(7)
    End If
    Do While Not RS.EOF
     If Check_Option(96) = 0 Then
      Print #PF, RS![Name] & ""; Tab(24); Format(Val(RS![POAmt] & ""), "0.00"); Tab(34); Format(Val(RS![poQAmt] & ""), "##"); EC(7)
     Else
      Data = Data & RS![Name] & "" & "|" & Format(Val(RS![POAmt] & ""), "0.00") & "|" & Format(Val(RS![poQAmt] & ""), "##") & ";"
     End If
    RS.MoveNext
    Loop
    If Check_Option(96) = 1 Then
       F = "+5000|+>1500|+>1500;"
       H = "Paid Outs|Amount|Quantity;"
       ReportView.View.TableBorder = tbBox
       ReportView.View.addTable F, H, Data, vbBlack, vbWhite
    End If
End If
SetNothing RS
Sql = "Select distinct [Name],"
Sql = Sql & " sum([Amount]) as DAmt,"
Sql = Sql & " sum([Quantity]) as QAmt"
Sql = Sql & " from [Discounts]  "
Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
Sql = Sql & " and not isnull([date])"
Sql = Sql & " Group By [Name]"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Discounts"; EC(7)
    Print #PF, String(40, 45); EC(7)
End If
Do While Not RS.EOF
If Check_Option(96) = 0 Then
  Print #PF, RS![Name] & ""; Tab(24); Format(Val(RS![DAmt] & ""), "0.00"); Tab(34); Format(Val(RS![QAmt] & ""), "##"); EC(7)
Else
  Data = Data & RS![Name] & "" & "|" & Format(Val(RS![DAmt] & ""), "0.00") & "|" & Format(Val(RS![QAmt] & ""), "##") & ";"
End If
RS.MoveNext
Loop
 If Check_Option(96) = 1 Then
   F = "+5000|+>1500|+>1500;"
   H = "Discounts|Amount|Quantity;"
   ReportView.View.TableBorder = tbBox
   ReportView.View.addTable F, H, Data, vbBlack, vbWhite
 End If
End If
SetNothing RS
Sql = "Select distinct [Name],"
Sql = Sql & " sum([Amount]) as DPAmt,"
Sql = Sql & " sum([Quantity]) as DPQAmt"
Sql = Sql & " from [Department Sales]  "
Sql = Sql & " Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
If FileCheck(MyRegisters(0).RegLoc & "INI Files\" & "\BrewPub.ini") Then
  Sql = Sql & " and [Name] <> 'Coupon' and [Name] <> 'Discount' "
End If  'added for central cash register mike 11/30/2004
Sql = Sql & " and not isnull([date])"
Sql = Sql & " Group By [Name]"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Department Sales"; EC(7)
    Print #PF, String(40, 45); EC(7)
End If
'*******************************
If FileCheck(MyRegisters(0).RegLoc & "INI Files\" & "\BrewPub.ini") Then
    GST = 0
    Do While Not RS.EOF
    GST = GST + Val(RS![dpamt] & "")
    RS.MoveNext
    Loop
    If Not RS.BOF Then RS.MoveFirst
End If
'***************************/***
Data = ""
Do While Not RS.EOF
If GST > 0 And Val(RS![dpamt] & "") <> 0 Then
 PAMT = (RS![dpamt] / GST)
  Select Case RS![Name]
    Case Is = "Coupon", "Discount"
      PAMT = PAMT * -1
  End Select
Else
 PAMT = 0
End If
If Check_Option(96) = 0 Then
  Print #PF, RS![Name] & ""; Tab(24); Format(Val(RS![dpamt] & ""), "0.00"); Tab(34); Format(Val(RS![DPQAmt] & ""), "##"); EC(7)
Else
  Data = Data & RS![Name] & "" & "|" & Format(Val(RS![dpamt] & ""), "0.00") & "|" & Format(Val(RS![DPQAmt] & ""), "##") & "|" & Format(PAMT, "0.00%") & ";"
End If
RS.MoveNext
Loop
 If Check_Option(96) = 1 Then
    F = "+5000|+>1500|+>1500|+>1500;"
    H = "Departments|Amount|Quantity|Percentage;"
    ReportView.View.TableBorder = tbBox
    ReportView.View.addTable F, H, Data, vbBlack, vbWhite
  End If
End If
DbClose Db, RS
Call UpdateWindow(Financal_Reports.hwnd)
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
  ReportView.View.EndDoc
  ReportView.Show 1
End If
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Summary Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub SSCommand28_Click()
End Sub

Private Sub SSCommand29_Click()
End Sub

Private Sub SSCommand30_Click()
End Sub

Private Sub SSCommand31_Click()

End Sub

Private Sub SSCommand32_Click()
End Sub

Private Sub SSCommand33_Click()
'If ProgramMenu = True Then
'    SetAccess SSCommand33.Caption
'    Exit Sub
'End If
If Can_Continue(SSCommand33.Caption) = False Then Exit Sub

Report_Name = "Rev"
CenterObject Me, Report_Type
Report_Type.ZOrder 0
Report_Type.Visible = True
Report_Type.Refresh
End Sub

Private Sub SSCommand34_Click()
'If ProgramMenu = True Then
'    SetAccess SSCommand34.Caption
'    Exit Sub
'End If
If Can_Continue(SSCommand34.Caption) = False Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    GroupRevenue SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If

End Sub
Public Sub SSCommand35_Click()
If Check_Option(199) = 1 Then
   See_Msg "Cannot Run Sales Reports From This Station", 5, "Option Disabled."
   Exit Sub
End If
'If ProgramMenu = True Then
'    SetAccess SSCommand35.Caption
'    Exit Sub
'End If
If Can_Continue(SSCommand35.Caption) = False Then Exit Sub
Dim TSold           As Currency
Dim QSold           As Integer
Dim CSold           As Currency
Dim Printall        As Integer
Dim pos             As Double
Dim Cdb             As ADODB.Connection
Dim Ctb             As ADODB.Recordset
Dim Heading         As String
Dim PrinterPort     As String
Dim PaidoutAmt      As Currency
Dim CashPayment     As Currency
Dim RecordNo        As Integer
Dim TSperc          As Double
Dim CatList         As New Collection
Dim catdb           As ADODB.Connection
Dim CatTb           As ADODB.Recordset
Dim coupons()       As Variant
Dim CashRefunds     As RefundAmts
Dim LL              As Integer
Dim NSR             As Currency
Dim TFeeIncome      As Currency
Dim H               As String
Dim F               As String
Dim Data            As String
Dim TotalCharges    As Currency
Dim NetSale         As Currency
Dim MealCount       As Integer
CSold = 0: TSperc = 0
PageHeader = "Daily Sales Recap"
SetPage ReportView.View, ""
'StartPage
With ReportView.View
   .CurrentY = "1.75in"
   .CurrentX = "1.75in"
   .CurrentY = .CurrentY - 100
   .FontSize = 12
   .textColor = vbBlack
   .MarginLeft = 2000
   .Text = "" & vbCrLf
   .Text = "" & vbCrLf
   .Text = "Daily Sales Recap For " & GetTodaysDate & vbCrLf
   
End With
On Error GoTo ErrRetry
    MealCount = ReturnDailyMealCount
    ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Cdb
    Sql = "Select Distinct [Department Report] from [Check Detail]"
    Sql = Sql & " Where [Deleted] = 0 and [Voided] = 0 and "
    'Sql = Sql & " [Department Report] <> 'Discount' and [Department Report] <> 'Coupon' and [Department Report] <> 'Frequent Diner' order by [Department Report] asc "

    Sql = Sql & " [Department Report] Not In ('Discount','Coupon','Online Tip','Frequent Diner','Online Fee','Online Delivery')"
    
    Sql = Sql & " order by [Department Report] asc"
    ConnectADORS Sql, Cdb, Ctb
       Do While Not Ctb.EOF
        On Error Resume Next
        If Ctb![Price] > 0 Then
           CatList.add CStr(Ctb![Department Report]), CStr(Ctb![Department Report])
        End If
        Ctb.MoveNext
       Loop
    DbClose Cdb, Ctb
ReDim Cat_Report(5, CatList.count)
For LL = 0 To CatList.count - 1
   Cat_Report(0, LL) = CatList.Item(LL + 1)
Next LL
RecordNo = CatList.count - 1
Set CatList(0) = Nothing
Dim TTL_Sold(0 To 2) As Variant
Sql = "Select Distinct [Menu Item],count([Menu Item]) as IQty,[Department Report], sum([Quantity]) as Qtysold,  sum([price]) as TotalSold"
Sql = Sql & " , sum([Cost Amount]) as TotalCost"
Sql = Sql & " , sum([Inclusive Amount]) as TotalInclusive"
Sql = Sql & " from [Check detail]"
Sql = Sql & " where [Voided] = 0 "
Sql = Sql & " and [close out day] = 0"
Sql = Sql & " group by [Menu Item], [Department Report]"
Sql = Sql & " Order by [Menu Item], [Department Report] asc"  '
    ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, catdb
    ConnectADORS Sql, catdb, CatTb
    Do While Not CatTb.EOF
       For GetCat = 0 To RecordNo
       If Trim(UCase(Cat_Report(0, GetCat))) = Trim(UCase(CatTb![Department Report])) Then
        If Val(CatTb!QtySold & "") = 0 Then
         Cat_Report(1, GetCat) = Cat_Report(1, GetCat) + Val(CatTb!Iqty & "")
         TTL_Sold(0) = TTL_Sold(0) + Val(CatTb!Iqty & "")
        Else
         Cat_Report(1, GetCat) = Cat_Report(1, GetCat) + Val(CatTb!QtySold & "")
         TTL_Sold(0) = TTL_Sold(0) + Val(Check_Null(CatTb!QtySold))
        End If
        Cat_Report(2, GetCat) = Cat_Report(2, GetCat) + Val(Check_Null(CatTb![Totalsold])) - Val(Check_Null(CatTb!TotalInclusive))
        Cat_Report(3, GetCat) = Cat_Report(3, GetCat) + Val(Check_Null(CatTb![totalcost]))
        
        TTL_Sold(1) = TTL_Sold(1) + Val(Check_Null(CatTb![Totalsold]))
        TTL_Sold(2) = TTL_Sold(2) + Val(Check_Null(CatTb![totalcost]))
        Exit For
       End If
       Next GetCat
      CatTb.MoveNext
    Loop
    DbClose catdb, CatTb

For Printall = 0 To RecordNo
    TSold = TSold + Val(Check_Null(Cat_Report(2, Printall)))
    QSold = QSold + Val(Check_Null(Cat_Report(1, Printall)))
    CSold = CSold + Val(Check_Null(Cat_Report(3, Printall)))
Next Printall
H = "Department Sales|Quantity|Dollars|Percent;"
F = "+<3500|^1200|+>1500|+>1500;"
Data = ""
For Printall = 0 To RecordNo
If Val(Cat_Report(1, Printall)) > 0 Then
      Data = Data & Cat_Report(0, Printall) & "|"
      Data = Data & Format(Cat_Report(1, Printall), "###") & "|"
      Data = Data & Format(Cat_Report(2, Printall), "0.00") & "|"
      If TSold > 0 And Val(Cat_Report(2, Printall)) > 0 Then
        pos = Cat_Report(2, Printall) / TSold
      Else
        pos = 0
      End If
      TSperc = TSperc + pos
      Data = Data & Format(pos * 100, "0.0000") & ";"
End If
'''''
Next Printall
Data = Data & "TOTAL " & "|"
Data = Data & "|"
Data = Data & Format(TSold, "0.00") & "|"
Data = Data & ";"

ReportView.View.TableBorder = tbBoxColumns
ReportView.View.addTable F, H, Data, vbGreen, vbWhite

TSold = 0
H = "Payments|Quantity|Dollars|Tips;"
F = "+<3500|^1200|+>1500|+>1500;"
Data = ""
'*************************
Dim ListRefunds As FinancialItems
Set ListRefunds = New FinancialItems
ListAllReturns ListRefunds, 0
'****************************
Dim Payments As FinancialItems
Set Payments = New FinancialItems
SortCreditCards Payments, 999, "", 0, "", ""   'was 0 changed to 999 for dropping delivery orders
TSold = 0: CSold = 0: TFeeIncome = 0
  For LL = 0 To UBound(CreditCards, 1) + 7
       Payments.Item(LL).Amount = Payments.Item(LL).Amount - Get_RefundAmts(ListRefunds, TotalRefund, UCase(Payments.Item(LL).Tender)) & ""
       Data = Data & Payments.Item(LL).Tender & "|"
       Data = Data & Payments.Item(LL).ItemQty & "|"
       Data = Data & Format(Payments.Item(LL).Amount, "0.00") & "|"
       Data = Data & Format(Payments.Item(LL).TipAmount, "0.00") & ";"
       TSold = TSold + Payments.Item(LL).Amount
       CSold = CSold + Payments.Item(LL).TipAmount
       If Payments.Item(LL).Tender = "Cash" Then CashPayment = Payments.Item(LL).Amount
       TFeeIncome = TFeeIncome + CalChargeBack(Payments.Item(LL).Tender, Payments.Item(LL).TipAmount)
    Next LL
Set ListRefunds = Nothing
Set Payments = Nothing
Data = Data & "TOTAL " & "|"
Data = Data & "|"
Data = Data & Format(TSold, "0.00") & "|"
TotalCharges = (TSold) - CashPayment
Data = Data & Format(CSold, "0.00") & ";"
ReportView.View.TableBorder = tbBoxColumns
ReportView.View.addTable F, H, Data, vbGreen, vbWhite
TSold = 0: CSold = 0
H = "Discounts|Quantity|Dollars;"
F = "+<3500|^1200|+>3000;"
Data = ""
Dim DiscountTotals As StoreDiscounts
Set DiscountTotals = New StoreDiscounts
Dim DeptTotals As StoreDiscounts
Set DeptTotals = New StoreDiscounts
Closing_Report.List_Coupons DiscountTotals, DeptTotals, False, 1, ""
  For LL = 1 To DiscountTotals.count
        Data = Data & DiscountTotals.Item(LL).DiscountName & "|"
        Data = Data & Format(DiscountTotals.Item(LL).DiscountQty, "###") & "|"
        Data = Data & Format(DiscountTotals.Item(LL).DiscountAmt, "0.00") & ";"
        TSold = TSold + DiscountTotals.Item(LL).DiscountAmt
 Next LL
Set DiscountTotals = Nothing
Set DeptTotals = Nothing
Data = Data & "TOTAL " & "|"
Data = Data & "|"
Data = Data & Format(TSold, "0.00") & ";"
ReportView.View.TableBorder = tbBoxColumns
ReportView.View.addTable F, H, Data, vbGreen, vbWhite
TSold = 0
Dim ExemptItems As FinancialItems
Set ExemptItems = New FinancialItems
SalesExemptItems ExemptItems, "", 0
If ExemptItems.count > 0 Then
 H = "Non Sales Revenue|Quantity|Dollars;"
 F = "+<3500|^1200|+>3000;"
 Data = ""
 TSold = 0
  For LL = 1 To ExemptItems.count
      Data = Data & ExemptItems.Item(LL).ItemName & "|"
      Data = Data & Format(ExemptItems.Item(LL).ItemQty, "###") & "|"
      Data = Data & Format(ExemptItems.Item(LL).Amount, "0.00") & ";"
      TSold = TSold + ExemptItems.Item(LL).Amount
 Next LL
 NSR = TSold

 Data = Data & "TOTAL " & "|"
 Data = Data & "|"
 Data = Data & Format(TSold, "0.00") & ";"
 ReportView.View.TableBorder = tbBoxColumns
 ReportView.View.addTable F, H, Data, vbGreen, vbWhite
End If
Set ExemptItems = Nothing
Dim Pdouts As SortAll
Set Pdouts = New SortAll
PaidoutAmt = GetPaidOuts("", Pdouts, 0, False, MyRegisterName, "", 0)
If PaidoutAmt > 0 Then
   H = "Paid Outs|Dollars;"
   F = "+<3500|+>1500;"
   Data = ""
   For LL = 1 To Pdouts.count
      Data = Data & Pdouts.Item(LL).MenuItem & "|"
      Data = Data & Format(Pdouts.Item(LL).SaleAmt, "0.00") & ";"
   Next LL
   Data = Data & "TOTAL " & "|"
   Data = Data & Format(PaidoutAmt, "0.00") & ";"
   ReportView.View.TableBorder = tbBoxColumns
   ReportView.View.addTable F, H, Data, vbGreen, vbWhite
End If
Set Pdouts = Nothing
H = "Sales|Dollars;"
F = "+<6200|+>1500;"
Data = ""
Sales_Breakdown True, 0
NetSale = GrossSales - Abs((NetDiscounts) + VoidSales + NetGratuity)
Data = Data & "Gross Sales  " & "|" & Format(GrossSales, "0.00") & ";"
Data = Data & "Voided Sales " & "|" & Format(VoidSales, "0.00") & ";"
Data = Data & "Sales Tax    " & "|" & Format(NetSTax, "0.00") & ";"
Data = Data & "Special Tax  " & "|" & Format(NetSpTax, "0.00") & ";"
Data = Data & SurChargeTxt & "|" & Format(NetLiqTax, "0.00") & ";"
Data = Data & "Non Sale Revenue" & "|" & Format(NSR, "0.00") & ";"
Data = Data & "Net Sales    " & "|" & Format((NetSale - NSR) - Abs(AmtOfRefunds), "0.00") & ";"
If AmtOfRefunds <> 0 Then PrintLine "Sales Refunded", 7632: PrintLine Format(AmtOfRefunds, "0.00"), -1
CashRefunds = Get_AuditRefund("", AllCash)
If CashRefunds.RefundAmt <> 0 Then
 Data = Data & "Cash Refunded " & "|" & Format(CashRefunds.RefundAmt + CashRefunds.RefundTax, "0.00") & ";"
End If
CashtoAcctFor = (GrossSales + ((Abs(NetSTax) + (Abs(NetSpTax)) + TFeeIncome) + (Abs(NetLiqTax)))) - (Abs(NetDiscounts) + VoidSales) - NetGratuity

If Input_Option(58) = "Yes" Then CashtoAcctFor = CashtoAcctFor + NetGratuity
CashtoAcctFor = CashtoAcctFor - Abs(AmtOfRefunds)
If Check_Option(237) = 1 Then CashtoAcctFor = CashtoAcctFor + CSold

Data = Data & "TOTAL TO ACCOUNT FOR " & "|" & Format(CashtoAcctFor, "0.00") & ";"
If Check_Option(237) = 1 Then
 Data = Data & "CASH TO ACCOUNT FOR              " & "|" & Format(((CashtoAcctFor)) - (TotalCharges + PaidoutAmt), "0.00") & ";"
Else
 Data = Data & "CASH TO ACCOUNT FOR (less tips)  " & "|" & Format(((CashtoAcctFor)) - (TotalCharges + PaidoutAmt), "0.00") & ";"
End If
If Input_Option(58) = "Yes" Then
 Data = Data & "gratuity included in total " & "|" & ";"
Else
 Data = Data & "gratuity not included in total " & "|;"
End If
If MealCount > 0 Then
 Data = Data & "Meal Count" & "|" & MealCount & ";"
 Data = Data & "Average" & "|" & Format(NetSale / MealCount, "0.00") & ";"
End If
ReportView.View.TableBorder = tbBoxColumns
ReportView.View.addTable F, H, Data, vbGreen, vbWhite
FinishJob 1
Exit Sub
ErrRetry:
  Error_Log Err.Number, Err.Description, "Load Recap", "None"
  Resume Next
End Sub
Private Sub SSCommand36_Click()
End Sub
Private Sub SSCommand37_Click()
If Can_Continue(SSCommand37.Caption) = False Then Exit Sub
PaymentbyHour.Visible = True
With List1
   .Clear
   .AddItem "Today"
   .AddItem "Monday"
   .AddItem "Tuesday"
   .AddItem "Wednesday"
   .AddItem "Thursday"
   .AddItem "Friday"
   .AddItem "Saturday"
   .AddItem "Sunday"
   .AddItem "Week To Date"
End With
With PickList
   CenterObject Me, PickList
   .Visible = True
   .ZOrder 0
End With
End Sub
Private Sub SSCommand38_Click()
Report_Name = ""
PaymentbyHour.Visible = False
With PickList
  .Visible = False
  .Left = -5000
End With
End Sub
Private Sub SSCommand4_Click()
End Sub
Private Sub SSCommand2_Click()
If Can_Continue(SSCommand2.Caption) = False Then Exit Sub
Sales_Mix.Show
End Sub
Private Sub SSCommand3_Click()
If Can_Continue(SSCommand3.Caption) = False Then Exit Sub
Time_Reports.Show
End Sub
Private Sub SSCommand42_Click()
SSCommand42.Caption = "Working....."
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    MealPeriodReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
SSCommand42.Caption = "Meal Period Report"
SSCommand42.Refresh
End Sub
Private Sub SSCommand5_Click()
Dim Cdb         As ADODB.Connection
Dim Ctb         As ADODB.Recordset
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim BeginDate   As String
Dim EndingDate  As String
Dim ResetReg    As Integer
Dim Destination As String

On Error GoTo errorhandler
If CloseWeek = True Then GoTo Cw  'if auto close is set will auto close week
If Can_Continue(SSCommand5.Caption) = False Then Exit Sub
If WeekIsClosed = True Then
  See_Msg "Week Closed", 5, "No Sales Found"
  Exit Sub
End If
'MkDir "C:\SqlBackUp\"
If Dir(MyRegisters(0).RegLoc & "Weekly Backup", vbDirectory) = "" Then
 MkDir MyRegisters(0).Svr & "Weekly BackUp"
End If
On Error GoTo errorhandler
'auto close week will only close closed records leaves all close out day records
'that are set to 0 open.
    If See_Msg("Weekly Close out Resets all Databases !!", 4, "Are You Sure ?") <> 6 Then Exit Sub
Cw:
        If FileCheck(MyRegisters(0).Svr & "Chk-Stat.mdb") = False Then
         If FileCheck(MyRegisters(0).Svr & "Chk-Stat.BAK") Then
          FileCopy MyRegisters(0).Svr & "Chk-Stat.BAK", MyRegisters(0).Svr & "Chk-Stat.mdb"
         End If
        End If
        '// makes sure we have a file to create a backup with
        '// Oliver had a site with not .mdb only bak files -- stop the weekly close for serveral weeks until we figured it out.
        '// 02/02/2025
        
        If CloseWeek = False Then
         Sale_Mix CloseMixSales, SelectAll, "", "", "", "None", 0
        End If
        ResetLogFiles
        ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Db
         Sql = "Select * from [Closed Checks] order by [order Date] asc"
        ConnectADORS Sql, Db, Tb
          If Not Tb.EOF Then
             BeginDate = Tb(19)
          End If
          Tb.Close
          Sql = "Select * from [Closed Checks] order by [order Date] Desc"
          ConnectADORS Sql, Db, Tb
                If Not Tb.EOF Then
                  EndingDate = Tb(19)
                End If
                DbClose Db, Tb
                Dim TotalDayInDB As Integer
                TotalDayInDB = DateDiff("d", BeginDate, EndingDate)
                If Trim(BeginDate) = "" Then BeginDate = Format(Date, "mm/dd/yyyy")
                If Not IsDate(BeginDate) Then BeginDate = Format(Date, "mm/dd/yyyy")
                Destination = Format(BeginDate, "MM-DD-YYYY") & " Week " & MyRegisters(0).regName & " Sales BackUp.MDB"
                  If FileCheck(MyRegisters(0).Svr & "Chk-Stat.mdb") Then
                   FileCopy MyRegisters(0).Svr & "Chk-Stat.mdb", MyRegisters(0).Svr & "Weekly BackUp\" & Destination
                  End If
                  WeeklyBackUp BeginDate, EndingDate, Destination, 0
                  ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Db
                  Sql = "Delete from [Check Detail] Where [Close Out Day] > 0"
                  Db.Execute Sql
                  Sql = "Delete from [Closed Checks] where [Close Out Day] > 0"
                  Db.Execute Sql
                  Sql = "Delete from [Payments] where [Close Out Day] > 0"
                  Db.Execute Sql
                  Sql = "Delete from [Error Correct] where [Close Out Day] > 0"
                  Db.Execute Sql
                  Sql = "Delete from [Item Refund] where [Close Out Day] > 0"
                  Db.Execute Sql
                  SetNothing Db
          If SvrPath.TerminalBkPath > "" Then
           If FileCheck(SvrPath.TerminalBkPath & CheckStatDb) = True Then
             ConnectADODB SvrPath.TerminalBkPath & CheckStatDb, CheckStatDb, Db
              Sql = "Delete from [Check Detail] Where [Close Out Day] > 0"""
              Db.Execute Sql
              Sql = "Delete from [Closed Checks] where [Close Out Day] > 0"
              Db.Execute Sql
              Sql = "Delete from [Payments] where [Close Out Day] > 0"
              Db.Execute Sql
              Sql = "Delete from [Error Correct] where [Close Out Day] > 0"
              Db.Execute Sql
             SetNothing Db
           End If
          End If
         Dim RM As GSubs
         Set RM = New GSubs
         RM.ResetRefunds
         Set RM = Nothing
                 
        If RemoteBackup_Path = "" Then RemoteBackup_Path = MyRegisters(0).Svr & "Weekly BackUp\"
        If TimeRecordsDb = "TimeRecordsSQL" Then
          Destination = Format(Date, "MM-DD-YYYY") & " Week " & "TimeRecordsBackUp.MDB"
          FileCopy MyRegisters(0).Svr & "Time Records.mdb", MyRegisters(0).Svr & "Weekly BackUp\" & Destination
          SqlToAccessBackUp MyRegisters(0).Svr & "Weekly BackUp\" & Destination, Destination, MyRegisters(0).Svr & TimeRecordsDb, TimeRecordsDb
          'PurgeTimeFiles
          On Error Resume Next
          If MyRegisters(0).Svr & "Weekly BackUp\" <> RemoteBackup_Path Then
           FileCopy MyRegisters(0).Svr & "Weekly BackUp\" & Destination, RemoteBackup_Path & Destination
          End If
        End If
        If HouseAcct = "House Accounts" Then
          Destination = Format(BeginDate, "MM-DD-YYYY") & " Week " & "HouseAcctBackUp.MDB"
          FileCopy MyRegisters(0).Svr & "Scan.mdb", MyRegisters(0).Svr & "Weekly BackUp\" & Destination
          SqlToAccessBackUp MyRegisters(0).Svr & "Weekly BackUp\" & Destination, Destination, MyRegisters(0).Svr & HouseAcct, "House Accounts"
          On Error Resume Next
          If MyRegisters(0).Svr & "Weekly BackUp\" <> RemoteBackup_Path Then
           FileCopy MyRegisters(0).Svr & "Weekly BackUp\" & Destination, RemoteBackup_Path & Destination
          End If
        End If

        If FrequentDinerDb = "Frequent DinerSql" Then
          Destination = Format(BeginDate, "MM-DD-YYYY") & " Week " & "FrequentDinerBackUp.MDB"
          FileCopy MyRegisters(0).Svr & "Frequent Diner.MDb", MyRegisters(0).Svr & Destination
          SqlToAccessBackUp MyRegisters(0).Svr & Destination, Destination, MyRegisters(0).Svr & FrequentDinerDb, "Frequent DinerSql"
        End If
        On Error GoTo errorhandler
        ResetFrequentDiner
        RemoveMenuLink
        SaveRemoteData MyRegisters(0).Svr, RemoteBackup_Path
        If SendEmails(5).Send Then
        Display.DSRBalanceSheet CStr(CDate(BeginDate)), CStr(CDate(BeginDate) + 6), True, 0
        Unload Display
        SendEMailEOD
        End If
    'resets all 0 balance gift certificates keeps file cleaned up once a week.
SkipBackUP:
On Error GoTo errorhandler
CleanDir
AutoPurge
Trans_Log "Weekly Close Complete...."
See_Msg "Weekly Update Complete", 5, ""
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Weekly Closing ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand6_Click()
If Can_Continue(SSCommand6.Caption) = False Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", MyRegisters(0).Svr & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
   If GetConfigInfo("Discounts", "Coupons Prompt for Zip Code", "No", "Functions.INI") = "ENABLED" Then
     If See_Msg("Coupon Report By Coupon Code", 4, "Coupons By Codes") = 6 Then
        ZipCouponReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
     Else
        CouponReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
     End If
   Else
    CouponReport SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
   End If
    Set SearchDates = Nothing
End If
End Sub
Public Sub CouponReport(StartDate As String, EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort As String
Dim ReportType As String
Dim ListDisc() As DiscValues
ReportType = "Coupon Report" & vbCrLf & " From " & StartDate & " to " & EndDate
'On Error Resume Next
Dim Db As ADODB.Connection
Dim CC As ADODB.Recordset
Dim RTS As Integer
Dim FDL As Integer
Dim CServer As String
Dim DTotal  As Currency
ReDim ListDisc(0)
RTS = -1
  ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Db
    Sql = "Select * From [Check Detail]"
    Sql = Sql & " Where Voided = 0"
    Sql = Sql & " and [Price] < 0"
    Sql = Sql & " and [Department Report] <> 'Non Sales Discount'    "
    If SvrPath.sServerIsOn = 1 Then
      Sql = Sql & " and TRY_CAST([Check Date] AS datetime) + TRY_CAST([recorded time] AS datetime) between '" & CDate(StartDate & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(EndDate) & " " & SetEndingTime) & "'"
    Else
      Sql = Sql & " and cdate([Check Date] & ' ' & [recorded time]) between #" & CDate(StartDate & " " & GetStartTime) & "# and #" & CDate(SetEndingDate(EndDate) & " " & SetEndingTime) & "#"
    End If
    Sql = Sql & " Order by [Servers Name],[menu item] asc"
  ConnectADORS Sql, Db, CC
     If CC.EOF Then
       ' See_Msg "No Discounts Found", 0, ""
     Else
       ReDim Preserve ListDisc(RTS + (RecordsCount(CC) + 1))
     End If
     Do While Not CC.EOF
        RTS = RTS + 1
        ListDisc(RTS).DItem = CC![Menu Item] & ""
        ListDisc(RTS).DAmt = Abs(Val(CC![Price] & "")) - Abs(Val(CC![Inclusive Amount] & ""))
        ListDisc(RTS).DServer = CC![Servers Name] & ""
        ListDisc(RTS).DCheck = Val(CC![Check Number] & "")
        ListDisc(RTS).DDate = CC![Check Date] & ""
        ListDisc(RTS).DTime = CC![Recorded Time] & ""
        ListDisc(RTS).DMGRName = CC![Voided By] & ""
     CC.MoveNext
   Loop
DbClose Db, CC
If RTS = -1 Then
   See_Msg "No Discounts Found", 0, ""
   Exit Sub
End If
    If Check_Option(96) = 0 Then
        Printer_Drivers "Receipt"
        PF = FreeFile
        PrinterPort = "Coupon Sales"
        Close
        Open PrinterPort For Output Shared As #PF
        CenterText = EC(8)
        LeftText = EC(9)
        'RightText = Chr(27) + Chr(97) + Chr(2)
        Print #PF, CenterText; ReportType; EC(7)
        Print #PF, LeftText & EC(7)
    Else
        StartPage
        PrintLine "", -1
        PrintCLine ReportType, -1, 0
        PrintLine "", -1
    End If
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Date    Check #    Name              Amt"; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Date ", 2000: PrintLine "Check Number ", 4000: PrintLine "Discount Name ", 7500: PrintLine "Amount", 9500: PrintLine "Manager", -1
        DrawLine
    End If
    CServer = ""
    DTotal = 0
    For FDL = 0 To RTS
       If CServer <> ListDisc(FDL).DServer Then
        If Check_Option(96) = 0 Then
            Print #PF, EC(1); ListDisc(FDL).DServer; EC(4); EC(7)
            Print #PF, String(40, 45); EC(7)
        Else
            PrintLine "Server Name " & ListDisc(FDL).DServer, -1
            DrawLine
        End If
        CServer = ListDisc(FDL).DServer
       End If
       If Check_Option(96) = 0 Then
        Print #PF, Format(ListDisc(FDL).DDate, "m/d"); Tab(6); ListDisc(FDL).DCheck;
        If ListDisc(FDL).DMGRName > "" Then
          Print #PF, Tab(15); ListDisc(FDL).DItem; Tab(36); Format(ListDisc(FDL).DAmt, CurrencyPreFix); EC(7)
          Print #PF, "Approved By " & ListDisc(FDL).DMGRName; EC(7)
        Else
          Print #PF, Tab(15); ListDisc(FDL).DItem; Tab(36); Format(ListDisc(FDL).DAmt, CurrencyPreFix); EC(7)
        End If
      Else
        PrintLine ListDisc(FDL).DDate, 2000: PrintLine ListDisc(FDL).DCheck, 4000:
        PrintLine ListDisc(FDL).DItem, 7500: PrintLine Format(ListDisc(FDL).DAmt, CurrencyPreFix), 9500
        PrintLine ListDisc(FDL).DMGRName, -1
      End If
      DTotal = DTotal + ListDisc(FDL).DAmt
    Next FDL
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
    End If
If Check_Option(96) = 0 Then
    Print #PF, "Total Qty / $ "; Tab(25); Format(RTS + 1, "0"); Tab(36); Format(DTotal, CurrencyPreFix); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    PrintLine "Total ", 4000
    PrintLine "Quantity " & Format(RTS + 1, "###"), 7500
    PrintLine Format(DTotal, CurrencyPreFix), -1
    DrawLine
    ReportView.View.EndDoc
    ReportView.Caption = "Coupon Report"
    ReportView.Show 1
End If
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Coupon Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub ZipCouponReport(StartDate As String, EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort As String
Dim ReportType As String
Dim ListDisc() As DiscValues
ReportType = "Coupon Report by Zip Code" & vbCrLf & " From " & StartDate & " to " & EndDate
'On Error Resume Next
Dim Db As ADODB.Connection
Dim CC As ADODB.Recordset
Dim RTS As Integer
Dim FDL As Integer
Dim CServer As String
Dim DTotal  As Currency
ReDim ListDisc(0)
Dim SRC As Integer
Dim IFound As Boolean
Dim TQty   As Integer
RTS = -1
TQty = 0
  ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Db
    Sql = "Select Distinct [Disc Type],[Menu Item],Sum([Price]) as tprice, Count([Menu Item]) as qty From [Check Detail]"
    Sql = Sql & " Where Voided = 0"
    Sql = Sql & " and [Price] < 0"
    Sql = Sql & " and [Cateory Name] = 'Coupon'"
    If SvrPath.sServerIsOn = 1 Then
      Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & CDate(StartDate) & "' and '" & CDate(EndDate) & "'"
    Else
      Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    End If
    Sql = Sql & " Group by [Disc type],[Menu Item]"
  ConnectADORS Sql, Db, CC
     
     
     Do While Not CC.EOF
      For SRC = 0 To RTS
       IFound = False
       If ListDisc(SRC).DItem = CC![Menu Item] & "" And ListDisc(SRC).ZipCode = CC![Disc Type] & "" Then
        ListDisc(SRC).DAmt = ListDisc(SRC).DAmt + Val(CC![TPrice] & "")
        ListDisc(SRC).DCheck = ListDisc(SRC).DCheck = Val(CC![qty] & "")
        IFound = True
       End If
      Next SRC
      If IFound = False Then
        RTS = RTS + 1
        ReDim Preserve ListDisc(RTS)
        ListDisc(RTS).DItem = CC![Menu Item] & ""
        ListDisc(RTS).DAmt = Val(CC![TPrice] & "")
        ListDisc(RTS).DCheck = Val(CC![qty] & "")
        ListDisc(RTS).ZipCode = CC![Disc Type] & ""
        TQty = TQty + Val(CC![qty] & "")
      End If
     CC.MoveNext
   Loop
DbClose Db, CC
If RTS = -1 Then
   See_Msg "No Coupons Found", 0, ""
   Exit Sub
End If

    If Check_Option(96) = 0 Then
        Printer_Drivers "Receipt"
        PF = FreeFile
        PrinterPort = "Coupon Sales"
        Close
        Open PrinterPort For Output Shared As #PF
        CenterText = EC(8)
        LeftText = EC(9)
        'RightText = Chr(27) + Chr(97) + Chr(2)
        Print #PF, CenterText; ReportType; EC(7)
        Print #PF, LeftText & EC(7)
    Else
        StartPage
        PrintLine "", -1
        PrintCLine ReportType, -1, 0
        PrintLine "", -1
    End If
    If Check_Option(96) = 0 Then
        Print #PF, EC(2); String(40, 45); EC(7)
        Print #PF, "Name             Quantity         Amt"; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
        PrintLine "Discount Name ", 2500: PrintLine "Quantity", 4000:  PrintLine "Amount", -1
        DrawLine
    End If
    CServer = ""
    DTotal = 0
    For FDL = 0 To RTS
       If CServer <> ListDisc(FDL).ZipCode Then
        If Check_Option(96) = 0 Then
            Print #PF, EC(1); "Zip Code " & ListDisc(FDL).ZipCode; EC(7)
            Print #PF, EC(2); String(40, 45); EC(7)
        Else
            PrintLine "Zip Code " & ListDisc(FDL).ZipCode, -1
            DrawLine
        End If
        CServer = ListDisc(FDL).ZipCode
       End If
       If Check_Option(96) = 0 Then
        Print #PF, ListDisc(FDL).DItem; Tab(22); ListDisc(FDL).DCheck;
        Print #PF, Tab(36); Format(Abs(ListDisc(FDL).DAmt), CurrencyPreFix); EC(7)
      Else
        PrintLine ListDisc(FDL).DItem, 2600: PrintLine ListDisc(FDL).DCheck, 4000
        PrintLine Format(Abs(ListDisc(FDL).DAmt), CurrencyPreFix), -1
      End If
      DTotal = DTotal + ListDisc(FDL).DAmt
    Next FDL
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
    End If
If Check_Option(96) = 0 Then
    Print #PF, "Total Qty / $ "; Tab(23); Format(TQty, "0"); Tab(36); Format(Abs(DTotal), CurrencyPreFix); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    PrintLine "Total ", 2600
    PrintLine Format(TQty, "###"), 4000
    PrintLine Format(Abs(DTotal), CurrencyPreFix), -1
    
    DrawLine
    ReportView.View.EndDoc
    ReportView.Caption = "Coupon Report By Zip"
    ReportView.Show 1
End If
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Zip Coupon Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Public Sub VoidItemReport(StartDate As String, EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort     As String
Dim ReportType      As String
Dim ListDisc()      As DiscValues
Dim Db              As ADODB.Connection
Dim CC              As ADODB.Recordset
Dim RTS             As Integer
Dim FDL             As Integer
Dim CServer         As String
Dim DTotal          As Currency
Dim H               As String
Dim F               As String
Dim Data            As String
ReDim ListDisc(0)
ReportType = "Void Report" & vbCrLf & " From " & StartDate & " to " & EndDate

RTS = -1
  ConnectADODB MyRegisters(0).Svr & CheckStatDb, CheckStatDb, Db
   If Check_Option(244) = 1 Then
      Sql = "Select * From [Check Detail]"
      Sql = Sql & " Where [Voided] > 0 and [Price] > 0"
      If See_Msg("Show Voids Only", 4, "Filter Out Deletes") = 6 Then
        Sql = Sql & " and [Void Reason] <> 'Cashier Correct'"
        Sql = Sql & " and [Void Reason] <> 'Manager Correct'"
        Sql = Sql & " and [Void Reason] <> 'Cancel Order'"
      End If
      'If SvrPath.sServerIsOn = 1 Then
       Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & CDate(StartDate) & "' and '" & CDate(EndDate) & "'"
       Sql = Sql & " Order by [Servers Name],TRY_CAST([Check date] AS datetime),[Check Number],[menu item] asc"
      'Else
      ' Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
      ' Sql = Sql & " Order by [Servers Name], TRY_CAST([Check date] AS datetime),[Check Number],[menu item] asc"
      'End If
   Else
    Sql = "Select * From [Check Detail]"
    Sql = Sql & " Where [Voided] > 0 and [Price] > 0"
    If SvrPath.sServerIsOn = 1 Then
      Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & CDate(StartDate) & "' and '" & CDate(EndDate) & "'"
      Sql = Sql & " Order by [Servers Name],TRY_CAST([Check date] AS datetime),[Check Number],[menu item] asc"
    Else
      Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & Format(CDate(CDate(StartDate)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(EndDate)), "yyyy-mm-dd hh:nn:ss") & "'"
      Sql = Sql & " Order by [Servers Name], TRY_CAST([Check date] AS datetime),[Check Number],[menu item] asc"
    End If
   End If
  ConnectADORS Sql, Db, CC
     If CC.EOF Then
       ' See_Msg "No Discounts Found", 0, ""
     Else
       ReDim Preserve ListDisc(RTS + (RecordsCount(CC) + 1))
     End If
     Do While Not CC.EOF
        RTS = RTS + 1
        ListDisc(RTS).DItem = CC![Menu Item] & ""
        ListDisc(RTS).DAmt = Val(CC![Price] & "")
        ListDisc(RTS).DServer = CC![Servers Name] & ""
        ListDisc(RTS).DCheck = Val(CC![Check Number] & "")
        ListDisc(RTS).DDate = CC![Check Date] & ""
        ListDisc(RTS).DTime = CC![Void Time] & ""
        ListDisc(RTS).DMGRName = CC![Voided By] & ""
        ListDisc(RTS).ZipCode = CC![Void Reason] & ""
     CC.MoveNext
   Loop
  DbClose Db, CC
  If RTS = -1 Then
    See_Msg "No Voids Found", 5, ""
    Exit Sub
  End If
    If Check_Option(96) = 0 Then
        Printer_Drivers "Receipt"
        PF = FreeFile
        PrinterPort = "Coupon Sales"
        Close
        Open PrinterPort For Output Shared As #PF
        CenterText = EC(8)
        LeftText = EC(9)
        'RightText = Chr(27) + Chr(97) + Chr(2)
        Print #PF, CenterText; ReportType; EC(7)
        Print #PF, LeftText & EC(7)
    Else
        SetPage ReportView.View, ReportType
    End If
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Date    Check #    Name              Amt"; EC(7)
        Print #PF, String(40, 45); EC(7)
    Else
        H = "Date|Check No|Item Name|Reason|Amount|Voided By|Void Time;"
        F = "+<1000|^1000|+<2500|+<3000|+>1000|+<2000|+<1000;"
    End If
    CServer = ""
    DTotal = 0
    Data = ""
    Dim PrintData As Boolean
    PrintData = False
    For FDL = 0 To RTS
       If CServer <> ListDisc(FDL).DServer Then
        If Check_Option(96) = 0 Then
           Print #PF, EC(1); ListDisc(FDL).DServer; EC(4); EC(7)
           Print #PF, String(40, 45); EC(7)
        Else
           If PrintData Then
             ReportView.View.TableBorder = tbBoxColumns
             ReportView.View.addTable F, H, Data, vbBlack, vbWhite
           End If
           ReportView.View.TableBorder = tbBox
           ReportView.View.addTable "+<4500;", "", "Name " & ListDisc(FDL).DServer & ";", vbBlack, vbWhite
           PrintData = True
           Data = ""
        End If
        CServer = ListDisc(FDL).DServer
       End If
       If Check_Option(96) = 0 Then
        Print #PF, Format(ListDisc(FDL).DDate, "m/d"); Tab(6); ListDisc(FDL).DCheck;
        Print #PF, Tab(15); ListDisc(FDL).DItem; Tab(36); Format(ListDisc(FDL).DAmt, CurrencyPreFix); EC(7)
      Else
        Data = Data & ListDisc(FDL).DDate & "|" & ListDisc(FDL).DCheck & "|"
        Data = Data & ListDisc(FDL).DItem & "|" & ListDisc(FDL).ZipCode & "|" & Format(ListDisc(FDL).DAmt, CurrencyPreFix) & "|"
        Data = Data & ListDisc(FDL).DMGRName & "|" & ListDisc(FDL).DTime & ";"
      End If
      DTotal = DTotal + ListDisc(FDL).DAmt
    Next FDL
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
    Else
      ReportView.View.TableBorder = tbBoxColumns
      ReportView.View.addTable F, H, Data, vbBlack, vbWhite
      Data = ""
    End If
If Check_Option(96) = 0 Then
    Print #PF, "Total Qty / $ "; Tab(25); Format(RTS + 1, "0"); Tab(36); Format(DTotal, CurrencyPreFix); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    Data = ""
    H = ""
    F = "+<5500|^1500|+>1500;"
    Data = Data & "Totals " & "|"
    Data = Data & "Quantity " & Format(RTS + 1, "###") & "|"
    Data = Data & Format(DTotal, CurrencyPreFix) & ";"
    ReportView.View.TableBorder = tbBoxColumns
    ReportView.View.addTable F, H, Data, vbBlack, vbWhite
    ReportView.View.EndDoc
    ReportView.Caption = "Void Item Report"
    ReportView.Show 1
End If
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Void Item Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Public Sub CouponHistoryReport(StartDate As String, EndDate As String)
On Error GoTo errorhandler
Dim PrinterPort As String
Dim ReportType As String
Dim ListDisc() As DiscValues
ReportType = "Coupon History Report" & vbCrLf & " From " & StartDate & " to " & EndDate
'On Error Resume Next
Dim Db As ADODB.Connection
Dim CC As ADODB.Recordset
Dim RTS As Long
Dim FDL As Long
Dim CServer As String
Dim DTotal  As Currency
Dim Reg As Integer
Dim STotal As Currency
RTS = -1
  ConnectADODB MyRegisters(0).Svr & ProductHistory, ProductHistory, Db
    Sql = "Select * From [Detail]"
    Sql = Sql & " Where [Item Price] < 0"
    Sql = Sql & " and TRY_CAST([Date] AS datetime) Between #" & CDate(StartDate & " " & GetStartTime) & "#"
    Sql = Sql & " and #" & CDate(SetEndingDate(EndDate) & " " & SetEndingTime) & "#"
    If SvrPath.sServerIsOn = 1 And SvrPath.sTerminalAddress > "" Then
      Sql = Sql & " Order by [Server Name],[menu item],[Date] asc"
    Else
      Sql = Sql & " Order by [Server Name],[menu item],TRY_CAST([Date] AS datetime) asc"
    End If
  ConnectADORS ReturnDateString(Sql), Db, CC
     If CC.EOF Then
       ' See_Msg "No Discounts Found", 0, ""
     Else
       ReDim Preserve ListDisc(RecordsCount(CC))
     End If
     Do While Not CC.EOF
        RTS = RTS + 1
        ListDisc(RTS).DItem = CC![Menu Item] & ""
        ListDisc(RTS).DAmt = Abs(Val(CC![Item Price] & ""))
        ListDisc(RTS).DServer = CC![Server Name] & ""
        ListDisc(RTS).DDate = Format(CC![Date] & "", "MM/DD/YYYY")
        ListDisc(RTS).DTime = Format(CC![Time] & "", "HH:MM ampm")
        'ListDisc(RTS).DMGRName = CC![Manager] & ""
     CC.MoveNext
   Loop
DbClose Db, CC
If RTS = -1 Then
   See_Msg "No Discounts Found", 0, ""
   Exit Sub
End If
    If Check_Option(96) = 0 Then
        Printer_Drivers "Receipt"
        PF = FreeFile
        PrinterPort = "Coupon Sales"
        Close
        Open PrinterPort For Output Shared As #PF
        CenterText = EC(8)
        LeftText = EC(9)
        'RightText = Chr(27) + Chr(97) + Chr(2)
        Print #PF, CenterText; ReportType; EC(7)
        Print #PF, LeftText & EC(7)
    Else
        StartPage
        PrintLine "", -1
        PrintCLine ReportType, -1, 0
        PrintLine "", -1
    End If
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Date    Time       Name              Amt"; EC(7)
    Else
        DrawLine
        PrintLine "Date ", 2000: PrintLine "Time ", 4000: PrintLine "Discount Name ", 7500: PrintLine "Amount ", -1 ' 9500: PrintLine "Mgr", -1
    End If
    CServer = ""
    DTotal = 0
    For FDL = 0 To RTS
       If CServer <> ListDisc(FDL).DServer Then
       If STotal <> 0 Then
        If Check_Option(96) = 0 Then
            Print #PF, String(40, 45); EC(7)
            Print #PF, EC(1); "Discount Total " & Format(STotal, "0.00"); EC(4); EC(7)
            Print #PF, String(40, 45); EC(7)
        Else
            DrawLine
            PrintLine "", 7500
            PrintLine "Discount Total --> " & Format(STotal, "0.00"), -1
            DrawLine
        End If
       End If
       STotal = 0
        If Check_Option(96) = 0 Then
            Print #PF, String(40, 45); EC(7)
            Print #PF, EC(1); ListDisc(FDL).DServer; EC(4); EC(7)
            Print #PF, String(40, 45); EC(7)
        Else
            DrawLine
            PrintLine "Server Name --> " & ListDisc(FDL).DServer, -1
            DrawLine
        End If
        CServer = ListDisc(FDL).DServer
       End If
       If Check_Option(96) = 0 Then
        Print #PF, Format(ListDisc(FDL).DDate, "m/d"); Tab(6); Format(ListDisc(FDL).DTime, "H:mm ampm");
        Print #PF, Tab(16); ListDisc(FDL).DItem; Tab(36); Format(ListDisc(FDL).DAmt, CurrencyPreFix); EC(7)
       Else
        PrintLine ListDisc(FDL).DDate, 2000: PrintLine Format(ListDisc(FDL).DTime, "H:mm ampm"), 4000:
        PrintLine ListDisc(FDL).DItem, 7500: PrintLine Format(ListDisc(FDL).DAmt, CurrencyPreFix), -1 ' 9500:
        'PrintLine ListDisc(FDL).DMGRName, -1
       End If
       DTotal = DTotal + ListDisc(FDL).DAmt
       STotal = STotal + ListDisc(FDL).DAmt
    Next FDL
    If Check_Option(96) = 0 Then
        Print #PF, String(40, 45); EC(7)
    Else
        DrawLine
    End If
If Check_Option(96) = 0 Then
    Print #PF, "Total Qty / $ "; Tab(25); Format(RTS + 1, "0"); Tab(36); Format(DTotal, CurrencyPreFix); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    PrintLine "Total ", 4000
    PrintLine "Quantity " & Format(RTS + 1, "###"), 7500
    PrintLine Format(DTotal, CurrencyPreFix), -1
    DrawLine
    ReportView.View.EndDoc
    ReportView.Caption = "Coupon History"
    ReportView.Show 1
End If
Exit Sub
errorhandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Coupon History Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand7_Click()
Report_Type.Visible = False
Report_Type.Left = -3400
End Sub
Private Sub SSCommand8_Click()
If Check_Option(199) = 1 Then
   See_Msg "Cannot Run Sales Reports From This Station", 5, "Option Disabled."
   Exit Sub
End If
If Can_Continue(SSCommand8.Caption) = False Then Exit Sub
vbHold = SSCommand8.Caption
SSCommand8.Caption = "Loading Please Wait...."
Red_Book.Show 1
SSCommand8.Caption = vbHold
End Sub
Private Sub SSCommand9_Click()
Dim ScreenReport As Boolean
SSCommand9.Caption = "Working..."
SSCommand9.Refresh
If Report_Name = "Dept" Then
  Dept_Report TaxedSales.Text, "WTD Ending " & Format(Date, "M/D/YY"), 0, "WTD", "", 0, True
ElseIf Report_Name = "Rev" Then
TableRev False, True
ElseIf Report_Name = "Term" Then
TerminalReport False, True
Else
  ScreenReport = True
  If Check_Option(96) = 0 Then
     If See_Msg("Print Report", 4, "") = 6 Then
        ScreenReport = False
     End If
  End If
Void_Report "WTD", "", ScreenReport, True
End If
Report_Type.Visible = False
SSCommand9.Caption = "Week To Date Report"
SSCommand9.Refresh
SSCommand7_Click
'RevGroup
End Sub


