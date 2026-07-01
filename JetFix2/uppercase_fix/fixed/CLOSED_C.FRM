VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Closed_Check 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Review"
   ClientHeight    =   10620
   ClientLeft      =   135
   ClientTop       =   750
   ClientWidth     =   13935
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CLOSED_C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10620
   ScaleWidth      =   13935
   Begin VB.PictureBox ViewChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00CACABC&
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
      Height          =   8775
      Left            =   -4080
      ScaleHeight     =   8745
      ScaleWidth      =   6465
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Next Check"
         Height          =   570
         Left            =   4920
         TabIndex        =   45
         Top             =   7920
         Width           =   1455
      End
      Begin VSFlex8DAOCtl.VSFlexGrid Ticket 
         Height          =   5415
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   2040
         Width           =   4335
         _cx             =   1978015198
         _cy             =   1978017103
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14211288
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
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   500
         Cols            =   47
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
      Begin BTNENHLib4.BtnEnh SSCommand13 
         Height          =   735
         Left            =   4920
         TabIndex        =   30
         Top             =   7080
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
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
         BackColorContainer=   8421504
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "CLOSED_C.frx":FC8F
         textLT          =   "CLOSED_C.frx":FCF7
         textCT          =   "CLOSED_C.frx":FD0F
         textRT          =   "CLOSED_C.frx":FD27
         textLM          =   "CLOSED_C.frx":FD3F
         textRM          =   "CLOSED_C.frx":FD57
         textLB          =   "CLOSED_C.frx":FD6F
         textCB          =   "CLOSED_C.frx":FD87
         textRB          =   "CLOSED_C.frx":FD9F
         colorBack       =   "CLOSED_C.frx":FDB7
         colorIntern     =   "CLOSED_C.frx":FDE1
         colorMO         =   "CLOSED_C.frx":FE0B
         colorFocus      =   "CLOSED_C.frx":FE35
         colorDisabled   =   "CLOSED_C.frx":FE5F
         colorPressed    =   "CLOSED_C.frx":FE89
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
         Office2003ColorSouth=   113
         Office2003ColorBorder=   13290172
         Office2003ColorNorthMouseOver=   128
         Office2003ColorSouthMouseOver=   113
         Office2003ColorBorderMouseOver=   16777215
         Office2003ColorNorthPressed=   2171390
         Office2003ColorSouthPressed=   2171390
         Office2003ColorBorderPressed=   2171390
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand6 
         Height          =   630
         Left            =   4920
         TabIndex        =   39
         Top             =   6360
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   1111
         _StockProps     =   66
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
         Picture         =   "C:\Users\ben\Downloads\double_chevron_down_32x32.bmp"
         PictureYOffset  =   -3
         BackColorContainer=   13290172
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "CLOSED_C.frx":FEB3
         textCaption     =   "CLOSED_C.frx":FF50
         textLT          =   "CLOSED_C.frx":FF68
         textCT          =   "CLOSED_C.frx":FF80
         textRT          =   "CLOSED_C.frx":FF98
         textLM          =   "CLOSED_C.frx":FFB0
         textRM          =   "CLOSED_C.frx":FFC8
         textLB          =   "CLOSED_C.frx":FFE0
         textCB          =   "CLOSED_C.frx":FFF8
         textRB          =   "CLOSED_C.frx":10010
         colorBack       =   "CLOSED_C.frx":10028
         colorIntern     =   "CLOSED_C.frx":10052
         colorMO         =   "CLOSED_C.frx":1007C
         colorFocus      =   "CLOSED_C.frx":100A6
         colorDisabled   =   "CLOSED_C.frx":100D0
         colorPressed    =   "CLOSED_C.frx":100FA
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
         Office2003ColorBorder=   13290172
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
      Begin BTNENHLib4.BtnEnh SSCommand7 
         Height          =   630
         Left            =   4920
         TabIndex        =   40
         Top             =   5640
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   1111
         _StockProps     =   66
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
         Picture         =   "C:\Users\ben\Downloads\double_chevron_up_32x32_variant.bmp"
         BackColorContainer=   13290172
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "CLOSED_C.frx":10124
         textCaption     =   "CLOSED_C.frx":101CD
         textLT          =   "CLOSED_C.frx":101E5
         textCT          =   "CLOSED_C.frx":101FD
         textRT          =   "CLOSED_C.frx":10215
         textLM          =   "CLOSED_C.frx":1022D
         textRM          =   "CLOSED_C.frx":10245
         textLB          =   "CLOSED_C.frx":1025D
         textCB          =   "CLOSED_C.frx":10275
         textRB          =   "CLOSED_C.frx":1028D
         colorBack       =   "CLOSED_C.frx":102A5
         colorIntern     =   "CLOSED_C.frx":102CF
         colorMO         =   "CLOSED_C.frx":102F9
         colorFocus      =   "CLOSED_C.frx":10323
         colorDisabled   =   "CLOSED_C.frx":1034D
         colorPressed    =   "CLOSED_C.frx":10377
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
         Office2003ColorBorder=   13290172
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
      Begin BTNENHLib4.BtnEnh SSCommand1 
         Height          =   855
         Left            =   4920
         TabIndex        =   41
         Top             =   120
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   $"CLOSED_C.frx":103A1
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
         BackColorContainer=   13290172
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "CLOSED_C.frx":103B0
         textLT          =   "CLOSED_C.frx":10426
         textCT          =   "CLOSED_C.frx":1043E
         textRT          =   "CLOSED_C.frx":10456
         textLM          =   "CLOSED_C.frx":1046E
         textRM          =   "CLOSED_C.frx":10486
         textLB          =   "CLOSED_C.frx":1049E
         textCB          =   "CLOSED_C.frx":104B6
         textRB          =   "CLOSED_C.frx":104CE
         colorBack       =   "CLOSED_C.frx":104E6
         colorIntern     =   "CLOSED_C.frx":10510
         colorMO         =   "CLOSED_C.frx":1053A
         colorFocus      =   "CLOSED_C.frx":10564
         colorDisabled   =   "CLOSED_C.frx":1058E
         colorPressed    =   "CLOSED_C.frx":105B8
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
         Office2003ColorBorder=   13290172
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
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   855
         Left            =   4920
         TabIndex        =   50
         Top             =   1080
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Reprint Original Check"
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
         BackColorContainer=   13290172
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "CLOSED_C.frx":105E2
         textLT          =   "CLOSED_C.frx":1066E
         textCT          =   "CLOSED_C.frx":10686
         textRT          =   "CLOSED_C.frx":1069E
         textLM          =   "CLOSED_C.frx":106B6
         textRM          =   "CLOSED_C.frx":106CE
         textLB          =   "CLOSED_C.frx":106E6
         textCB          =   "CLOSED_C.frx":106FE
         textRB          =   "CLOSED_C.frx":10716
         colorBack       =   "CLOSED_C.frx":1072E
         colorIntern     =   "CLOSED_C.frx":10758
         colorMO         =   "CLOSED_C.frx":10782
         colorFocus      =   "CLOSED_C.frx":107AC
         colorDisabled   =   "CLOSED_C.frx":107D6
         colorPressed    =   "CLOSED_C.frx":10800
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
         Office2003ColorBorder=   13290172
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
      Begin VB.Label STAX 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   375
         Left            =   2760
         TabIndex        =   53
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   52
         Top             =   7560
         Width           =   2295
      End
      Begin VB.Label Cashier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   48
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Close By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Server 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label TableNumb 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label NIP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Table Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number In Party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label ServerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label CKNumb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label TBLNUmb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label PartyNumb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time Check Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label TimeCheckClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label CheckDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Ck_Totals 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   8040
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Total"
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
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   8040
         Width           =   2295
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2880
         Top             =   7920
         Width           =   1695
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00F8F6F4&
         FillStyle       =   0  'Solid
         Height          =   8535
         Left            =   120
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.PictureBox SearchBox 
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
      Height          =   8055
      Left            =   12240
      ScaleHeight     =   8025
      ScaleWidth      =   9945
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CheckBox ByTime 
         BackColor       =   &H00404040&
         Caption         =   "Show Items By Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   7560
         Width           =   3135
      End
      Begin VSFlex8DAOCtl.VSFlexGrid SearchItem 
         Height          =   6615
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   9495
         _cx             =   1975337324
         _cy             =   1975332244
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   65535
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   16711680
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
         Rows            =   100
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
      Begin BTNENHLib4.BtnEnh Sscommand12 
         Height          =   735
         Left            =   7680
         TabIndex        =   31
         Top             =   7200
         Width           =   2175
         _Version        =   589828
         _ExtentX        =   3836
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
         Surface         =   15
         BackColorContainer=   4210752
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "CLOSED_C.frx":1082A
         textLT          =   "CLOSED_C.frx":10892
         textCT          =   "CLOSED_C.frx":108AA
         textRT          =   "CLOSED_C.frx":108C2
         textLM          =   "CLOSED_C.frx":108DA
         textRM          =   "CLOSED_C.frx":108F2
         textLB          =   "CLOSED_C.frx":1090A
         textCB          =   "CLOSED_C.frx":10922
         textRB          =   "CLOSED_C.frx":1093A
         colorBack       =   "CLOSED_C.frx":10952
         colorIntern     =   "CLOSED_C.frx":1097C
         colorMO         =   "CLOSED_C.frx":109A6
         colorFocus      =   "CLOSED_C.frx":109D0
         colorDisabled   =   "CLOSED_C.frx":109FA
         colorPressed    =   "CLOSED_C.frx":10A24
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
      Begin VB.Label SearchCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search Results"
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
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   8895
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   6975
         Left            =   120
         Top             =   120
         Width           =   9735
      End
   End
   Begin VB.ListBox SearchDate 
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
      Height          =   5700
      Left            =   1560
      TabIndex        =   25
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox EndTime 
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
      Height          =   285
      Left            =   6360
      TabIndex        =   29
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox StartTime 
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
      Height          =   285
      Left            =   4800
      TabIndex        =   28
      Top             =   1200
      Width           =   1455
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   0
      Left            =   5160
      TabIndex        =   32
      Top             =   9480
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "List All Transactions"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":10A4E
      textLT          =   "CLOSED_C.frx":10AD8
      textCT          =   "CLOSED_C.frx":10AF0
      textRT          =   "CLOSED_C.frx":10B08
      textLM          =   "CLOSED_C.frx":10B20
      textRM          =   "CLOSED_C.frx":10B38
      textLB          =   "CLOSED_C.frx":10B50
      textCB          =   "CLOSED_C.frx":10B68
      textRB          =   "CLOSED_C.frx":10B80
      colorBack       =   "CLOSED_C.frx":10B98
      colorIntern     =   "CLOSED_C.frx":10BC2
      colorMO         =   "CLOSED_C.frx":10BEC
      colorFocus      =   "CLOSED_C.frx":10C16
      colorDisabled   =   "CLOSED_C.frx":10C40
      colorPressed    =   "CLOSED_C.frx":10C6A
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   1
      Left            =   5160
      TabIndex        =   33
      Top             =   3960
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Time"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":10C94
      textLT          =   "CLOSED_C.frx":10CFC
      textCT          =   "CLOSED_C.frx":10D14
      textRT          =   "CLOSED_C.frx":10D2C
      textLM          =   "CLOSED_C.frx":10D44
      textRM          =   "CLOSED_C.frx":10D5C
      textLB          =   "CLOSED_C.frx":10D74
      textCB          =   "CLOSED_C.frx":10D8C
      textRB          =   "CLOSED_C.frx":10DA4
      colorBack       =   "CLOSED_C.frx":10DBC
      colorIntern     =   "CLOSED_C.frx":10DE6
      colorMO         =   "CLOSED_C.frx":10E10
      colorFocus      =   "CLOSED_C.frx":10E3A
      colorDisabled   =   "CLOSED_C.frx":10E64
      colorPressed    =   "CLOSED_C.frx":10E8E
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   2
      Left            =   5160
      TabIndex        =   34
      Top             =   4680
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Employee Name"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":10EB8
      textLT          =   "CLOSED_C.frx":10F32
      textCT          =   "CLOSED_C.frx":10F4A
      textRT          =   "CLOSED_C.frx":10F62
      textLM          =   "CLOSED_C.frx":10F7A
      textRM          =   "CLOSED_C.frx":10F92
      textLB          =   "CLOSED_C.frx":10FAA
      textCB          =   "CLOSED_C.frx":10FC2
      textRB          =   "CLOSED_C.frx":10FDA
      colorBack       =   "CLOSED_C.frx":10FF2
      colorIntern     =   "CLOSED_C.frx":1101C
      colorMO         =   "CLOSED_C.frx":11046
      colorFocus      =   "CLOSED_C.frx":11070
      colorDisabled   =   "CLOSED_C.frx":1109A
      colorPressed    =   "CLOSED_C.frx":110C4
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   3
      Left            =   5160
      TabIndex        =   35
      Top             =   6120
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Menu Item"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":110EE
      textLT          =   "CLOSED_C.frx":11160
      textCT          =   "CLOSED_C.frx":11178
      textRT          =   "CLOSED_C.frx":11190
      textLM          =   "CLOSED_C.frx":111A8
      textRM          =   "CLOSED_C.frx":111C0
      textLB          =   "CLOSED_C.frx":111D8
      textCB          =   "CLOSED_C.frx":111F0
      textRB          =   "CLOSED_C.frx":11208
      colorBack       =   "CLOSED_C.frx":11220
      colorIntern     =   "CLOSED_C.frx":1124A
      colorMO         =   "CLOSED_C.frx":11274
      colorFocus      =   "CLOSED_C.frx":1129E
      colorDisabled   =   "CLOSED_C.frx":112C8
      colorPressed    =   "CLOSED_C.frx":112F2
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   5
      Left            =   5160
      TabIndex        =   36
      Top             =   2520
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Check No"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":1131C
      textLT          =   "CLOSED_C.frx":1138C
      textCT          =   "CLOSED_C.frx":113A4
      textRT          =   "CLOSED_C.frx":113BC
      textLM          =   "CLOSED_C.frx":113D4
      textRM          =   "CLOSED_C.frx":113EC
      textLB          =   "CLOSED_C.frx":11404
      textCB          =   "CLOSED_C.frx":1141C
      textRB          =   "CLOSED_C.frx":11434
      colorBack       =   "CLOSED_C.frx":1144C
      colorIntern     =   "CLOSED_C.frx":11476
      colorMO         =   "CLOSED_C.frx":114A0
      colorFocus      =   "CLOSED_C.frx":114CA
      colorDisabled   =   "CLOSED_C.frx":114F4
      colorPressed    =   "CLOSED_C.frx":1151E
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   4
      Left            =   5160
      TabIndex        =   37
      Top             =   6840
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Voided Sales"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":11548
      textLT          =   "CLOSED_C.frx":115C0
      textCT          =   "CLOSED_C.frx":115D8
      textRT          =   "CLOSED_C.frx":115F0
      textLM          =   "CLOSED_C.frx":11608
      textRM          =   "CLOSED_C.frx":11620
      textLB          =   "CLOSED_C.frx":11638
      textCB          =   "CLOSED_C.frx":11650
      textRB          =   "CLOSED_C.frx":11668
      colorBack       =   "CLOSED_C.frx":11680
      colorIntern     =   "CLOSED_C.frx":116AA
      colorMO         =   "CLOSED_C.frx":116D4
      colorFocus      =   "CLOSED_C.frx":116FE
      colorDisabled   =   "CLOSED_C.frx":11728
      colorPressed    =   "CLOSED_C.frx":11752
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorNorth=   4152316
      Office2003ColorSouth=   1916924
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   4152316
      Office2003ColorSouthMouseOver=   4152316
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   11648510
      Office2003ColorSouthPressed=   11648510
      Office2003ColorBorderPressed=   4152316
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh Backup 
      Height          =   615
      Left            =   1560
      TabIndex        =   38
      Top             =   8280
      Width           =   3135
      _Version        =   589828
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Show Current Week"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":1177C
      textLT          =   "CLOSED_C.frx":117FE
      textCT          =   "CLOSED_C.frx":11816
      textRT          =   "CLOSED_C.frx":1182E
      textLM          =   "CLOSED_C.frx":11846
      textRM          =   "CLOSED_C.frx":1185E
      textLB          =   "CLOSED_C.frx":11876
      textCB          =   "CLOSED_C.frx":1188E
      textRB          =   "CLOSED_C.frx":118A6
      colorBack       =   "CLOSED_C.frx":118BE
      colorIntern     =   "CLOSED_C.frx":118E8
      colorMO         =   "CLOSED_C.frx":11912
      colorFocus      =   "CLOSED_C.frx":1193C
      colorDisabled   =   "CLOSED_C.frx":11966
      colorPressed    =   "CLOSED_C.frx":11990
      TextCaptionAlignment=   0
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
      Office2003ColorNorth=   8404992
      Office2003ColorSouth=   7419904
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   8404992
      Office2003ColorSouthMouseOver=   8404992
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   15680287
      Office2003ColorSouthPressed=   15680287
      Office2003ColorBorderPressed=   15680287
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   6
      Left            =   5160
      TabIndex        =   42
      Top             =   7560
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   $"CLOSED_C.frx":119BA
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":119CD
      textLT          =   "CLOSED_C.frx":11A4B
      textCT          =   "CLOSED_C.frx":11A63
      textRT          =   "CLOSED_C.frx":11A7B
      textLM          =   "CLOSED_C.frx":11A93
      textRM          =   "CLOSED_C.frx":11AAB
      textLB          =   "CLOSED_C.frx":11AC3
      textCB          =   "CLOSED_C.frx":11ADB
      textRB          =   "CLOSED_C.frx":11AF3
      colorBack       =   "CLOSED_C.frx":11B0B
      colorIntern     =   "CLOSED_C.frx":11B35
      colorMO         =   "CLOSED_C.frx":11B5F
      colorFocus      =   "CLOSED_C.frx":11B89
      colorDisabled   =   "CLOSED_C.frx":11BB3
      colorPressed    =   "CLOSED_C.frx":11BDD
      Orientation     =   3
      TextCaptionAlignment=   0
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
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   7
      Left            =   5160
      TabIndex        =   44
      Top             =   3240
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Table/Tab Name"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":11C07
      textLT          =   "CLOSED_C.frx":11C83
      textCT          =   "CLOSED_C.frx":11C9B
      textRT          =   "CLOSED_C.frx":11CB3
      textLM          =   "CLOSED_C.frx":11CCB
      textRM          =   "CLOSED_C.frx":11CE3
      textLB          =   "CLOSED_C.frx":11CFB
      textCB          =   "CLOSED_C.frx":11D13
      textRB          =   "CLOSED_C.frx":11D2B
      colorBack       =   "CLOSED_C.frx":11D43
      colorIntern     =   "CLOSED_C.frx":11D6D
      colorMO         =   "CLOSED_C.frx":11D97
      colorFocus      =   "CLOSED_C.frx":11DC1
      colorDisabled   =   "CLOSED_C.frx":11DEB
      colorPressed    =   "CLOSED_C.frx":11E15
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   46
      Top             =   9000
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   66
      Caption         =   "Keyboard"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":11E3F
      textLT          =   "CLOSED_C.frx":11EAF
      textCT          =   "CLOSED_C.frx":11EC7
      textRT          =   "CLOSED_C.frx":11EDF
      textLM          =   "CLOSED_C.frx":11EF7
      textRM          =   "CLOSED_C.frx":11F0F
      textLB          =   "CLOSED_C.frx":11F27
      textCB          =   "CLOSED_C.frx":11F3F
      textRB          =   "CLOSED_C.frx":11F57
      colorBack       =   "CLOSED_C.frx":11F6F
      colorIntern     =   "CLOSED_C.frx":11F99
      colorMO         =   "CLOSED_C.frx":11FC3
      colorFocus      =   "CLOSED_C.frx":11FED
      colorDisabled   =   "CLOSED_C.frx":12017
      colorPressed    =   "CLOSED_C.frx":12041
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   9
      Left            =   5160
      TabIndex        =   51
      Top             =   8280
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Address"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":1206B
      textLT          =   "CLOSED_C.frx":120D9
      textCT          =   "CLOSED_C.frx":120F1
      textRT          =   "CLOSED_C.frx":12109
      textLM          =   "CLOSED_C.frx":12121
      textRM          =   "CLOSED_C.frx":12139
      textLB          =   "CLOSED_C.frx":12151
      textCB          =   "CLOSED_C.frx":12169
      textRB          =   "CLOSED_C.frx":12181
      colorBack       =   "CLOSED_C.frx":12199
      colorIntern     =   "CLOSED_C.frx":121C3
      colorMO         =   "CLOSED_C.frx":121ED
      colorFocus      =   "CLOSED_C.frx":12217
      colorDisabled   =   "CLOSED_C.frx":12241
      colorPressed    =   "CLOSED_C.frx":1226B
      Orientation     =   3
      TextCaptionAlignment=   0
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
   Begin BTNENHLib4.BtnEnh BtnEnh4 
      Height          =   615
      Left            =   1560
      TabIndex        =   54
      Top             =   9000
      Width           =   3135
      _Version        =   589828
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Search History"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":12295
      textLT          =   "CLOSED_C.frx":12311
      textCT          =   "CLOSED_C.frx":12329
      textRT          =   "CLOSED_C.frx":12341
      textLM          =   "CLOSED_C.frx":12359
      textRM          =   "CLOSED_C.frx":12371
      textLB          =   "CLOSED_C.frx":12389
      textCB          =   "CLOSED_C.frx":123A1
      textRB          =   "CLOSED_C.frx":123B9
      colorBack       =   "CLOSED_C.frx":123D1
      colorIntern     =   "CLOSED_C.frx":123FB
      colorMO         =   "CLOSED_C.frx":12425
      colorFocus      =   "CLOSED_C.frx":1244F
      colorDisabled   =   "CLOSED_C.frx":12479
      colorPressed    =   "CLOSED_C.frx":124A3
      TextCaptionAlignment=   0
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
      Office2003ColorNorth=   8404992
      Office2003ColorSouth=   7419904
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   8404992
      Office2003ColorSouthMouseOver=   8404992
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   15680287
      Office2003ColorSouthPressed=   15680287
      Office2003ColorBorderPressed=   15680287
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin VB.Frame Message 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Make A Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   8400
      TabIndex        =   55
      Top             =   2400
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "Clear List"
         Height          =   615
         Left            =   120
         TabIndex        =   57
         Top             =   8400
         Width           =   3735
      End
      Begin VB.ListBox List1 
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
         ForeColor       =   &H00000000&
         Height          =   6780
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   3975
      End
   End
   Begin BTNENHLib4.BtnEnh SSCommand2 
      Height          =   975
      Left            =   8400
      TabIndex        =   58
      Top             =   9240
      Width           =   3975
      _Version        =   589828
      _ExtentX        =   7011
      _ExtentY        =   1720
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
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":124CD
      textLT          =   "CLOSED_C.frx":12535
      textCT          =   "CLOSED_C.frx":1254D
      textRT          =   "CLOSED_C.frx":12565
      textLM          =   "CLOSED_C.frx":1257D
      textRM          =   "CLOSED_C.frx":12595
      textLB          =   "CLOSED_C.frx":125AD
      textCB          =   "CLOSED_C.frx":125C5
      textRB          =   "CLOSED_C.frx":125DD
      colorBack       =   "CLOSED_C.frx":125F5
      colorIntern     =   "CLOSED_C.frx":1261F
      colorMO         =   "CLOSED_C.frx":12649
      colorFocus      =   "CLOSED_C.frx":12673
      colorDisabled   =   "CLOSED_C.frx":1269D
      colorPressed    =   "CLOSED_C.frx":126C7
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
      Office2003ColorSouth=   113
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   128
      Office2003ColorSouthMouseOver=   113
      Office2003ColorBorderMouseOver=   16777215
      Office2003ColorNorthPressed=   2171390
      Office2003ColorSouthPressed=   2171390
      Office2003ColorBorderPressed=   2171390
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh SearchHow 
      Height          =   615
      Index           =   10
      Left            =   5160
      TabIndex        =   59
      Top             =   5400
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Payment Type"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Surface         =   15
      BackColorContainer=   8421504
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "CLOSED_C.frx":126F1
      textLT          =   "CLOSED_C.frx":12769
      textCT          =   "CLOSED_C.frx":12781
      textRT          =   "CLOSED_C.frx":12799
      textLM          =   "CLOSED_C.frx":127B1
      textRM          =   "CLOSED_C.frx":127C9
      textLB          =   "CLOSED_C.frx":127E1
      textCB          =   "CLOSED_C.frx":127F9
      textRB          =   "CLOSED_C.frx":12811
      colorBack       =   "CLOSED_C.frx":12829
      colorIntern     =   "CLOSED_C.frx":12853
      colorMO         =   "CLOSED_C.frx":1287D
      colorFocus      =   "CLOSED_C.frx":128A7
      colorDisabled   =   "CLOSED_C.frx":128D1
      colorPressed    =   "CLOSED_C.frx":128FB
      Orientation     =   3
      TextCaptionAlignment=   0
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
      Office2003ColorSouth=   15263976
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   12632256
      Office2003ColorNorthPressed=   16777215
      Office2003ColorSouthPressed=   16645629
      Office2003ColorBorderPressed=   16739100
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Make a Selection:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   365
      Left            =   8400
      TabIndex        =   60
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9000
      TabIndex        =   49
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Criteria:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   365
      Left            =   5040
      TabIndex        =   27
      Top             =   2040
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   365
      Left            =   1560
      TabIndex        =   26
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00242424&
      FillColor       =   &H00F8F6F4&
      FillStyle       =   0  'Solid
      Height          =   7815
      Left            =   5040
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Date:"
      ForeColor       =   &H00F8F6F4&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Searchwhen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF7B00&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F8F6F4&
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F8F6F4&
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label FileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Count:"
      ForeColor       =   &H00F8F6F4&
      Height          =   255
      Left            =   9000
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00402606&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1560
      Top             =   600
      Width           =   10815
   End
   Begin VB.Menu Quit 
      Caption         =   "Exit"
   End
   Begin VB.Menu STS 
      Caption         =   "Serving Time Summary"
   End
   Begin VB.Menu SBI 
      Caption         =   "Sales By Individual"
   End
   Begin VB.Menu UT 
      Caption         =   "Utilities"
      Begin VB.Menu ImportBack 
         Caption         =   "Import BackUps"
      End
      Begin VB.Menu SMC 
         Caption         =   "Search For Missing Checks"
      End
   End
End
Attribute VB_Name = "Closed_Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CheckOpenTime As String
Private BegIn_Line As Integer
Private CameraNo   As String
Private SearchWhat As String
Private PhoneSearch As String
Private DayNumber   As String
Private SearchDay   As Integer
Private Counter     As Integer
Private ToTDiff     As Currency
Private Search_Path As String
Private FiletoLookat As String
Private GTotal As Currency
Private NTotal As Currency
Private I_total As Currency
Private LastCheckNo  As Long
Private LN      As Long
Private TDVAT As Currency
Private TVDIF As Currency
Private Type Dates
     StartDate As String
     EndDate   As String
End Type
Private Enum dbTableName
     closedchecks
     checkdetail
     CheckPayments
End Enum
 
Private Sub Backup_Click()
LoadCurrentWeek
End Sub

'***********************

'***********************
Private Sub Backup_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
End Sub
Private Sub BtnEnh1_Click()
End Sub
Private Sub BtnEnh2_Click()
End Sub
Private Sub BtnEnh3_Click()
Dim MyCheck As String
MyCheck = CheckDate.Caption ' & " " & TimeCheckClosed.Caption
If ReprintCustomerReceipt(MyCheck, CKNumb) = False Then
   See_Msg "Original Copy Not Found.", 5, "Reprinting Check"
   SSCommand1_Click
End If
End Sub
Private Sub BtnEnh4_Click()
LoadDates "Select Distinct [Date] from [Sales] order by [Date] asc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
FiletoLookat = "MSaleData"
Dim Db      As ADODB.Connection
Dim Rcset   As ADODB.Recordset
  ConnectADODB SvrPath.TerminalPath & FiletoLookat, Db
     Sql = "Select distinct [Order Date] from [Closed checks]"
     Sql = Sql & " Where TRY_CAST([order date] AS date) Between '" & SearchDates.Item(1).StartDate & "' and '" & SearchDates.Item(1).EndDate & "'"
  ConnectADORS Sql, Db, Rcset
  SearchDate.Clear
  Do While Not Rcset.EOF
     SearchDate.AddItem Format(Rcset.fields("Order Date").Value, "MM/DD/YYYY")
     Rcset.MoveNext
  Loop
DbClose Db, Rcset
FileName.Caption = FiletoLookat
'LoadHistoryList
End Sub
Private Sub Command1_Click()
Dim LT As Integer
  If Counter >= SearchItem.rows Then
   See_Msg "End Of Day", 5, ""
   Exit Sub
  End If
  If IsNumeric(SearchItem.TextMatrix(Counter, 0)) Then
   View_Checks Val(SearchItem.TextMatrix(Counter, 0))
  End If
  Counter = Counter + 1
End Sub

Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Command3_Click()

End Sub

Private Sub EndTime_Click()
Dim St As Datetime.DateOrTime
Set St = New Datetime.DateOrTime
EndTime.Text = St.ReturnTime
Set St = Nothing
End Sub
Private Sub ExitButton_Click()
End Sub
Public Sub SalesByInd(StartDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Records     As ADODB.Recordset
Dim Tb          As ADODB.Recordset
Dim RTB         As ADODB.Recordset
Dim T_Rec       As Long
Dim R           As Long
Dim CSvrName    As String
Dim Data        As String
Dim T_Voids     As Currency
Dim T_Discount  As Currency
Dim T_Gross     As Currency
Dim T_Payment   As Currency
Dim T_Items     As Long
Dim T_Count     As Long
Dim TFileName   As String
Dim LCount      As Long
Dim Orderdate   As String
If FiletoLookat = "Chk StatSql" Or FiletoLookat = "MSaleData" Then
 ConnectADODB Search_Path & FiletoLookat, Db
   Sql = "Select * from [Closed Checks]"
   Sql = Sql & " Where TRY_CAST([Order date] AS datetime) + TRY_CAST([Open Time] AS datetime)  Between '" & CDate(StartDate & " " & GetStartTime()) & "'"
   Sql = Sql & " and '" & CDate(SetEndingDate(EndDate) & " " & SetEndingTime()) & "'"
   Sql = Sql & " Order by [Server Name],TRY_CAST([Order Date] AS datetime),TRY_CAST([Open Time] AS datetime) asc"
 Else
 TFileName = FiletoLookat
 ConnectADODB Search_Path & TFileName, Db
  Sql = "Select * from [Closed Checks] "
  Sql = Sql & " Where TRY_CAST([Order date] AS datetime) + TRY_CAST([Open Time] AS datetime) Between '" & Format(CDate(CDate(StartDate & " " & GetStartTime())), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(EndDate) & " " & SetEndingTime())), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Order by [Server Name],TRY_CAST([Order Date] AS datetime),TRY_CAST([Open Time] AS datetime) asc"
End If
 ConnectADORS Sql, Db, Tb
 If Tb.EOF Then
    See_Msg "No Sales found", 5, ""
    DbClose Db, Tb
    Exit Sub
 End If
Update.Show
Update.UPdateWhere.Caption = "Compiling sales data...."
Update.Refresh
'Update.UPdateWhere.Caption
T_Rec = RecordsCount(Tb)
R = 0
 CSvrName = "": Data = ""
 Do While Not Tb.EOF
        R = R + 1
        Update.Progress.ProgressBarValue = (R / T_Rec) * 100
        Update.Progress.Refresh
        'Sleep 0
        If Tb(0) <> CSvrName Then
         LCount = 0
         Data = Data & Tb(0) & "|"
         Data = Data & ">>" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & ";"
         CSvrName = Tb(0)
         T_Voids = 0: T_Discount = 0: T_Gross = 0: T_Payment = 0: T_Items = 0
        End If
        LCount = LCount + 1
        Data = Data & Format(Tb![Order Date], "MM/DD/YY") & "|"
        Data = Data & Format(Tb![Open Time], "HH:MM ampm") & "|"
        Data = Data & Tb![Register Name] & "|"
        Data = Data & Tb![Check Number] & "|"
        Orderdate = Tb![Order Date] & ""
        Sql = "Select [Servers Name], Sum([Quantity]) as TItems From [Check Detail] where [Check Number] = " & Tb![Check Number]
        Sql = Sql & " and [Close Out Day] = " & Tb![Close out day]
        Sql = Sql & " and [Servers Name] = '" & Tb![Server Name] & "'"
        If FiletoLookat = "Chk StatSql" Or FiletoLookat = "MSaleData" Then
          Sql = Sql & " and TRY_CAST([Check Date] AS date) = '" & Orderdate & "'"
        Else
          Sql = Sql & " and TRY_CAST([Check Date] AS datetime) = '" & Format(CDate(Orderdate), "yyyy-mm-dd hh:nn:ss") & "'"
        End If
        Sql = Sql & " Group By [Servers Name]"
        
        ConnectADORS Sql, Db, RTB
          If Not RTB.EOF Then
            Data = Data & RTB![TItems] & "|"
            T_Items = T_Items + RTB![TItems]
          Else
             Data = Data & "No Items" & "|"
          End If
         Data = Data & Format(Tb![Gross Sales], "0.00") & "|"
         T_Gross = T_Gross + Tb![Gross Sales]
         
         Data = Data & Format((Abs(Tb![Gross Discounts]) - Abs(Tb![Voided Discounts])), "0.00") & "|"
         T_Discount = T_Discount + (Abs(Tb![Gross Discounts]) - Abs(Tb![Voided Discounts]))
         
         Data = Data & Format(Tb![Voided Sales], "0.00") & "|"
         T_Voids = T_Voids + Tb![Voided Sales]
         
         Data = Data & Format(Tb![Check Total], "0.00") & "|"
         T_Payment = T_Payment + Tb![Check Total]
            
            Sql = "Select * From [Payments] where [Check Number] = " & Tb![Check Number]
            If FiletoLookat = "Chk StatSql" Or FiletoLookat = "MSaleData" Then
             Sql = Sql & " and TRY_CAST([date] AS date) = '" & Orderdate & "'"
            Else
             Sql = Sql & " and TRY_CAST([date] AS datetime) = '" & Format(CDate(Orderdate), "yyyy-mm-dd hh:nn:ss") & "'"
            End If
            Sql = Sql & " and [Close Out Day] = " & Tb![Close out day]
            ConnectADORS Sql, Db, Records
              If Not Records.EOF Then
                Data = Data & Records![Payment Type] & ";"
              Else
                Data = Data & "No Payment" & ";"
              End If
          SetNothing RTB: SetNothing Records
    Tb.MoveNext
    If Tb.EOF Then
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "Totals" & "|"
         Data = Data & "" & "|"
         Data = Data & T_Items & "|"
         Data = Data & Format(T_Gross, "0.00") & "|"
         Data = Data & Format(T_Discount, "0.00") & "|"
         Data = Data & Format(T_Voids, "0.00") & "|"
         Data = Data & Format(T_Payment, "0.00") & "|"
         Data = Data & "" & ";"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "Average" & "|"
         Data = Data & "Sale" & "|"
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & Format(T_Payment / LCount, "0.00") & "|"
         Data = Data & "" & ";"
    
    Else
      If CSvrName <> Tb![Server Name] Then
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "Totals" & "|"
         Data = Data & "" & "|"
         Data = Data & T_Items & "|"
         Data = Data & Format(T_Gross, "0.00") & "|"
         Data = Data & Format(T_Discount, "0.00") & "|"
         Data = Data & Format(T_Voids, "0.00") & "|"
         Data = Data & Format(T_Payment, "0.00") & "|"
         Data = Data & "" & ";"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "Average" & "|"
         Data = Data & "Sale" & "|"
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & Format(T_Payment / LCount, "0.00") & "|"
         Data = Data & "" & ";"
         
     End If
   End If
Loop
Unload Update
DbClose Db, Tb
Dim H$
H = "Date|Time|Terminal|Check #|#|Gross Sale|Discount|Voids|Check Total|Tender;"
Dim F$
F = "+<1600|+<1000|+>1600|+>750|+>500|+>1000|+>1000|+>1000|+>1000|+>1500;"
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

Public Sub SummaryServeTimes(StartDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Records     As ADODB.Recordset
Dim Tb          As ADODB.Recordset
Dim RTB         As ADODB.Recordset
Dim SvrRs       As ADODB.Recordset
Dim T_Rec       As Long
Dim R           As Long
Dim CSvrName    As String
Dim Data        As String
Dim T_Voids     As Currency
Dim T_Discount  As Currency
Dim T_Gross     As Currency
Dim T_Net       As Currency
Dim T_Count     As Long
Dim TFileName   As String
Dim LCount      As Long
Dim T_Min       As Long
If FiletoLookat = "Chk StatSql" Then
 ConnectADODB Search_Path & FiletoLookat, Db
  Sql = "Select [Server Name] from [Closed Checks] where TRY_CAST([order date] AS datetime) Between '" & StartDate & "' and '" & EndDate & "'"
  Sql = Sql & " Order by [Server Name]"
Else
  TFileName = FiletoLookat
  ConnectADODB Search_Path & TFileName, Db
   Sql = "Select [Server Name] from [Closed Checks] where TRY_CAST([order date] AS datetime) Between '" & Format(CDate(StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
   Sql = Sql & " Group by [Server Name]"
End If
 ConnectADORS Sql, Db, Tb
 If Tb.EOF Then
    See_Msg "No Sales found", 5, ""
    DbClose Db, Tb
    Exit Sub
 End If
Update.Show
Update.UPdateWhere.Caption = "Compiling sales data...."
Update.Refresh
 T_Rec = RecordsCount(Tb)
 R = 0
 Data = ""
 Do While Not Tb.EOF
   T_Voids = 0: T_Discount = 0: T_Gross = 0:   LCount = 0: T_Min = 0: T_Net = 0
   If FiletoLookat = "Chk StatSql" Then
     Sql = "Select * from [Closed Checks] where TRY_CAST([order date] AS datetime) Between '" & StartDate & "' and '" & EndDate & "'"
     Sql = Sql & " and [Server Name] = '" & Tb![Server Name] & "'"
     Sql = Sql & " Order by TRY_CAST([Open Time] AS datetime) asc"
   Else
     Sql = "Select * from [Closed Checks] where TRY_CAST([order date] AS datetime) Between '" & Format(CDate(StartDate), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(EndDate), "yyyy-mm-dd hh:nn:ss") & "'"
     Sql = Sql & " and [Server Name] = '" & Tb![Server Name] & "'"
     Sql = Sql & " Order by TRY_CAST([Open Time] AS datetime) asc"
   End If
   ConnectADORS Sql, Db, SvrRs
     R = 0
     Do While Not SvrRs.EOF
         R = R + 1
         Update.Progress.ProgressBarValue = (R / T_Rec) * 100
         Update.Progress.Refresh
         LCount = LCount + 1
         If CDate(Format(SvrRs![Close Time], "HH:MM ampm")) < CDate(Format(SvrRs![Open Time], "HH:MM ampm")) Then
          T_Min = T_Min + DateDiff("n", SvrRs![Order Date] & " " & SvrRs![Open Time], CDate(SvrRs![Order Date]) + 1 & " " & SvrRs![Close Time])
         Else
          T_Min = T_Min + DateDiff("n", SvrRs![Order Date] & " " & SvrRs![Open Time], SvrRs![Order Date] & " " & SvrRs![Close Time])
         End If
         T_Gross = T_Gross + SvrRs![Gross Sales]
         T_Discount = T_Discount + (Abs(SvrRs![Gross Discounts]) - Abs(SvrRs![Voided Discounts]))
         T_Voids = T_Voids + SvrRs![Voided Sales]
         T_Net = ((T_Gross - T_Discount) - T_Voids)
        SvrRs.MoveNext
     Loop
     Data = Data & Tb(0) & "|"
     Data = Data & Format(T_Gross, "0.00") & "|"
     Data = Data & Format(T_Discount, "0.00") & "|"
     Data = Data & Format(T_Voids, "0.00") & "|"
     Data = Data & Format(T_Net, "0.00") & "|"
     If T_Net > 0 Then
       Data = Data & Format(T_Net / LCount, "0.00") & "|"
     Else
       Data = Data & "" & "|"
     End If
     If T_Min > 0 Then
       Data = Data & Format(T_Min / LCount, "0.00") & ";"
     Else
       Data = Data & "" & ";"
     End If
    Tb.MoveNext
Loop
Unload Update
DbClose Db, Tb
Dim H$
H = "Name|Gross Sales|Discounts|Voids|Net Sales|Avg Sales|Avg Srv Time;"
Dim F$
F = "<2000|+>1200|+>1200|+>1200|+>1200|+>1200|+>1200;"
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

Private Sub Form_Load()
On Local Error GoTo lderror
ResizeScreen1024 Me
CenterForm Me
SetMemberToNothing
ExtraData = ""
'need to be set to nothing list last charge on next reprinted check
With SearchItem
    .ColAlignment(0) = 1
    .ColWidth(0) = 950 * WidScale800x600
    .ColAlignment(1) = 1
    .ColWidth(1) = 1500 * WidScale800x600
    .ColAlignment(2) = 1
    .ColWidth(2) = 2200 * WidScale800x600
    .ColAlignment(3) = 1
    .ColWidth(3) = 1200 * WidScale800x600
    .ColAlignment(4) = 8
    .ColWidth(4) = 1000 * WidScale800x600
    .TextMatrix(0, 0) = "Chk No"
    .TextMatrix(0, 1) = "Type"
    .TextMatrix(0, 2) = "Name"
    .TextMatrix(0, 3) = "Amount"
    .TextMatrix(0, 4) = "Time"
End With
    FiletoLookat = CheckStatDb
    Search_Path = SvrPath.TerminalPath
    BegIn_Line = -1
    Dim RS As Integer
    Ticket(0).HighLight = flexHighlightNever
    SeatNumber = 0: If Number_In_Party = 0 Then Number_In_Party = 4
    RS = 0
    Ticket(RS).HighLight = flexHighlightNever
    Ticket(RS).rows = MaxRows
    Ticket(RS).ColAlignment(0) = flexAlignLeftCenter
    Ticket(RS).ColAlignment(1) = flexAlignLeftCenter
    Ticket(RS).ColAlignment(2) = flexAlignRightCenter
    Dim i As Integer
    For i = 3 To TTL_Col
        Ticket(RS).ColWidth(i) = 0
    Next
    Ticket(RS).ColWidth(0) = 425 * WidScale800x600
    Ticket(RS).ColWidth(1) = 2050 * WidScale800x600
    Ticket(RS).ColWidth(2) = 925 * WidScale800x600
    StartTime = Format(Time, "HH:MM")
    EndTime = Format(Time, "HH:MM")
    Searchwhen.Caption = Format(Date, "mm/dd/yyyy")
    FileName.Caption = CheckStatDb
    SearchDate.Clear
    If Check_Option(92) = 1 Then SearchHow(6).Caption = "Member No"
    DeleteFile SvrPath.LocalTerminalPath & "DAT Files\" & "OutPutData.Dat"
'used to track checks and payments
LoadCurrentWeek
LoadHistoryList
If BegIn_Line = -1 Then Exit Sub
SearchDate.Selected(0) = True
Sleep 0
Exit Sub
lderror:
MsgBox ("Error : " & Error), 0, "Error Number : " & Err, 0, 0
'Resume
Resume Next
End Sub
Private Sub LoadCurrentWeek()
    SearchDate.Clear
    Dim Db    As ADODB.Connection
    Dim Rcset As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db

    Sql = "Select distinct TRY_CAST([Order Date] AS datetime) as [Order Date] " & _
          "from [Closed Checks] " & _
          "order by TRY_CAST([Order Date] AS datetime) asc"

    ConnectADORS Sql, Db, Rcset
    Do While Not Rcset.EOF
        SearchDate.AddItem Format(Rcset.fields("Order Date").Value, "MM/DD/YYYY")
        Rcset.MoveNext
    Loop
    DbClose Db, Rcset
    FileName.Caption = CheckStatDb
End Sub
Private Sub LoadHistoryList()
 On Error Resume Next
 Dim BkUpFile As String
  
  BkUpFile = Dir(SvrPath.TerminalPath & "* BackUp.mdb")
  If BkUpFile = vbNullString Then Exit Sub
  List1.Clear
  Do Until BkUpFile = vbNullString
   If InStr(BkUpFile, "House") > 0 Or InStr(BkUpFile, "Time") > 0 Or InStr(BkUpFile, "Frequent") > 0 Then
     
   Else
     List1.AddItem BkUpFile
   End If
   BkUpFile = Dir
  Loop
End Sub
Private Function GetDates(FileName As String) As Dates
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
ConnectADODB Search_Path & FileName, Db
  Sql = "Select TRY_CAST([Check date] AS datetime) from [Check Detail] Order By TRY_CAST([Check Date] AS datetime) asc"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
If Not RS.BOF Then RS.MoveFirst
 GetDates.StartDate = Format(RS.fields("Check Date").Value, "MM/DD/YYYY")
If Not RS.EOF Then RS.MoveLast
 GetDates.EndDate = Format(RS.fields("Check Date").Value, "MM/DD/YYYY")
End If
DbClose Db, RS
End Function

Private Sub Form_Unload(Cancel As Integer)
ExtraData = ""
End Sub

Private Sub ImportBack_Click()
If See_Msg("Import Sale Data To History?", 4, "") <> 6 Then Exit Sub
LoadHistoryList
Dim PDate   As Dates
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim result  As String
Dim C       As Integer
For C = 0 To List1.ListCount - 1
 FiletoLookat = List1.List(C)
  PDate = GetDates(FiletoLookat)
  If IsDate(PDate.StartDate) And IsDate(PDate.EndDate) Then
     Sql = "Delete from [Closed Checks] where TRY_CAST([Order Date] AS date) Between '" & PDate.StartDate & "' and '" & PDate.EndDate & "'"
     ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
     Db.Execute Sql
     Sql = "Delete from [Check Detail] where TRY_CAST([Check Date] AS date) Between '" & PDate.StartDate & "' and '" & PDate.EndDate & "'"
     Db.Execute Sql
     Sql = "Delete from [Payments] where TRY_CAST([Date] AS date) Between '" & PDate.StartDate & "' and '" & PDate.EndDate & "'"
     Db.Execute Sql
     Db.Close
  End If
  
  If InStr(FiletoLookat, "Week") > 0 Then
   result = MoveTableData(SvrPath.TerminalPath, FiletoLookat, SvrPath.TerminalPath, "MSaleData", "Closed Checks", "CopyOnly")
   See_Msg "Import Complete for Closed Checks", 5, result
   result = MoveTableData(SvrPath.TerminalPath, FiletoLookat, SvrPath.TerminalPath, "MSaleData", "Check Detail", "CopyOnly")
   See_Msg "Import Complete for Check Detail", 5, result
   result = MoveTableData(SvrPath.TerminalPath, FiletoLookat, SvrPath.TerminalPath, "MSaleData", "Payments", "CopyOnly")
   See_Msg "Import Complete for Payments", 5, result
   CopyFile Search_Path & FiletoLookat, SvrPath.TerminalPath & "Daily BackUp Files\" & FiletoLookat, True
   DeleteFile Search_Path & FiletoLookat
  End If
Next C
LoadHistoryList
End Sub
Private Sub List1_Click()
On Error Resume Next
Dim sDb     As ADODB.Connection
Dim SRS     As ADODB.Recordset
If SearchWhat = "Server" Or SearchWhat = "Menu Items" Or SearchWhat = "Customer Name" Or SearchWhat = "Payments" Then
ConnectADODB Search_Path & FiletoLookat, sDb
    Select Case SearchWhat
        Case Is = "Server"
            Sql = "Select distinct [Check Number],[Servers Name] from [Check Detail] Where [Servers Name] = '" & List1.List(List1.ListIndex) & "'"
            Sql = Sql & " and [close out day] = " & SearchDay
            'Sql = Sql & GetDateTimeRange(Searchwhen, checkdetail)
            If Right$(UCase(FiletoLookat), 3) = "MDB" Then
              Sql = Sql & GetMDBDateTimeRange(Searchwhen, checkdetail)
            Else
              Sql = Sql & GetDateTimeRange(Searchwhen, checkdetail)
            End If
        Case Is = "Customer Name"
            Dim GETCK() As String
            GETCK = Split(List1.List(List1.ListIndex), "-CK#-")
            Sql = "Select distinct [Check Number],[Servers Name] from [Check Detail] Where [Check Number] = " & Val(GETCK(1))
            'Sql = Sql & " and [Check Date] = '" & Searchwhen & "'"
            If Right$(UCase(FiletoLookat), 3) = "MDB" Then
              Sql = Sql & GetMDBDateTimeRange(Searchwhen, checkdetail)
            Else
              Sql = Sql & GetDateTimeRange(Searchwhen, checkdetail)
            End If
           ' Sql = Sql & GetDateTimeRange(Searchwhen, checkdetail)
        Case Is = "Payments"
            Sql = "Select distinct [Check Number] from [Payments] Where [Payment Type] = '" & List1.List(List1.ListIndex) & "'"
            Sql = Sql & " and [close out day] = " & SearchDay
            If Right$(UCase(FiletoLookat), 3) = "MDB" Then
              Sql = Sql & GetMDBDateTimeRange(Searchwhen, CheckPayments)
            Else
              Sql = Sql & GetDateTimeRange(Searchwhen, CheckPayments)
            End If
        
        Case Else
            Sql = "Select distinct [Check Number],[Servers Name] from [Check Detail] Where [Menu Item] = '" & List1.List(List1.ListIndex) & "'"
            'Sql = Sql & " and [Check Date] = '" & Searchwhen & "'"
            If Right$(UCase(FiletoLookat), 3) = "MDB" Then
              Sql = Sql & GetMDBDateTimeRange(Searchwhen, checkdetail)
            Else
              Sql = Sql & GetDateTimeRange(Searchwhen, checkdetail)
            End If
            'If SVRLST.Text <> "None" Then
            '   Sql = Sql & " and [Servers Name] = '" & SVRLST.Text & "'"
            'End If
    End Select
ConnectADORS Sql, sDb, SRS
   Dim NR As Integer
   If Not SRS.EOF Then
      With SearchItem
          .rows = RecordsCount(SRS) + 1
           NR = 0
            .Clear
            .TextMatrix(0, 0) = "Chk No"
            .TextMatrix(0, 1) = ""
            .TextMatrix(0, 2) = "Name"
            .TextMatrix(0, 3) = ""
            .TextMatrix(0, 4) = ""
           If System_config = ExpressService Then
              .TextMatrix(0, 1) = ""
           Else
              .TextMatrix(0, 1) = ""
           End If
            'SearchBox.Left = 1320 * WidScale: SearchBox.Top = 600 * TopScale
            CenterObject Me, SearchBox

            Do While Not SRS.EOF
              NR = NR + 1
              .TextMatrix(NR, 0) = SRS![Check Number]
              .TextMatrix(NR, 2) = SRS![Servers Name]
              '.TextMatrix(NR, 3) = Format(SRS![Recorded time], "HH:MM ampm")
             SRS.MoveNext
            Loop
            SearchBox.Visible = True
            SearchBox.ZOrder 0
            Sleep 0
      End With

    Else
      Call See_Msg("No Data Found", 5, "")
    End If
    DbClose sDb, SRS
    SSCommand9_Click
    Exit Sub
    End If
If SearchWhat = "Back Up" Or SearchWhat = "Import" Or SearchWhat = "" Then
    FiletoLookat = List1.List(List1.ListIndex)
    Search_Path = SvrPath.TerminalPath
    FileName.Caption = FiletoLookat
End If
If SearchWhat <> "Import" Then
  WhichFile Search_Path, FiletoLookat
End If
SSCommand9_Click
End Sub
Private Sub WhichFile(PathtoFile As String, FileName As String)
Search_Path = PathtoFile
FiletoLookat = FileName
If InStr(FiletoLookat, "House") > 0 Then
   See_Msg "House Account Back Up", 5, "Cannot View"
   Exit Sub
End If
Dim Db As ADODB.Connection, Rcset As ADODB.Recordset
ConnectADODB PathtoFile & FileName, Db
    Sql = "Select distinct [Order Date] from [Closed checks] order by [Order Date] asc"
ConnectADORS Sql, Db, Rcset
  If Not Rcset.EOF Then SearchDate.Clear
  Do While Not Rcset.EOF
     SearchDate.AddItem Rcset![Order Date]
     Rcset.MoveNext
  Loop
DbClose Db, Rcset
End Sub
Private Sub Quit_Click()
SSCommand2_Click
End Sub
Private Sub SBI_Click()
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & ClosingHistory, SL
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
    SalesByInd SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
  End If
Else
   See_Msg "Not Dates Found", 5, ""
End If

End Sub

Private Sub SearchDate_Click()
 If Trim$(SearchDate.List(SearchDate.ListIndex)) = "" Then Exit Sub
 Dim Db As ADODB.Connection
 Dim Tb As ADODB.Recordset
 If Left$(FiletoLookat, 3) = "MDB" Then
   If FileCheck(Search_Path & FiletoLookat) = False Then
      FiletoLookat = "MSaleData"
   End If
 End If
 If ConnectADODB(Search_Path & FiletoLookat, Db) = False Then
    See_Msg "Error Connection To " & FiletoLookat, 5, ""
    Exit Sub
 End If
 Sql = "Select * from [Closed Checks] where [Order Date] = '" & SearchDate.List(SearchDate.ListIndex) & "'"
 ConnectADORS Sql, Db, Tb
 Label6.Caption = RecordsCount(Tb)
 Searchwhen = SearchDate.List(SearchDate.ListIndex)
 If Not Tb.EOF Then
    SearchDay = Val(Tb![Close out day] & "")
 End If
 Sleep 0
 DbClose Db, Tb
End Sub
Private Sub Searching_Click()
End Sub
Private Sub SearchHow_Click(Index As Integer)
Dim sDb As ADODB.Connection, SRS As ADODB.Recordset
Dim NR As Integer, TotalSale As Currency, Scheck As Variant
GTotal = 0: NTotal = 0: I_total = 0: LN = 0: TDVAT = 0: TVDIF = 0
PhoneSearch = ""
If Right$(UCase(FiletoLookat), 3) = "MDB" Then
     If FileCheck(Search_Path & FiletoLookat) = False Then
        FiletoLookat = "Chk StatSql"
     End If
End If
List1.Clear
ConnectADODB Search_Path & FiletoLookat, sDb
   Select Case Index
      Case Is = 0
          Sql = "Select * from [Closed Checks]"
          If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
             Sql = Sql & " Where TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
          Else
             Sql = Sql & " Where cdate([order Date] & ' ' & [Open time]) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
          End If
          Sql = Sql & " Order by [Check Number] asc"
      Case Is = 1
          Sql = "Select * from [Closed Checks] Where "
         If IsDate(StartTime) And IsDate(EndTime) Then
          Sql = Sql & " TRY_CAST([Open Time] AS datetime) between '" & CDate(StartTime) & "'"
          Sql = Sql & " and '" & CDate(EndTime) & "'"
          Sql = Sql & " and TRY_CAST([order date] AS datetime) = '" & Searchwhen & "'"
         Else
           See_Msg "Invalid Time", 5, ""
           Exit Sub
         End If
      Case Is = 2
          SearchWhat = "Server"
          Sql = "Select Distinct [Server Name] from [Closed Checks] where [order date] = '" & Format(Searchwhen, "mm/dd/yyyy") & "'"
          ConnectADORS Sql, sDb, SRS
             If Not SRS.EOF Then
              'List1.Clear
              Do While Not SRS.EOF
                List1.AddItem SRS![Server Name]
                SRS.MoveNext
              Loop
             End If
            DbClose sDb, SRS
           Exit Sub
      Case Is = 3
          SearchWhat = "Menu Items"
          Sql = "Select Distinct [Menu Item] from [Check Detail] where [check date] = '" & Format(Searchwhen, "mm/dd/yyyy") & "'"
          ConnectADORS Sql, sDb, SRS
             If Not SRS.EOF Then
              'List1.Clear
              Do While Not SRS.EOF
                List1.AddItem SRS![Menu Item]
                SRS.MoveNext
              Loop
             End If
            DbClose sDb, SRS
           Exit Sub
      Case Is = 4
         Sql = "Select * from [Closed Checks] Where [voided Sales] > 0"
          If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
             Sql = Sql & " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
          Else
             Sql = Sql & " and cdate([order Date] & ' ' & [Open time]) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
          End If
      Case Is = 5
          Scheck = ShowNumberPad("Enter Check No.", False, 0, 0)
          If Val(Scheck) = 0 Then Exit Sub
          Sql = "Select * from [Closed Checks] Where [Check Number] = " & Scheck
          If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
             Sql = Sql & " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
          Else
             Sql = Sql & " and cdate([order Date] & ' ' & [Open time]) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
          End If
      Case Is = 6
        Scheck = ShowKeyBoard("Enter Phone or Name")
        SearchWhat = "Customer Name"
        If Scheck = "" Then Exit Sub
        Sql = "Select * from [Closed Checks] where [order date] = '" & Format(Searchwhen, "mm/dd/yyyy") & "'"
        Sql = Sql & " and [Member Name] like '" & Scheck & "%'"
          ConnectADORS Sql, sDb, SRS
             If Not SRS.EOF Then
             'List1.Clear
             Do While Not SRS.EOF
               List1.AddItem SRS![Member Name] & "-CK#-" & SRS![Check Number]
               SRS.MoveNext
             Loop
            End If
           DbClose sDb, SRS
           Exit Sub
      Case Is = 7
        Scheck = ShowKeyBoard("Enter Table Or Tab Name")
        Sql = "Select * from [Closed Checks] Where [Table Number] like '" & Scheck & "%'"
        If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
           Sql = Sql & " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
        Else
           Sql = Sql & " and cdate([order Date] & ' ' & [Open time]) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
        End If
        PhoneSearch = Scheck
      Case Is = 8
        Scheck = ShowKeyBoard("Item Name")
        PhoneSearch = Scheck
        SearchWhat = "Menu Items"
        Sql = "Select * from [Check Detail] Where [Menu Item] like '" & Scheck & "%'"
        If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
           Sql = Sql & " and TRY_CAST([Check Date] AS datetime) + TRY_CAST([Recorded time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
        Else
           Sql = Sql & " and cdate([Check Date] & ' ' & [Recorded time]) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
        End If
        ConnectADORS Sql, sDb, SRS
         If Not SRS.EOF Then
           'List1.Clear
           Do While Not SRS.EOF
             List1.AddItem SRS![Menu Item]
             SRS.MoveNext
           Loop
           End If
           DbClose sDb, SRS
         Exit Sub
      Case Is = 9
        PhoneSearch = ShowKeyBoard("Enter Address")
        Scheck = GetDeliveryAddressCheck(PhoneSearch)
        If Val(Scheck) = 0 Then Exit Sub
        Sql = "Select * from [Closed Checks] Where [Check Number] = " & Scheck
        If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
           Sql = Sql & " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
        Else
           Sql = Sql & " and cdate([order Date] & ' ' & [Open time]) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
        End If
       Case Is = 10
          SearchWhat = "Payments"
          If Right$(UCase(FiletoLookat), 3) = "MDB" Then
           Sql = "Select Distinct [Payment Type] from [Payments] where TRY_CAST([date] AS datetime) + TRY_CAST([time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
          Else
           Sql = "Select Distinct [Payment Type] from [Payments] where TRY_CAST([date] AS datetime) + TRY_CAST([time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
          End If
          ConnectADORS Sql, sDb, SRS
             If Not SRS.EOF Then
            '  List1.Clear
              Do While Not SRS.EOF
                List1.AddItem SRS![Payment Type]
                SRS.MoveNext
              Loop
             End If
            DbClose sDb, SRS
           Exit Sub

   End Select

ConnectADORS Sql, sDb, SRS
   If Not SRS.EOF Then
      With SearchItem
          .rows = RecordsCount(SRS) + 1
           NR = 0
           .Clear
            .TextMatrix(0, 0) = "Chk No"
            .TextMatrix(0, 1) = "Type"
            .TextMatrix(0, 2) = "Name"
            .TextMatrix(0, 3) = "Time"
            .TextMatrix(0, 4) = "Amount"
           If System_config = ExpressService Then
              .TextMatrix(0, 1) = "Order"
           Else
              .TextMatrix(0, 1) = "Table"
           End If
            'SearchBox.Left = 1320 * WidScale: SearchBox.Top = 600 * TopScale
            CenterObject Me, SearchBox
            Do While Not SRS.EOF
              NR = NR + 1
              .TextMatrix(NR, 0) = SRS![Check Number]
            If System_config = ExpressService Then
              .TextMatrix(NR, 1) = SRS![Order Destination]
             Else
              .TextMatrix(NR, 1) = SRS![Table Number]
            End If
              .TextMatrix(NR, 2) = SRS![Server Name]
              .TextMatrix(NR, 3) = Format(SRS![Close Time], "HH:MM ampm")
              TotalSale = (SRS![Gross Sales] + SRS![Gross Sales Tax] + SRS![Gross Special Tax] + SRS![Gross Liquor Tax])
              TotalSale = TotalSale - SRS![Voided Sales] ' + (SRS![Voided Sales Tax]) + (SRS![Voided Special tax]) + (SRS![Voided liquor Tax]))
              TotalSale = TotalSale - (Abs(SRS![Gross Discounts]) - Abs(SRS![Voided Discounts]))
              .TextMatrix(NR, 4) = Format(TotalSale, "0.00")
              'View_Checks SRS![Check Number]
             SRS.MoveNext
            Loop
            SearchBox.Visible = True
            SearchBox.ZOrder 0
            Sleep 0
      End With
Else
  Call See_Msg("No Data Found", 5, "")
  PhoneSearch = ""
End If
DbClose sDb, SRS
End Sub
Private Sub SearchItem_Click()
'ToTDiff = 0
If SearchItem.Row < 0 Then Exit Sub
If Val(SearchItem.TextMatrix(SearchItem.Row, 0)) > 0 Then
   View_Checks Val(SearchItem.TextMatrix(SearchItem.Row, 0))
   'If Gross_Sales > 0 Or Chk_Sales_data(6) <> 0 Then
    CenterObject Me, ViewChk
    ViewChk.Visible = True
    ViewChk.ZOrder 0
    Counter = SearchItem.Row
    'Turned off for Dariy Queen Free Water.
   'End If
   Sleep 0
End If
End Sub

Private Sub ServerName_Click()
 On Error GoTo erview
' Label7.Visible = True
 Dim Cod As Integer   'Close out day
 Ticket(0).TopRow = 0
 Gross_Sales = 0
 Dim ItemTotal As Currency
 Chk_Sales_data(6) = 0
 Dim Db As ADODB.Connection, Records As ADODB.Recordset
 Dim Tb As ADODB.Recordset, RTB As ADODB.Recordset
 Dim RCount As Long
 ConnectADODB Search_Path & FiletoLookat, Db
 Sql = "Select * from [Closed Checks] Where [order date] = '" & Searchwhen & "' order by [Check number] asc"
 ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
   LoggedOnAs(0).EMPName = Tb![Server Name]
   Table_Number = Tb![Table Number]
   OrderIs = Tb![Order Destination]
   Check_no = Tb![Check Number]
   Number_In_Party = Tb![Number in Party]
   CheckDate = Trim$(Tb![Order Date])
   Cod = Tb![Close out day]
   CheckOpenTime = Tb![Open Time] & ""
   TimeCheckClosed.Caption = Format(Tb![Close Time], "HH:MM ampm")
   Ck_Totals.Caption = Format(Val(Tb![Check Total] & ""), "0.00")
   Cashier.Caption = Tb.fields("Cashier Name").Value & ""
ReDim_Array
Grid_array(0, 0, 3) = Format(Tb![Close Time], "HH:MM ampm")
Grid_array(0, 0, 4) = CheckDate
ItemTotal = 0
        Sql = "Select * From [Check Detail] where [Check Number] = " & Check_no
        Sql = Sql & " and [Close Out Day] = " & Cod
        Sql = Sql & " and [Servers Name] = '" & LoggedOnAs(0).EMPName & "'"
        Sql = Sql & " Order By Abs([Selection Number]),[Pos] asc "
        ConnectADORS Sql, Db, RTB
        If RTB.EOF Then
        MsgBox ("Search For Check Detail Failed.."), 0, "", 0, 0
        DbClose Db, RTB
        Chk_Sales_data(5) = 0
        Chk_Sales_data(6) = 0
        Exit Sub
        End If
        Dim Rw  As Integer
        Dim Cs  As Integer
        Rw = 0: Cs = RTB(10)
        Dim ShowMsg As Boolean
        ShowMsg = False
        Do While Not RTB.EOF
             If IsDate(CheckDate) And IsDate(RTB![Check Date]) Then
               If Abs(DateDiff("d", CheckDate, RTB![Check Date])) > 1 And ShowMsg = False Then
                 ShowMsg = True
                 See_Msg "More Than Seven Days of Data Have Been Found", 5, "Please Close System Every Week"
               End If
             End If 'Looks for more than skiped weekly close outs
             If RTB(2) = 0 Then
               Grid_array(0, Rw, 0) = ""
               Else
               Grid_array(0, Rw, 0) = RTB![Quantity]
               End If
               If RTB(8) > 0 Then
                 Grid_array(0, Rw, 1) = "void--> " & RTB![Menu Item]
               Else
                 Grid_array(0, Rw, 1) = RTB![Menu Item]
                 ItemTotal = ItemTotal + Val(RTB![Price] & "")
               End If
               If RTB![Price] = 0 Then
                 Grid_array(0, Rw, 2) = ""
               Else
                 Grid_array(0, Rw, 2) = Align(2, RTB![Price])
               End If
               Grid_array(0, Rw, 10) = RTB![Tax Table]
               Grid_array(0, Rw, 11) = RTB![Special Tax]
               Grid_array(0, Rw, 12) = RTB![Liquor Tax]
               Grid_array(0, Rw, 13) = RTB![voided]
               Grid_array(0, Rw, 14) = RTB![Selection Number]  'hold selection number for voided items
               Grid_array(0, Rw, 24) = RTB![Seat Number]
               Grid_array(0, Rw, 15) = RTB![Sort Level]
               Grid_array(0, Rw, 16) = RTB![Printed]
               Grid_array(0, Rw, 17) = RTB![Disc Type]
               Grid_array(0, Rw, 18) = RTB![Disc Amount]
               Grid_array(0, Rw, 19) = Grid_array(0, Rw, 0)
              '*****
               Grid_array(0, Rw, 25) = RTB![Menu Number]
               Grid_array(0, Rw, 26) = RTB![Department Report]
               Grid_array(0, Rw, 23) = RTB![Cateory Name]
               Grid_array(0, Rw, 30) = RTB![Menu Item]
               Grid_array(0, Rw, 32) = Val(RTB![Inclusive Tax] & "")
               Grid_array(0, Rw, 33) = RTB![Check Date] & "" & " " & RTB![Recorded Time] & ""
               'Grid_array(0, Rw, 34) = Val(rTb![Member Name] & "")
             Rw = Rw + 1
             If Rw >= MaxRows Then Redim_Grid
             NextRow(0) = Rw
           RTB.MoveNext
          Loop
SetNothing RTB
Sales_Totals
Ck_Totals.Caption = Align(7, Chk_Sales_data(5))
Tb.MoveNext
Loop
DbClose Db, RTB
skip:
Sleep 0
Exit Sub
erview:
MsgBox ("Error : " & Error), 0 + vbExclamation, "Error # : " & Err, 0, 0
Resume Next
End Sub
Private Sub SMC_Click()
On Error GoTo ERH
 If Trim$(SearchDate.List(SearchDate.ListIndex)) = "" Then Exit Sub
 Dim Db         As ADODB.Connection
 Dim Tb         As ADODB.Recordset
 Dim RS         As ADODB.Recordset
 Dim PaymentAmt As Currency
 Dim CloseAmt   As Currency
 Dim SQL1       As String
 
 SearchDate = SearchDate.List(SearchDate.ListIndex)
 If Left$(FiletoLookat, 3) <> "MDB" Then
   If FileCheck(Search_Path & FiletoLookat) = False Then
      FiletoLookat = "MSaleData"
      Sql = "Select Distinct [Check Number], sum([Payment Amount]) as [TPayment] from [Payments] where TRY_CAST([Date] AS date) = '" & SearchDate.List(SearchDate.ListIndex) & "'"
   End If
 Else
   Sql = "Select Distinct [Check Number], sum([Payment Amount]) as [TPayment] from [Payments] where TRY_CAST([Date] AS datetime) = '" & Format(CDate(SearchDate.List(SearchDate.ListIndex)), "yyyy-mm-dd hh:nn:ss") & "'"
 End If
 If ConnectADODB(Search_Path & FiletoLookat, Db) = False Then
    See_Msg "Error Connection To " & FiletoLookat, 5, ""
    Exit Sub
 End If
 Sql = Sql & " and [close out day] = " & SearchDay
 Sql = Sql & " Group by [Check Number]"
 ConnectADORS Sql, Db, Tb
  Do While Not Tb.EOF
    PaymentAmt = Tb.fields("TPayment").Value
    SQL1 = "Select * from [Closed Checks] Where [Check Number] = " & Tb.fields("Check Number").Value
   If FiletoLookat = "MSaleData" Then
    SQL1 = SQL1 & " and TRY_CAST([Order Date] AS date) = '" & SearchDate & "'"
   Else
    SQL1 = SQL1 & " and TRY_CAST([Order Date] AS datetime) = '" & Format(CDate(SearchDate), "yyyy-mm-dd hh:nn:ss") & "'"
   End If
   ConnectADORS SQL1, Db, RS
     If RS.EOF Then
       See_Msg "Closed Check Missing", 0, Tb.fields("Check Number").Value & " Payment " & Format(PaymentAmt, "0.00")
     Else
       If PaymentAmt <> RS.fields("Check Total").Value Then
         See_Msg Tb.fields("Check Number").Value & " Payment " & Format(PaymentAmt, "0.00"), 0, "Check total " & Format(RS.fields("Check Total").Value, "0.00")
       End If
     End If
   RS.Close
    SQL1 = "Select * from [Check Detail] Where [Check Number] = " & Tb.fields("Check Number").Value
    If FiletoLookat = "MSaleData" Then
     SQL1 = SQL1 & " and TRY_CAST([Check Date] AS date) = '" & SearchDate & "'"
    Else
     SQL1 = SQL1 & " and TRY_CAST([Check Date] AS datetime) = '" & Format(CDate(SearchDate), "yyyy-mm-dd hh:nn:ss") & "'"
    End If
    SQL1 = SQL1 & " and [Close out Day] = " & SearchDay
   ConnectADORS SQL1, Db, RS
     If RS.EOF Then
       See_Msg "Check Detail Missing", 0, Tb.fields("Check Number").Value
     End If
   RS.Close
  
  Tb.MoveNext
 Loop
 
 DbClose Db, Tb
 See_Msg "Search Complete", 5, "Done.."
 Exit Sub
ERH:
Error_Log Err.Number, Err.Description, "Searching lost checks", "None"
MsgBox ("Error : " & Error), 0 + vbExclamation, "Error # : " & Err, 0, 0
'Resume
End Sub

Private Sub SSCommand1_Click()
'********** PRINT CHECK *********
 On Error GoTo erview
PRINTIT:
Dim HoldValue           As Integer
Dim HoldDPValue         As Single
Dim HoldCreditAmount    As Single
HoldValue = Check_Option(132)
Check_Option(132) = 0
HoldDPValue = DualPricing
CashDiscountData.CreditAmount = 0: DualPricing = 0

If AllPrinter.Item("Receipt").PrinterPort = "None" Then
    SetPrinter 0
    PrintCLine StoreName, -1, 0
    PrintCLine StoreAddress, -1, 0
    PrintLine "", -1
    PrintLine "", 3500: PrintLine TimeCheckClosed, -1
    PrintLine "Server Name   " & ServerName.Caption, -1
    PrintLine "Check Number  " & CKNumb, -1
    PrintLine "Check Date    " & CheckDate.Caption, -1
    PrintLine "Check Time    " & TimeCheckClosed.Caption, -1
    DrawLine
    Dim R As Integer
    Dim C As Integer
    With Ticket(0)
     For R = 6 To .rows
        For C = 0 To 2
        If C = 1 And Trim(.TextMatrix(R - 6, C)) = "" Then GoTo EReport
         Select Case C
            Case Is = 0
               PrintLine .TextMatrix(R - 6, C), 1500
            Case Is = 1
               PrintLine .TextMatrix(R - 6, C), 4500
            Case Is = 2
               PrintLine .TextMatrix(R - 6, C), -1
         End Select
        Next C
    Next R
EReport:
    End With
    PrintLine "", -1
    DrawLine
    PrintLine "", -1
    PrintLine "Sales Tax", 4500
    PrintLine Format(Format(Chk_Sales_data(1) + Chk_Sales_data(3) + Chk_Sales_data(2), "0.00"), "0.00"), -1
    PrintLine "Check Total", 4500
    PrintLine Format(Ck_Totals.Caption, "0.00"), -1
    FinishJob 1
Else
  Sold_To = "Transaction Date " & CheckDate.Caption
  Grid_array(0, 0, 3) = Format(TimeCheckClosed.Caption, "HH:MM ampm")
  Grid_array(0, 0, 4) = CheckDate
  
  If Chk_Sales_data(6) < 0 Then
    If See_Msg("Print Voids On Ticket?", 4, "") = 6 Then
      Receipt "Voided Check", PrintReceipt, "Void", AbortSale
    Else
      UserVoid.Manager = "": UserVoid.Reason = "": UserVoid.Time = "": Emv_Back.r_Last4 = ""
      Receipt "Reprint of Closed Check", PrintReceipt, "Reprint", AbortSale
    End If
  Else
    Receipt "Reprint of Closed Check", PrintReceipt, "Reprint", AbortSale
  End If
  ExtraData = ""
  
  Sold_To = ""
  '//
  Check_Option(132) = HoldValue
  DualPricing = HoldDPValue
  CashDiscountData.CreditAmount = HoldCreditAmount
  '// stop double check total with dual pricing
End If
Exit Sub
erview:
MsgBox ("Error : " & Error), 0 + vbExclamation, "Error # : " & Err, 0, 0
Resume Next
End Sub
Private Sub SSCommand12_Click()
SearchBox.Visible = False
SearchBox.Left = -4000
Sleep 0
End Sub
Private Sub SSCommand13_Click()
SendAKey &H21, vbKeyHome
ViewChk.Visible = False
ViewChk.Left = -4000
End Sub
Private Sub SSCommand2_Click()
Close
Check_no = 0
Call MemoryCheck(Closed_Check)
End Sub
Private Sub SSCommand3_Click()
End Sub
Private Sub SSCommand6_Click()
Ticket(0).HighLight = flexHighlightNever
With Ticket(0)
  If .Row + 10 >= .rows Then Exit Sub
   .TopRow = .Row + 10
   .Row = .TopRow
End With
End Sub
Private Sub SSCommand7_Click()
Ticket(0).SetFocus
With Ticket(0)
  If .Row - 10 < 0 Then
   .TopRow = 0
  Else
   .TopRow = .Row - 10
   .Row = .TopRow
  End If
End With
End Sub
Private Function GetMDBDateTimeRange(DateIn As String, TableName As dbTableName) As String
  If TableName = closedchecks Then
   GetMDBDateTimeRange = " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(DateIn & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(DateIn) & " " & SetEndingTime) & "'"
  ElseIf TableName = checkdetail Then
    GetMDBDateTimeRange = " and TRY_CAST([Check Date] AS datetime) + TRY_CAST([Recorded time] AS datetime) between '" & CDate(DateIn & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(DateIn) & " " & SetEndingTime) & "'"
  Else
    GetMDBDateTimeRange = " and TRY_CAST([date] AS datetime) + TRY_CAST([time] AS datetime) between '" & CDate(DateIn & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(DateIn) & " " & SetEndingTime) & "'"
  End If
End Function

Private Function GetDateTimeRange(DateIn As String, TableName As dbTableName) As String
  If TableName = closedchecks Then
   GetDateTimeRange = " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(DateIn & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(DateIn) & " " & SetEndingTime) & "'"
  ElseIf TableName = checkdetail Then
    GetDateTimeRange = " and TRY_CAST([Check Date] AS datetime) + TRY_CAST([Recorded time] AS datetime) between '" & CDate(DateIn & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(DateIn) & " " & SetEndingTime) & "'"
  Else
    GetDateTimeRange = " and TRY_CAST([date] AS datetime) + TRY_CAST([time] AS datetime) between '" & CDate(DateIn & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(DateIn) & " " & SetEndingTime) & "'"
  End If
End Function
Public Sub View_Checks(FindCheckNo As Long)
 On Error GoTo erview
 Label2.Visible = True: Label6.Visible = True
 Dim Cod            As Integer
 Dim IncX           As Currency
 Dim VItem          As Currency
 Dim ItemTotal      As Currency
 Dim Db             As ADODB.Connection
 Dim Records        As ADODB.Recordset
 Dim Tb             As ADODB.Recordset
 Dim RTB            As ADODB.Recordset
 Dim FD             As ADODB.Connection
 Dim FR             As ADODB.Recordset
 Dim MemberName     As String
 Dim FindSvrCheck() As String
 UserVoid.Manager = "": UserVoid.Reason = "": UserVoid.Time = ""
ConnectADODB Search_Path & FiletoLookat, Db
Sql = "Select * from [Closed Checks] where [check number] = " & FindCheckNo    '& " and [order date] = '" & Searchwhen & "'"
If InStr(UCase(FiletoLookat), ".MDB") = 0 Then
   Sql = Sql & " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & CDate(Searchwhen & " " & GetStartTime) & "' and '" & CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime) & "'"
Else
   Sql = Sql & " and TRY_CAST([order Date] AS datetime) + TRY_CAST([Open time] AS datetime) between '" & Format(CDate(CDate(Searchwhen & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(Searchwhen) & " " & SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
End If
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
      See_Msg "Check Information Not Found", 5, "Can Not Review Check"
      DbClose Db, Tb
      Exit Sub
End If
Dim DTTTl   As Currency
Dim GSale   As Currency
Dim GSTx    As Currency
Dim GSPTx   As Currency
Dim GLTx    As Currency
Dim GDis    As Currency
Dim VSale   As Currency
Dim VSTx    As Currency
Dim VSPTx   As Currency
Dim VLTx    As Currency
Dim VDis    As Currency
Dim CkTTL   As Currency
Dim ICLTT   As Currency
Dim GEXTT   As Currency
Dim VEXTT   As Currency
Dim NSale   As Currency
Dim NTx     As Currency
Dim NSPTx   As Currency
Dim NDic    As Currency
Dim NEx     As Currency
Dim CKAMT   As Currency
Dim VDT     As Currency
Dim IDT     As Currency
Dim Row     As Integer
Dim RS      As Integer
Dim Cl      As Integer
Dim St      As Integer
Dim LP      As Integer
Dim CKDate  As String
Dim CkTime  As String
If Not Tb.EOF Then
   GSale = Tb![Gross Sales]
   GSTx = Tb![Gross Sales Tax]
   GSPTx = Tb![Gross Special Tax]
   GLTx = Tb![Gross Liquor Tax]
   GDis = Abs(Tb![Gross Discounts])
   
   VSale = Tb![Voided Sales]
   'VSTx = Tb![Voided Sales Tax]
   'VSPTx = Tb![Voided Special tax]
   'VLTx = Tb![Voided liquor Tax]
   VDis = Abs(Tb![Voided Discounts])
   
   ICLTT = Tb![Inclusive Total]
   
   GEXTT = Tb![Gross Exempt]
   VEXTT = Val(Tb![Voided Exempt] & "")
   
   
   NTx = GSTx '- VSTx
   NSPTx = GSPTx '- VSPTx
   NDic = GDis - VDis
   NEx = GEXTT - VEXTT
   NSale = (GSale - VSale) - NDic
   CKAMT = (((NSale + NTx) + NSPTx))
   LoggedOnAs(0).EMPName = Tb(0)
   Table_Number = Tb(1)
   OrderIs = Tb(18)
   Check_no = Tb(2)
   Number_In_Party = Tb(3)
   CheckDate = Trim$(Tb(19))
   Cod = Tb![Close out day]
   CheckOpenTime = Format(Tb![Open Time] & "", "HH:MM ampm")
   TimeCheckClosed.Caption = Format(CheckOpenTime, "HH:MM ampm")
   Ck_Totals.Caption = Format(Val(Tb![Check Total] & ""), "0.00")
   CkTime = Format(Tb.fields("Close Time").Value, "HH:MM ampm")
   CameraNo = Tb![Member Name] & ""
   MemberName = Tb![Member Name] & ""
   Cashier.Caption = Tb.fields("Cashier Name").Value & ""
   TBLNUmb.Caption = Table_Number
   ServerName.Caption = SearchItem.TextMatrix(SearchItem.Row, 2)
   CKNumb.Caption = Check_no
   PartyNumb.Caption = Number_In_Party
   CheckDate.Caption = Searchwhen
End If
ReDim_Array
        Sql = "Select * From [Check Detail] where [Check Number] = " & Check_no
        If Right$(UCase(FiletoLookat), 3) = "MDB" Then
           Sql = Sql & " and TRY_CAST([Check date] AS datetime) + TRY_CAST([Recorded Time] AS datetime) Between '" & Format(CDate(CDate(CheckDate & " " & CheckOpenTime)), "yyyy-mm-dd hh:nn:ss") & "'"
           Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(CheckDate) & " " & SetEndingTime())), "yyyy-mm-dd hh:nn:ss") & "'"
        Else
           Sql = Sql & " and TRY_CAST([Check date] AS datetime) + TRY_CAST([Recorded Time] AS datetime) Between '" & CDate(CheckDate & " " & GetStartTime) & "'"
           Sql = Sql & " and '" & CDate(SetEndingDate(CheckDate) & " " & SetEndingTime()) & "'"
        End If
        Sql = Sql & " and [Close Out Day] = " & Cod
        Sql = Sql & " and [Servers Name] = '" & LoggedOnAs(0).EMPName & "'"
        Sql = Sql & " Order By [Seat Number], Abs([Selection Number]),[Pos] asc "
        ConnectADORS Sql, Db, RTB
        If RTB.EOF Then
         MsgBox ("Search For Check Detail Failed.."), 0, "", 0, 0
         DbClose Db, RTB
         Exit Sub
        End If
        'CkTime = Format(CheckOpenTime, "HH:MM ampm")
        CKDate = CheckDate
        ItemTotal = 0: VItem = 0:   Row = 0: DTTTl = 0: IDT = 0: St = 0
        Do While Not RTB.EOF
              If St <> Val(RTB.fields("Seat Number").Value & "") And Val(RTB.fields("Voided").Value & "") = 0 Then
               St = Val(RTB.fields("Seat Number").Value & "")
               Row = 0
              End If
              If RTB![Quantity] = 0 Then
                Grid_array(St, Row, 0) = ""
              Else
                Grid_array(St, Row, 0) = RTB![Quantity]
              End If
              If RTB![voided] > 0 Then
                 Grid_array(St, Row, 1) = "void--> " & RTB![Menu Item]
                 If RTB![Price] > 0 Then
                  VItem = VItem + RTB![Price] - RTB![Inclusive Amount]
                 End If
                 UserVoid.Manager = RTB![Voided By] & ""
                 UserVoid.Reason = RTB![Void Reason] & ""
                 UserVoid.Time = RTB![Void Time] & ""
              Else
                 Grid_array(St, Row, 1) = RTB![Menu Item]
                 ItemTotal = ItemTotal + Val(RTB![Price] & "") - RTB![Inclusive Amount]
                 IncX = IncX + RTB![Inclusive Amount]
                 IDT = IDT + RTB![Inclusive Amount]
              End If
              If RTB![Price] = 0 Then
                 Grid_array(St, Row, 2) = ""
              Else
                 Grid_array(St, Row, 2) = Align(2, Val(RTB![Price] & ""))
                If RTB(8) = 0 Then DTTTl = DTTTl + Val(RTB![Price] & "") - RTB![Inclusive Amount]
              End If
               Grid_array(St, Row, 10) = RTB![Tax Table]
               Grid_array(St, Row, 11) = RTB![Special Tax]
               Grid_array(St, Row, 12) = RTB![Liquor Tax]
               Grid_array(St, Row, 13) = RTB![voided]
               Grid_array(St, Row, 14) = RTB![Selection Number]  'hold selection number for voided items
               Grid_array(St, Row, 24) = RTB![Seat Number]
               Grid_array(St, Row, 15) = RTB![Sort Level]
               Grid_array(St, Row, 16) = RTB![Printed]
               Grid_array(St, Row, 17) = RTB![Disc Type]
               Grid_array(St, Row, 18) = RTB![Disc Amount]
               Grid_array(St, Row, 19) = Grid_array(0, Row, 0)
               Grid_array(St, Row, 25) = RTB![Menu Number]
               Grid_array(St, Row, 26) = RTB![Department Report]
               Grid_array(St, Row, 23) = RTB![Cateory Name]
               Grid_array(St, Row, 30) = RTB![Menu Item]
               Grid_array(St, Row, 32) = Val(RTB![Inclusive Tax] & "")
               Grid_array(St, Row, 33) = RTB![Check Date] & "" & " " & RTB![Recorded Time] & ""
               Grid_array(St, Row, 34) = Val(RTB![Member Name] & "")
               Row = Row + 1
               NextRow(St) = NextRow(St) + 1
           RTB.MoveNext
          Loop
   SetNothing RTB
Sales_Totals

Ck_Totals.Caption = Format(Chk_Sales_data(5), "0.00")
STAX.Caption = Format((NTx + NSPTx) + GLTx, "0.00")
'If Format((NSale * 0.075), "0.00") <> NTx Then
' MsgBox NSale * 0.075, vbCritical, NTx
'End If
'Sql = "Select * from [Closed Checks] where [check number] = " & FindCheckNo  '& " and [order date] = '" & Searchwhen & "'"
'If Right$(UCase(FiletoLookat), 3) = "MDB" Then
' 'Sql = Sql & GetMDBDateTimeRange(CheckDate, closedchecks)
'Else
' Sql = Sql & GetDateTimeRange(CheckDate, closedchecks)
'End If
'Sql = Sql & " and [Close out day] = " & Cod
'Sql = Sql & " and [Server Name] = '" & LoggedOnAs(0).EMPName & "'"
'ConnectADORS Sql, Db, Tb '
'LN = LN + 1
'Dim ff As Integer
'ff = FreeFile
'If Not Tb.EOF Then
' If GetConfigInfo("Misc", "Check Balance", "OFF", "FUNCTIONS.INI") = "ENABLED" Then
'    Open SvrPath.LocalTerminalPath & "DAT Files\" & "OutPutData.Dat" For Append As #ff
'    Print #1, FindCheckNo & " ---- " & Tb![Register Name] & "   Line  " & LN
'    Print #1, "G Sales       " & Format(Tb(6), "0.00")
'    Print #1, "All Tax       " & Format(Tb(7), "0.00")
'    Print #1, "H Inc Tax     " & Format(Chk_Sales_data(12) + Chk_Sales_data(20), "0.00")
'    Print #1, "D Inc Tax     " & Format(IncX, "0.00")
'    Sql = "Select sum([Payment Amount]) as [Payment Amount] from [Payments] where [check number] = " & FindCheckNo & " and [Close out day] = " & Cod
'    'Sql = Sql & " and [date] = '" & CheckDate & "'"
'    ' will drop payments on checks started before Midnight
'    Print #1, "Total Sales   " & Format(Tb![Check Total], "0.00")
'    ConnectADORS Sql, Db, RTB
'    If Not RTB.EOF Then
'     Print #1, "Payment Amt   " & Format(Val(RTB![Payment Amount] & ""), "0.00")
'     If Val(RTB![Payment Amount] & "") <> Val(Tb![Check Total] & "") Then
'      Print #1, "No Match Payment  " & Format(Val(RTB![Payment Amount] & "") - Val(Tb![Check Total] & ""), "0.00")
'     End If
'    Else
'     Print #1, "No Payment Found "
'    End If
'    SetNothing RTB
'    Print #1, "Items         " & Format(ItemTotal, "0.00")
'    Print #1, "Net Sale      " & Format(((Tb(6) - Tb(11)) - (Abs(Tb(10)) - Abs(Tb(15)))), "0.00")
'    Print #1, "NetDisc       " & Format(Tb(10) - Tb(15), "0.00")
'    Print #1, "VoidDisc      " & Format(Tb(15), "0.00")
'    Print #1, "VoidH         " & Format(Tb(11), "0.00")  'header
'    Print #1, "VoidD         " & Format(VItem, "0.00")   'detail
'    If Tb(11) <> VItem Then
'    Print #1, "Void Detail No Match    " & Format(Tb(11) - VItem, "0.00")
'    End If
'    I_total = I_total + ItemTotal
'    NTotal = NTotal + Format(((Tb(6) - Tb(11)) - (Abs(Tb(10)) - Abs(Tb(15)))), "0.00")
'    GTotal = GTotal + Tb(6)
'    If ItemTotal <> Format(((Tb(6) - Tb(11)) - (Abs(Tb(10)) - Abs(Tb(15)))), "0.00") Then
'    Print #1, "No Match Check Balance  " & Format(ItemTotal - (Tb(6) - Tb(11)) - (Abs(Tb(10)) - Abs(Tb(15))), "0.00")
'    MsgBox ItemTotal, vbOKOnly, Format(((Tb(6) - Tb(11)) - (Abs(Tb(10)) - Abs(Tb(15)))), "0.00")
'    End If
'    Print #1, "GSale        >>> " & Format(GTotal, "0.00")   'detail
'    Print #1, "NSale        >>> " & Format(NTotal, "0.00")   'detail
'    Print #1, "ItemTTL      >>> " & Format(I_total, "0.00")
'    LastCheckNo = FindCheckNo
'    Close ff
'  End If
'End If
'SetNothing Tb
''*******************************************************
On Error GoTo skip
Sql = "Select * from [Payments] where [check number] = " & FindCheckNo & " and [Close out day] = " & Cod
If Right$(UCase(FiletoLookat), 3) = "MDB" Then
   Sql = Sql & " and TRY_CAST([date] AS datetime) + TRY_CAST([Time] AS datetime) Between '" & Format(CDate(CDate(CheckDate & " " & GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "'"
   Sql = Sql & " and '" & Format(CDate(CDate(SetEndingDate(CheckDate) & " " & SetEndingTime())), "yyyy-mm-dd hh:nn:ss") & "'"
Else
   Sql = Sql & " and TRY_CAST([date] AS datetime) + TRY_CAST([Time] AS datetime) Between '" & CDate(CheckDate & " " & GetStartTime) & "'"
   Sql = Sql & " and '" & CDate(SetEndingDate(CheckDate) & " " & SetEndingTime()) & "'"
End If
ConnectADORS Sql, Db, RTB
Grid_array(St, Row, 0) = "": Grid_array(St, Row, 1) = "": Grid_array(St, Row, 2) = ""
Row = Row + 1
Grid_array(St, Row, 0) = "": Grid_array(St, Row, 1) = "Check Payments": Grid_array(St, Row, 2) = ""
ExtraData = EC(2) & EC(9) & EC(3) & String(40, 45) & vbCrLf
ExtraData = ExtraData & "Check Payments" & vbCrLf
Row = Row + 1
If RTB.EOF Then
    Grid_array(St, Row, 0) = ""
    Grid_array(St, Row, 1) = "Payment Not Found"
    Grid_array(St, Row, 2) = ""
    See_Msg "Payment Not Found", 5, ""
    Row = Row + 1
    ExtraData = ExtraData & "Payment Not Found" & vbCrLf
End If
If GetConfigInfo("Functions", "Analize", "OFF", "FUNCTIONS.INI") = "ENABLED" Then
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Void Sale"
   Grid_array(St, Row, 2) = Format(VSale, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Void Detail"
   Grid_array(St, Row, 2) = Format(VItem, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Detail Total"
   Grid_array(St, Row, 2) = Format(DTTTl, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Net Total"
   Grid_array(St, Row, 2) = Format(NSale, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Inclusive Total"
   Grid_array(St, Row, 2) = Format(ICLTT, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Non Sales"
   Grid_array(St, Row, 2) = Format(GEXTT - VEXTT, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Check Total"
   Grid_array(St, Row, 2) = Format(CKAMT, "0.00")
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Inclusive Added"
   Grid_array(St, Row, 2) = Format(IDT + NSale, "0.00")
   TDVAT = (NSale - DTTTl)
   Row = Row + 1
   Grid_array(St, Row, 0) = ""
   Grid_array(St, Row, 1) = "Difference"
   Grid_array(St, Row, 2) = Format(TDVAT, "0.00")
End If
Row = Row + 1
ReDim How_Paid(0)
LP = 0
Do While Not RTB.EOF
    ReDim Preserve How_Paid(LP)
    Grid_array(St, Row, 0) = ""
    How_Paid(LP).PaymentName = RTB![Payment Type] & ""
    How_Paid(LP).PaymentAmt = Format(Val(RTB![Payment Amount] & ""), "0.00")
    Grid_array(St, Row, 1) = RTB![Payment Type] & ""
    Grid_array(St, Row, 2) = Format(Val(RTB![Payment Amount] & ""), "0.00")
    ExtraData = ExtraData & Grid_array(St, Row, 1) & "     "
    ExtraData = ExtraData & Grid_array(St, Row, 2) & vbCrLf
    If Val(RTB![Tip Amount]) > 0 Then
       Row = Row + 1
       LP = LP + 1
       ReDim Preserve How_Paid(LP)
       How_Paid(LP).PaymentName = "Tip Amount"
       How_Paid(LP).PaymentAmt = Format(Val(RTB![Tip Amount] & ""), "0.00")
       Grid_array(St, Row, 1) = "Tip Amount"
       Grid_array(St, Row, 2) = Format(Val(RTB![Tip Amount] & ""), "0.00")
       ExtraData = ExtraData & Grid_array(St, Row, 1) & "     "
       ExtraData = ExtraData & Grid_array(St, Row, 2) & vbCrLf
    End If
    Row = Row + 1
    LP = LP + 1
  RTB.MoveNext
Loop

If MemberName > "" Then  'added will print member Name on check or phone number for call in orders.
    If How_Paid(0).PaymentName <> MemberName Then
     Grid_array(St, Row, 0) = ""
     Grid_array(St, Row, 1) = MemberName
     Grid_array(St, Row, 2) = ""
     ExtraData = ExtraData & Grid_array(St, Row, 1) & vbCrLf
     Row = Row + 1
     If OrderIs = "Delivery" Then
      Grid_array(St, Row, 1) = GetDeliveryAddress(MemberName, FindCheckNo)
      ExtraData = ExtraData & Grid_array(St, Row, 1) & vbCrLf
      Row = Row + 1
     End If
    End If
End If
If IsDate(CkTime) And IsDate(CheckOpenTime) And IsDate(CKDate) Then
    Grid_array(St, Row, 0) = ""
    Grid_array(St, Row, 1) = "(Mins Open " & DateDiff("n", CDate(CKDate) & " " & CDate(CheckOpenTime), CDate(CKDate) & " " & CDate(CkTime)) & ")"
    ExtraData = ExtraData & Grid_array(St, Row, 1) & vbCrLf
    Row = Row + 1
End If
 ExtraData = ExtraData & EC(2) & EC(9) & EC(3) & String(40, 45) & vbCrLf
skip:
DbClose Db, RTB
NextRow(St) = Row
On Error GoTo erview
Ticket(0).HighLight = flexHighlightNever
Row = 0
Ticket(0).TopRow = 0
With Ticket(0)
 .Clear
 For St = 0 To Number_In_Party - 1
  For RS = 0 To NextRow(St) - 1
   For Cl = 0 To TTL_Col
      If Cl = 2 And Val(Grid_array(St, RS, Cl)) < 0 Then
         .Cell(flexcpForeColor, Row, 0, Row, 2) = &H800080
      End If
      If Cl = 0 And Val(Grid_array(St, RS, 0)) = 0 And Trim(Grid_array(St, RS, 1)) > "" Then
         .Cell(flexcpForeColor, Row, 0, Row, 2) = &H808080
      End If
      If Cl = 1 And Trim(Grid_array(St, RS, 1)) = ">> Check Open Mins ->" Then
         .Cell(flexcpForeColor, Row, 0, Row, 2) = vbRed
      End If
      If ByTime.Value = 1 Then
         .TextMatrix(RS, Cl) = Grid_array(St, RS, Cl)
         If IsDate(Grid_array(St, RS, 33)) And IsDate(CheckOpenTime) And IsDate(CKDate) Then
           .TextMatrix(Row, 2) = "Min( " & DateDiff("n", CDate(CKDate) & " " & CDate(CheckOpenTime), CDate(Grid_array(St, RS, 33))) & " )"
         End If
      Else
         .TextMatrix(Row, Cl) = Grid_array(St, RS, Cl)
      End If
   Next Cl
   Row = Row + 1
  Next RS
 Next St
End With
Ticket(0).HighLight = flexHighlightNever
Sleep 0
Exit Sub
erview:
MsgBox ("Error : " & Error), 0 + vbExclamation, "Error # : " & Err, 0, 0
Resume Next
End Sub
Private Sub SSCommand9_Click()
End Sub
Private Sub StartTime_Click()
Dim St As Datetime.DateOrTime
Set St = New Datetime.DateOrTime
StartTime.Text = St.ReturnTime
Set St = Nothing
End Sub
'Private Sub GetVideoClip(DateTimeSTR As String, CamNo As Long)
'Dim UTC_second As String
'Dim ChannelNo  As String
'Dim GMT_Local  As String
'Dim UTC_Time   As String
'GetSystemTime SysTime
'UTC_Time = SysTime.wMonth & "/" & SysTime.wDay & "/" & SysTime.wYear & " " & SysTime.wHour & ":" & SysTime.wMinute & ":" & SysTime.wSecond
'GMT_Local = DateDiff("s", Now, CDate(UTC_Time))
'If CamNo - 1 < 0 Then
'    ChannelNo = "01"
'Else
'    If (CamNo - 1) < 10 Then
'      ChannelNo = "0" & CStr(CamNo - 1)
'    Else
'      ChannelNo = CStr(CamNo - 1)
'    End If
'End If
'UTC_second = CStr(DateDiff("s", "01/01/1970 00:00:00", CDate(DateTimeSTR)))
'If RemoteCamera.CameraReplay > "" Then
' Shell RemoteCamera.CameraReplay & " " & CStr(Val(UTC_second) + Val(GMT_Local)) & "_" & ChannelNo, vbNormalFocus
'Else
' See_Msg "No location for replay set", 5, ""
'End If
'End Sub

Private Sub STS_Click()
Dim SL As ADODB.Connection
Dim SRd As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & ClosingHistory, SL
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
    SummaryServeTimes SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
  End If
Else
   See_Msg "Not Dates Found", 5, ""
End If

End Sub

Private Sub Ticket_Click(Index As Integer)
On Error GoTo ERH
'If RemoteCamera.CameraIP > "" Then
'  GetVideoClip Ticket(0).TextMatrix(Ticket(0).Row, 33), Ticket(0).TextMatrix(Ticket(0).Row, 34)
'End If
ERH:
Err.Clear
End Sub

