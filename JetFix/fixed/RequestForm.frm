VERSION 5.00
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form RequestListForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "RequestForm"
   ClientHeight    =   8625
   ClientLeft      =   540
   ClientTop       =   735
   ClientWidth     =   14325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BTNENHLib4.BtnEnh Exit 
      Height          =   855
      Left            =   10920
      TabIndex        =   0
      Top             =   7560
      Width           =   3015
      _Version        =   589828
      _ExtentX        =   5318
      _ExtentY        =   1508
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
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":0000
      textLT          =   "RequestForm.frx":0068
      textCT          =   "RequestForm.frx":0080
      textRT          =   "RequestForm.frx":0098
      textLM          =   "RequestForm.frx":00B0
      textRM          =   "RequestForm.frx":00C8
      textLB          =   "RequestForm.frx":00E0
      textCB          =   "RequestForm.frx":00F8
      textRB          =   "RequestForm.frx":0110
      colorBack       =   "RequestForm.frx":0128
      colorIntern     =   "RequestForm.frx":0152
      colorMO         =   "RequestForm.frx":017C
      colorFocus      =   "RequestForm.frx":01A6
      colorDisabled   =   "RequestForm.frx":01D0
      colorPressed    =   "RequestForm.frx":01FA
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
   Begin BTNENHLib4.BtnEnh UpArrow 
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _Version        =   589828
      _ExtentX        =   2566
      _ExtentY        =   1508
      _StockProps     =   66
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
      Picture         =   "C:\Users\ben\Downloads\double_chevron_up_32x32_variant.bmp"
      PictureYOffset  =   3
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPicture      =   "RequestForm.frx":0224
      textCaption     =   "RequestForm.frx":02CD
      textLT          =   "RequestForm.frx":02E5
      textCT          =   "RequestForm.frx":02FD
      textRT          =   "RequestForm.frx":0315
      textLM          =   "RequestForm.frx":032D
      textRM          =   "RequestForm.frx":0345
      textLB          =   "RequestForm.frx":035D
      textCB          =   "RequestForm.frx":0375
      textRB          =   "RequestForm.frx":038D
      colorBack       =   "RequestForm.frx":03A5
      colorIntern     =   "RequestForm.frx":03CF
      colorMO         =   "RequestForm.frx":03F9
      colorFocus      =   "RequestForm.frx":0423
      colorDisabled   =   "RequestForm.frx":044D
      colorPressed    =   "RequestForm.frx":0477
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
   Begin BTNENHLib4.BtnEnh DwnArrow 
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _Version        =   589828
      _ExtentX        =   2566
      _ExtentY        =   1508
      _StockProps     =   66
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
      Picture         =   "C:\Users\ben\Downloads\double_chevron_down_32x32.bmp"
      PictureYOffset  =   -3
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPicture      =   "RequestForm.frx":04A1
      textCaption     =   "RequestForm.frx":053E
      textLT          =   "RequestForm.frx":0556
      textCT          =   "RequestForm.frx":056E
      textRT          =   "RequestForm.frx":0586
      textLM          =   "RequestForm.frx":059E
      textRM          =   "RequestForm.frx":05B6
      textLB          =   "RequestForm.frx":05CE
      textCB          =   "RequestForm.frx":05E6
      textRB          =   "RequestForm.frx":05FE
      colorBack       =   "RequestForm.frx":0616
      colorIntern     =   "RequestForm.frx":0640
      colorMO         =   "RequestForm.frx":066A
      colorFocus      =   "RequestForm.frx":0694
      colorDisabled   =   "RequestForm.frx":06BE
      colorPressed    =   "RequestForm.frx":06E8
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":0712
      textLT          =   "RequestForm.frx":077A
      textCT          =   "RequestForm.frx":0792
      textRT          =   "RequestForm.frx":07AA
      textLM          =   "RequestForm.frx":07C2
      textRM          =   "RequestForm.frx":07DA
      textLB          =   "RequestForm.frx":07F2
      textCB          =   "RequestForm.frx":080A
      textRB          =   "RequestForm.frx":0822
      colorBack       =   "RequestForm.frx":083A
      colorIntern     =   "RequestForm.frx":0864
      colorMO         =   "RequestForm.frx":088E
      colorFocus      =   "RequestForm.frx":08B8
      colorDisabled   =   "RequestForm.frx":08E2
      colorPressed    =   "RequestForm.frx":090C
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SSCommand12 
      Height          =   855
      Left            =   7800
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   3015
      _Version        =   589828
      _ExtentX        =   5318
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Find Check"
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
      textCaption     =   "RequestForm.frx":0936
      textLT          =   "RequestForm.frx":09AA
      textCT          =   "RequestForm.frx":09C2
      textRT          =   "RequestForm.frx":09DA
      textLM          =   "RequestForm.frx":09F2
      textRM          =   "RequestForm.frx":0A0A
      textLB          =   "RequestForm.frx":0A22
      textCB          =   "RequestForm.frx":0A3A
      textRB          =   "RequestForm.frx":0A52
      colorBack       =   "RequestForm.frx":0A6A
      colorIntern     =   "RequestForm.frx":0A94
      colorMO         =   "RequestForm.frx":0ABE
      colorFocus      =   "RequestForm.frx":0AE8
      colorDisabled   =   "RequestForm.frx":0B12
      colorPressed    =   "RequestForm.frx":0B3C
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
      Office2003ColorNorth=   16384
      Office2003ColorSouth=   16384
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   16384
      Office2003ColorSouthMouseOver=   16384
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   16384
      Office2003ColorSouthPressed=   16384
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":0B66
      textLT          =   "RequestForm.frx":0BCE
      textCT          =   "RequestForm.frx":0BE6
      textRT          =   "RequestForm.frx":0BFE
      textLM          =   "RequestForm.frx":0C16
      textRM          =   "RequestForm.frx":0C2E
      textLB          =   "RequestForm.frx":0C46
      textCB          =   "RequestForm.frx":0C5E
      textRB          =   "RequestForm.frx":0C76
      colorBack       =   "RequestForm.frx":0C8E
      colorIntern     =   "RequestForm.frx":0CB8
      colorMO         =   "RequestForm.frx":0CE2
      colorFocus      =   "RequestForm.frx":0D0C
      colorDisabled   =   "RequestForm.frx":0D36
      colorPressed    =   "RequestForm.frx":0D60
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":0D8A
      textLT          =   "RequestForm.frx":0DF2
      textCT          =   "RequestForm.frx":0E0A
      textRT          =   "RequestForm.frx":0E22
      textLM          =   "RequestForm.frx":0E3A
      textRM          =   "RequestForm.frx":0E52
      textLB          =   "RequestForm.frx":0E6A
      textCB          =   "RequestForm.frx":0E82
      textRB          =   "RequestForm.frx":0E9A
      colorBack       =   "RequestForm.frx":0EB2
      colorIntern     =   "RequestForm.frx":0EDC
      colorMO         =   "RequestForm.frx":0F06
      colorFocus      =   "RequestForm.frx":0F30
      colorDisabled   =   "RequestForm.frx":0F5A
      colorPressed    =   "RequestForm.frx":0F84
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":0FAE
      textLT          =   "RequestForm.frx":1016
      textCT          =   "RequestForm.frx":102E
      textRT          =   "RequestForm.frx":1046
      textLM          =   "RequestForm.frx":105E
      textRM          =   "RequestForm.frx":1076
      textLB          =   "RequestForm.frx":108E
      textCB          =   "RequestForm.frx":10A6
      textRB          =   "RequestForm.frx":10BE
      colorBack       =   "RequestForm.frx":10D6
      colorIntern     =   "RequestForm.frx":1100
      colorMO         =   "RequestForm.frx":112A
      colorFocus      =   "RequestForm.frx":1154
      colorDisabled   =   "RequestForm.frx":117E
      colorPressed    =   "RequestForm.frx":11A8
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":11D2
      textLT          =   "RequestForm.frx":123A
      textCT          =   "RequestForm.frx":1252
      textRT          =   "RequestForm.frx":126A
      textLM          =   "RequestForm.frx":1282
      textRM          =   "RequestForm.frx":129A
      textLB          =   "RequestForm.frx":12B2
      textCB          =   "RequestForm.frx":12CA
      textRB          =   "RequestForm.frx":12E2
      colorBack       =   "RequestForm.frx":12FA
      colorIntern     =   "RequestForm.frx":1324
      colorMO         =   "RequestForm.frx":134E
      colorFocus      =   "RequestForm.frx":1378
      colorDisabled   =   "RequestForm.frx":13A2
      colorPressed    =   "RequestForm.frx":13CC
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   5
      Left            =   480
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":13F6
      textLT          =   "RequestForm.frx":145E
      textCT          =   "RequestForm.frx":1476
      textRT          =   "RequestForm.frx":148E
      textLM          =   "RequestForm.frx":14A6
      textRM          =   "RequestForm.frx":14BE
      textLB          =   "RequestForm.frx":14D6
      textCB          =   "RequestForm.frx":14EE
      textRB          =   "RequestForm.frx":1506
      colorBack       =   "RequestForm.frx":151E
      colorIntern     =   "RequestForm.frx":1548
      colorMO         =   "RequestForm.frx":1572
      colorFocus      =   "RequestForm.frx":159C
      colorDisabled   =   "RequestForm.frx":15C6
      colorPressed    =   "RequestForm.frx":15F0
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   6
      Left            =   480
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":161A
      textLT          =   "RequestForm.frx":1682
      textCT          =   "RequestForm.frx":169A
      textRT          =   "RequestForm.frx":16B2
      textLM          =   "RequestForm.frx":16CA
      textRM          =   "RequestForm.frx":16E2
      textLB          =   "RequestForm.frx":16FA
      textCB          =   "RequestForm.frx":1712
      textRB          =   "RequestForm.frx":172A
      colorBack       =   "RequestForm.frx":1742
      colorIntern     =   "RequestForm.frx":176C
      colorMO         =   "RequestForm.frx":1796
      colorFocus      =   "RequestForm.frx":17C0
      colorDisabled   =   "RequestForm.frx":17EA
      colorPressed    =   "RequestForm.frx":1814
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   7
      Left            =   3840
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":183E
      textLT          =   "RequestForm.frx":18A6
      textCT          =   "RequestForm.frx":18BE
      textRT          =   "RequestForm.frx":18D6
      textLM          =   "RequestForm.frx":18EE
      textRM          =   "RequestForm.frx":1906
      textLB          =   "RequestForm.frx":191E
      textCB          =   "RequestForm.frx":1936
      textRB          =   "RequestForm.frx":194E
      colorBack       =   "RequestForm.frx":1966
      colorIntern     =   "RequestForm.frx":1990
      colorMO         =   "RequestForm.frx":19BA
      colorFocus      =   "RequestForm.frx":19E4
      colorDisabled   =   "RequestForm.frx":1A0E
      colorPressed    =   "RequestForm.frx":1A38
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   8
      Left            =   3840
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":1A62
      textLT          =   "RequestForm.frx":1ACA
      textCT          =   "RequestForm.frx":1AE2
      textRT          =   "RequestForm.frx":1AFA
      textLM          =   "RequestForm.frx":1B12
      textRM          =   "RequestForm.frx":1B2A
      textLB          =   "RequestForm.frx":1B42
      textCB          =   "RequestForm.frx":1B5A
      textRB          =   "RequestForm.frx":1B72
      colorBack       =   "RequestForm.frx":1B8A
      colorIntern     =   "RequestForm.frx":1BB4
      colorMO         =   "RequestForm.frx":1BDE
      colorFocus      =   "RequestForm.frx":1C08
      colorDisabled   =   "RequestForm.frx":1C32
      colorPressed    =   "RequestForm.frx":1C5C
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   9
      Left            =   3840
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":1C86
      textLT          =   "RequestForm.frx":1CEE
      textCT          =   "RequestForm.frx":1D06
      textRT          =   "RequestForm.frx":1D1E
      textLM          =   "RequestForm.frx":1D36
      textRM          =   "RequestForm.frx":1D4E
      textLB          =   "RequestForm.frx":1D66
      textCB          =   "RequestForm.frx":1D7E
      textRB          =   "RequestForm.frx":1D96
      colorBack       =   "RequestForm.frx":1DAE
      colorIntern     =   "RequestForm.frx":1DD8
      colorMO         =   "RequestForm.frx":1E02
      colorFocus      =   "RequestForm.frx":1E2C
      colorDisabled   =   "RequestForm.frx":1E56
      colorPressed    =   "RequestForm.frx":1E80
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   10
      Left            =   3840
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":1EAA
      textLT          =   "RequestForm.frx":1F12
      textCT          =   "RequestForm.frx":1F2A
      textRT          =   "RequestForm.frx":1F42
      textLM          =   "RequestForm.frx":1F5A
      textRM          =   "RequestForm.frx":1F72
      textLB          =   "RequestForm.frx":1F8A
      textCB          =   "RequestForm.frx":1FA2
      textRB          =   "RequestForm.frx":1FBA
      colorBack       =   "RequestForm.frx":1FD2
      colorIntern     =   "RequestForm.frx":1FFC
      colorMO         =   "RequestForm.frx":2026
      colorFocus      =   "RequestForm.frx":2050
      colorDisabled   =   "RequestForm.frx":207A
      colorPressed    =   "RequestForm.frx":20A4
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   11
      Left            =   3840
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":20CE
      textLT          =   "RequestForm.frx":2136
      textCT          =   "RequestForm.frx":214E
      textRT          =   "RequestForm.frx":2166
      textLM          =   "RequestForm.frx":217E
      textRM          =   "RequestForm.frx":2196
      textLB          =   "RequestForm.frx":21AE
      textCB          =   "RequestForm.frx":21C6
      textRB          =   "RequestForm.frx":21DE
      colorBack       =   "RequestForm.frx":21F6
      colorIntern     =   "RequestForm.frx":2220
      colorMO         =   "RequestForm.frx":224A
      colorFocus      =   "RequestForm.frx":2274
      colorDisabled   =   "RequestForm.frx":229E
      colorPressed    =   "RequestForm.frx":22C8
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   12
      Left            =   3840
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":22F2
      textLT          =   "RequestForm.frx":235A
      textCT          =   "RequestForm.frx":2372
      textRT          =   "RequestForm.frx":238A
      textLM          =   "RequestForm.frx":23A2
      textRM          =   "RequestForm.frx":23BA
      textLB          =   "RequestForm.frx":23D2
      textCB          =   "RequestForm.frx":23EA
      textRB          =   "RequestForm.frx":2402
      colorBack       =   "RequestForm.frx":241A
      colorIntern     =   "RequestForm.frx":2444
      colorMO         =   "RequestForm.frx":246E
      colorFocus      =   "RequestForm.frx":2498
      colorDisabled   =   "RequestForm.frx":24C2
      colorPressed    =   "RequestForm.frx":24EC
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   13
      Left            =   3840
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":2516
      textLT          =   "RequestForm.frx":257E
      textCT          =   "RequestForm.frx":2596
      textRT          =   "RequestForm.frx":25AE
      textLM          =   "RequestForm.frx":25C6
      textRM          =   "RequestForm.frx":25DE
      textLB          =   "RequestForm.frx":25F6
      textCB          =   "RequestForm.frx":260E
      textRB          =   "RequestForm.frx":2626
      colorBack       =   "RequestForm.frx":263E
      colorIntern     =   "RequestForm.frx":2668
      colorMO         =   "RequestForm.frx":2692
      colorFocus      =   "RequestForm.frx":26BC
      colorDisabled   =   "RequestForm.frx":26E6
      colorPressed    =   "RequestForm.frx":2710
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   14
      Left            =   7200
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":273A
      textLT          =   "RequestForm.frx":27A2
      textCT          =   "RequestForm.frx":27BA
      textRT          =   "RequestForm.frx":27D2
      textLM          =   "RequestForm.frx":27EA
      textRM          =   "RequestForm.frx":2802
      textLB          =   "RequestForm.frx":281A
      textCB          =   "RequestForm.frx":2832
      textRB          =   "RequestForm.frx":284A
      colorBack       =   "RequestForm.frx":2862
      colorIntern     =   "RequestForm.frx":288C
      colorMO         =   "RequestForm.frx":28B6
      colorFocus      =   "RequestForm.frx":28E0
      colorDisabled   =   "RequestForm.frx":290A
      colorPressed    =   "RequestForm.frx":2934
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   15
      Left            =   7200
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":295E
      textLT          =   "RequestForm.frx":29C6
      textCT          =   "RequestForm.frx":29DE
      textRT          =   "RequestForm.frx":29F6
      textLM          =   "RequestForm.frx":2A0E
      textRM          =   "RequestForm.frx":2A26
      textLB          =   "RequestForm.frx":2A3E
      textCB          =   "RequestForm.frx":2A56
      textRB          =   "RequestForm.frx":2A6E
      colorBack       =   "RequestForm.frx":2A86
      colorIntern     =   "RequestForm.frx":2AB0
      colorMO         =   "RequestForm.frx":2ADA
      colorFocus      =   "RequestForm.frx":2B04
      colorDisabled   =   "RequestForm.frx":2B2E
      colorPressed    =   "RequestForm.frx":2B58
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   16
      Left            =   7200
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":2B82
      textLT          =   "RequestForm.frx":2BEA
      textCT          =   "RequestForm.frx":2C02
      textRT          =   "RequestForm.frx":2C1A
      textLM          =   "RequestForm.frx":2C32
      textRM          =   "RequestForm.frx":2C4A
      textLB          =   "RequestForm.frx":2C62
      textCB          =   "RequestForm.frx":2C7A
      textRB          =   "RequestForm.frx":2C92
      colorBack       =   "RequestForm.frx":2CAA
      colorIntern     =   "RequestForm.frx":2CD4
      colorMO         =   "RequestForm.frx":2CFE
      colorFocus      =   "RequestForm.frx":2D28
      colorDisabled   =   "RequestForm.frx":2D52
      colorPressed    =   "RequestForm.frx":2D7C
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   17
      Left            =   7200
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":2DA6
      textLT          =   "RequestForm.frx":2E0E
      textCT          =   "RequestForm.frx":2E26
      textRT          =   "RequestForm.frx":2E3E
      textLM          =   "RequestForm.frx":2E56
      textRM          =   "RequestForm.frx":2E6E
      textLB          =   "RequestForm.frx":2E86
      textCB          =   "RequestForm.frx":2E9E
      textRB          =   "RequestForm.frx":2EB6
      colorBack       =   "RequestForm.frx":2ECE
      colorIntern     =   "RequestForm.frx":2EF8
      colorMO         =   "RequestForm.frx":2F22
      colorFocus      =   "RequestForm.frx":2F4C
      colorDisabled   =   "RequestForm.frx":2F76
      colorPressed    =   "RequestForm.frx":2FA0
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   18
      Left            =   7200
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":2FCA
      textLT          =   "RequestForm.frx":3032
      textCT          =   "RequestForm.frx":304A
      textRT          =   "RequestForm.frx":3062
      textLM          =   "RequestForm.frx":307A
      textRM          =   "RequestForm.frx":3092
      textLB          =   "RequestForm.frx":30AA
      textCB          =   "RequestForm.frx":30C2
      textRB          =   "RequestForm.frx":30DA
      colorBack       =   "RequestForm.frx":30F2
      colorIntern     =   "RequestForm.frx":311C
      colorMO         =   "RequestForm.frx":3146
      colorFocus      =   "RequestForm.frx":3170
      colorDisabled   =   "RequestForm.frx":319A
      colorPressed    =   "RequestForm.frx":31C4
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   19
      Left            =   7200
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":31EE
      textLT          =   "RequestForm.frx":3256
      textCT          =   "RequestForm.frx":326E
      textRT          =   "RequestForm.frx":3286
      textLM          =   "RequestForm.frx":329E
      textRM          =   "RequestForm.frx":32B6
      textLB          =   "RequestForm.frx":32CE
      textCB          =   "RequestForm.frx":32E6
      textRB          =   "RequestForm.frx":32FE
      colorBack       =   "RequestForm.frx":3316
      colorIntern     =   "RequestForm.frx":3340
      colorMO         =   "RequestForm.frx":336A
      colorFocus      =   "RequestForm.frx":3394
      colorDisabled   =   "RequestForm.frx":33BE
      colorPressed    =   "RequestForm.frx":33E8
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   20
      Left            =   7200
      TabIndex        =   26
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":3412
      textLT          =   "RequestForm.frx":347A
      textCT          =   "RequestForm.frx":3492
      textRT          =   "RequestForm.frx":34AA
      textLM          =   "RequestForm.frx":34C2
      textRM          =   "RequestForm.frx":34DA
      textLB          =   "RequestForm.frx":34F2
      textCB          =   "RequestForm.frx":350A
      textRB          =   "RequestForm.frx":3522
      colorBack       =   "RequestForm.frx":353A
      colorIntern     =   "RequestForm.frx":3564
      colorMO         =   "RequestForm.frx":358E
      colorFocus      =   "RequestForm.frx":35B8
      colorDisabled   =   "RequestForm.frx":35E2
      colorPressed    =   "RequestForm.frx":360C
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   21
      Left            =   10560
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":3636
      textLT          =   "RequestForm.frx":369E
      textCT          =   "RequestForm.frx":36B6
      textRT          =   "RequestForm.frx":36CE
      textLM          =   "RequestForm.frx":36E6
      textRM          =   "RequestForm.frx":36FE
      textLB          =   "RequestForm.frx":3716
      textCB          =   "RequestForm.frx":372E
      textRB          =   "RequestForm.frx":3746
      colorBack       =   "RequestForm.frx":375E
      colorIntern     =   "RequestForm.frx":3788
      colorMO         =   "RequestForm.frx":37B2
      colorFocus      =   "RequestForm.frx":37DC
      colorDisabled   =   "RequestForm.frx":3806
      colorPressed    =   "RequestForm.frx":3830
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   22
      Left            =   10560
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":385A
      textLT          =   "RequestForm.frx":38C2
      textCT          =   "RequestForm.frx":38DA
      textRT          =   "RequestForm.frx":38F2
      textLM          =   "RequestForm.frx":390A
      textRM          =   "RequestForm.frx":3922
      textLB          =   "RequestForm.frx":393A
      textCB          =   "RequestForm.frx":3952
      textRB          =   "RequestForm.frx":396A
      colorBack       =   "RequestForm.frx":3982
      colorIntern     =   "RequestForm.frx":39AC
      colorMO         =   "RequestForm.frx":39D6
      colorFocus      =   "RequestForm.frx":3A00
      colorDisabled   =   "RequestForm.frx":3A2A
      colorPressed    =   "RequestForm.frx":3A54
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   23
      Left            =   10560
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":3A7E
      textLT          =   "RequestForm.frx":3AE6
      textCT          =   "RequestForm.frx":3AFE
      textRT          =   "RequestForm.frx":3B16
      textLM          =   "RequestForm.frx":3B2E
      textRM          =   "RequestForm.frx":3B46
      textLB          =   "RequestForm.frx":3B5E
      textCB          =   "RequestForm.frx":3B76
      textRB          =   "RequestForm.frx":3B8E
      colorBack       =   "RequestForm.frx":3BA6
      colorIntern     =   "RequestForm.frx":3BD0
      colorMO         =   "RequestForm.frx":3BFA
      colorFocus      =   "RequestForm.frx":3C24
      colorDisabled   =   "RequestForm.frx":3C4E
      colorPressed    =   "RequestForm.frx":3C78
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   24
      Left            =   10560
      TabIndex        =   30
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":3CA2
      textLT          =   "RequestForm.frx":3D0A
      textCT          =   "RequestForm.frx":3D22
      textRT          =   "RequestForm.frx":3D3A
      textLM          =   "RequestForm.frx":3D52
      textRM          =   "RequestForm.frx":3D6A
      textLB          =   "RequestForm.frx":3D82
      textCB          =   "RequestForm.frx":3D9A
      textRB          =   "RequestForm.frx":3DB2
      colorBack       =   "RequestForm.frx":3DCA
      colorIntern     =   "RequestForm.frx":3DF4
      colorMO         =   "RequestForm.frx":3E1E
      colorFocus      =   "RequestForm.frx":3E48
      colorDisabled   =   "RequestForm.frx":3E72
      colorPressed    =   "RequestForm.frx":3E9C
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   25
      Left            =   10560
      TabIndex        =   31
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":3EC6
      textLT          =   "RequestForm.frx":3F2E
      textCT          =   "RequestForm.frx":3F46
      textRT          =   "RequestForm.frx":3F5E
      textLM          =   "RequestForm.frx":3F76
      textRM          =   "RequestForm.frx":3F8E
      textLB          =   "RequestForm.frx":3FA6
      textCB          =   "RequestForm.frx":3FBE
      textRB          =   "RequestForm.frx":3FD6
      colorBack       =   "RequestForm.frx":3FEE
      colorIntern     =   "RequestForm.frx":4018
      colorMO         =   "RequestForm.frx":4042
      colorFocus      =   "RequestForm.frx":406C
      colorDisabled   =   "RequestForm.frx":4096
      colorPressed    =   "RequestForm.frx":40C0
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   26
      Left            =   10560
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":40EA
      textLT          =   "RequestForm.frx":4152
      textCT          =   "RequestForm.frx":416A
      textRT          =   "RequestForm.frx":4182
      textLM          =   "RequestForm.frx":419A
      textRM          =   "RequestForm.frx":41B2
      textLB          =   "RequestForm.frx":41CA
      textCB          =   "RequestForm.frx":41E2
      textRB          =   "RequestForm.frx":41FA
      colorBack       =   "RequestForm.frx":4212
      colorIntern     =   "RequestForm.frx":423C
      colorMO         =   "RequestForm.frx":4266
      colorFocus      =   "RequestForm.frx":4290
      colorDisabled   =   "RequestForm.frx":42BA
      colorPressed    =   "RequestForm.frx":42E4
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin BTNENHLib4.BtnEnh SelectionButton 
      Height          =   975
      Index           =   27
      Left            =   10560
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      _Version        =   589828
      _ExtentX        =   5953
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Text"
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
      Surface         =   16
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "RequestForm.frx":430E
      textLT          =   "RequestForm.frx":4376
      textCT          =   "RequestForm.frx":438E
      textRT          =   "RequestForm.frx":43A6
      textLM          =   "RequestForm.frx":43BE
      textRM          =   "RequestForm.frx":43D6
      textLB          =   "RequestForm.frx":43EE
      textCB          =   "RequestForm.frx":4406
      textRB          =   "RequestForm.frx":441E
      colorBack       =   "RequestForm.frx":4436
      colorIntern     =   "RequestForm.frx":4460
      colorMO         =   "RequestForm.frx":448A
      colorFocus      =   "RequestForm.frx":44B4
      colorDisabled   =   "RequestForm.frx":44DE
      colorPressed    =   "RequestForm.frx":4508
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   0
      Office2007ColorNorthBottom=   0
      Office2007ColorSouthTop=   0
      Office2007ColorSouthBottom=   4473924
      Office2007ColorInnerBorderMouseOver=   16776960
      Office2007ColorOuterBorderMouseOver=   16776960
      Office2007ColorNorthTopMouseOver=   0
      Office2007ColorNorthBottomMouseOver=   0
      Office2007ColorSouthTopMouseOver=   0
      Office2007ColorSouthBottomMouseOver=   8421504
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
   Begin VB.Label NRFound 
      Alignment       =   2  'Center
      Caption         =   "No Records Found"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   6975
      Left            =   360
      Top             =   360
      Width           =   13695
   End
   Begin VB.Label Get_What 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   12975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   360
      Top             =   7440
      Width           =   13695
   End
End
Attribute VB_Name = "RequestListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TScroll         As Long
Private Total_Display   As Long
Private Bk              As Long
Private Display         As Long
Private ScanCheck       As String
Private ButtonFore      As String
Private ButtonBack      As String
Private FSize           As Single
Private Type ArryProp
        Name As String
        CheckNo As Long
        SvrName As String
        BtnColor As Variant
End Type
Dim SelectionArry() As ArryProp
Private Sub DwnArrow_Click()
    TScroll = TScroll + 28
    Scroller
End Sub
Private Sub Exit_Click()
GetSelected = ""
Call MemoryCheck(RequestListForm)
End Sub
Private Sub LoadTipList()
On Error GoTo ErrorHandler
    Dim MYCheckSeries   As Long
    Dim Db              As ADODB.Connection
    Dim RS              As ADODB.Recordset
    MYCheckSeries = 0
    Total_Display = -1
    ConnectADODB SvrPath.TerminalPath & DebitDb, Db
        If LoggedOnAs(0).EmpAccessLevel < Access_Required("Managers Menu") Then
            Sql = "Select * from [Debit] Where [Tender] = 'EMV' and [Close Out day] = 0 and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by [Check Number] desc"
        Else
            Sql = "Select * from [Debit] Where [Tender] = 'EMV' and [Close Out day] = 0 ORDER BY [Server Name] DESC, [Check Number] DESC"
        End If
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)
    If Not RS.EOF Then
     ReDim SelectionArry(RecordsCount(RS))
     Do While Not RS.EOF
        Total_Display = Total_Display + 1
        If Nz(RS![tip]) = 0 Then
         SelectionArry(Total_Display).Name = "CHK NO " & RS![Check Number] & " Acct. " & "EMV-" & RS.fields("Account Number").Value & " -No Tip Applied"
        Else
         SelectionArry(Total_Display).Name = "Chk No " & RS![Check Number] & " Acct. " & "EMV-" & RS.fields("Account Number").Value & " -Tip Applied- " & Format(Val(RS![tip] & ""), "0.00")
        End If
        MYCheckSeries = CInt(RS![Check Number] / 1000)
        RS.MoveNext
     Loop
    End If
    DbClose Db, RS
    '// sort credit cards above  manager see by server and check number
    '// servers only see their check in desc order
    '// below is any house charges or gift cards
    Sql = "Select * from [HouseCharges] Where [Day of Charge] = 0 and [voided] = 0"
    If MYCheckSeries > 0 Then
     If LoggedOnAs(0).EmpAccessLevel < Access_Required("Managers Menu") Then
       Sql = Sql & " and TRY_CONVERT(int, [Check Number] / 1000) = " & MYCheckSeries
     End If
    End If
    If ConnectADODB(HouseAcct, Db) = False Then
      Exit Sub
    End If
    'If ConnectADORS(sql, Db, Rs) = False Then Exit Sub
    Set RS = OpenRSFast(Db, Sql)
    If Not RS.EOF Then
     ReDim Preserve SelectionArry(UBound(SelectionArry, 1) + RecordsCount(RS))
     Do While Not RS.EOF
      Total_Display = Total_Display + 1
      If Nz(RS![Tip Amount]) = 0 Then
       SelectionArry(Total_Display).Name = "CHK NO " & RS![Check Number] & " Acct. " & Trim$(RS![Scan Number] & "") & " -No Tip Applied"
      Else
       SelectionArry(Total_Display).Name = "Chk No " & RS![Check Number] & " Acct. " & Trim$(RS![Scan Number] & "") & " -Tip Applied"
      End If
      RS.MoveNext
     Loop
    End If
    DbClose Db, RS
    Scroller
    Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Tip List ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub ReturnOpenTabs(SearchHow As String)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim SBy As String
Dim TBGA As String
Dim DataStr As String
Dim ReturnTableNo As Boolean
DataStr = "[Table Number],[Seat Number],[Check Number],[server Name],[Check total]"
ReturnTableNo = False
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
If SearchHow > "" Then
   Sql = SearchHow
   ReturnTableNo = True
Else
    If Check_Option(93) = 0 Then  'tab sharing
       If SvrPath.GroupAudit Then
          TBGA = " Where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EmpAssignedTo & "'"
       Else
          TBGA = " Where [Close Out Day] = 0 and [Transfer From] = '" & SvrPath.TerminalGroupName & "'"
       End If
    Else
          TBGA = " Where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EMPName & "'"
    End If
    If Check_Option(140) = 1 Then  'allow open tabs from QSR Options
      If Check_Option(93) = 0 Then
        Sql = "Select " & DataStr & " from [Open Checks] Where [Close Out Day] = 0 Group By " & DataStr & " order by [Table Number] asc"
      Else
        Sql = "Select " & DataStr & " from [Open Checks] Where [Close Out Day] = 0 and [Transfer From] = '" & SvrPath.TerminalGroupName & "' Group By " & DataStr & " order by [Table Number] asc"
      End If
    Else
       If Val(LoggedOnAs(0).EmpAccessLevel) > Access_Required("View All Checks") Then  'And Check_Option(93) = 0
           Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 Group By " & DataStr & " order by [Table Number] asc"
       Else
           Sql = "Select " & DataStr & " from [Open Checks] " & TBGA & " Group By " & DataStr & " order by [Table Number], [check number] asc"
       End If
    End If
End If
'ConnectADORS sql, Db, Tb
Set Tb = OpenRSFast(Db, Sql)
ReDim SelectionArry(RecordsCount(Tb))
Total_Display = -1
Dim CdsOnFile As Integer
Dim s As Integer
CdsOnFile = -1
Dim ChargesOnFile() As Long
If Check_Option(156) = 1 Then
 CdsOnFile = CreditOnFile(ChargesOnFile())
End If
Do While Not Tb.EOF
   Total_Display = Total_Display + 1
   If ReturnTableNo Or Tb![Table Number] = "Dine_In" Then
     SelectionArry(Total_Display).Name = "Chk No " & Tb![Check Number] & " - " & Tb![Table Number]
   Else
     SelectionArry(Total_Display).Name = Tb![Table Number] & " " & Tb![Seat Number] & "" & "     ( " & Format(Tb![Check Total], CurrencyPreFix) & " )"
   End If
   SelectionArry(Total_Display).CheckNo = Tb![Check Number] & ""
   SelectionArry(Total_Display).SvrName = Tb![Server Name] & ""
   SelectionArry(Total_Display).BtnColor = ""
   If CdsOnFile > -1 Then
     For s = 0 To CdsOnFile
      If SelectionArry(Total_Display).CheckNo = ChargesOnFile(s) Then
        SelectionArry(Total_Display).BtnColor = &H404080
       Exit For
      End If
     Next s
   End If
 Tb.MoveNext
Loop
DbClose Db, Tb
If Total_Display = -1 Then
  ClearRecords
Else
 NRFound.Visible = False
End If
Exit Sub
ErrorHandler:
Err.Clear
End Sub
Public Function GetScanCheck(ScanNo As String) As Integer
On Error GoTo ErrorHandler
Dim FC As Integer
For FC = 0 To Total_Display
  If SelectionArry(FC).CheckNo = ScanNo & "" Then
   GetScanCheck = FC
   Exit Function
  End If
Next FC
See_Msg "Check Not Found", 0, ""
GetScanCheck = -1
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Request List Form Get Scan Check ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub ReturnTabs()
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim SBy As String
Dim TBGA As String
Dim DataStr As String
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
Sql = ""
DataStr = "[Table Number],[Seat Number],[Check Number],[server Name],[Check total]"
If Check_Option(39) = 1 Then
 If SvrPath.TerminalGroupName <> MyRegisterName And LoggedOnAs(0).EmpAccessLevel < Access_Required("Managers Menu") And Check_Option(93) = 0 Then
    If SvrPath.GroupAudit Then
       Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EmpAssignedTo & "' Group By " & DataStr & " order by [Table Number], [check number] asc"
    Else
       Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 and [Transfer From] = '" & SvrPath.TerminalGroupName & "' Group By " & DataStr & " order by [Table Number], [check number] asc"
    End If
 Else
        Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EMPName & "' Group By " & DataStr & " order by [Table Number], [check number] asc"
 End If
 If LoggedOnAs(0).EmpAccessLevel >= Access_Required("Managers Menu") And SvrPath.GroupAudit = False Then
    Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 Group By " & DataStr & " order by [Table Number] asc"
 End If
Else
    If LoggedOnAs(0).EmpAccessLevel >= Access_Required("View All Checks") Then
      Sql = "Select " & DataStr & " from [Open Checks] Where [Close Out Day] = 0 Group By " & DataStr & " order by [Table Number] asc"
    Else
     If SvrPath.GroupAudit Then
      Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EmpAssignedTo & "' Group By " & DataStr & "Order  by [Table Number], [check number] asc"
     Else
      Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EMPName & "' Group By " & DataStr & "Order  by [Table Number], [check number] asc"
     End If
    End If
 End If
If Sql = "" Then 'Makes sure sql call is made
      Sql = "Select " & DataStr & " from [Open Checks] where [Close Out Day] = 0 and [server Name] = '" & LoggedOnAs(0).EMPName & "' Group By " & DataStr & " order by [Table Number], [check number] asc"
End If
'*******************old segment
'ConnectADORS sql, Db, Tb
Set Tb = OpenRSFast(Db, Sql)
ReDim SelectionArry(RecordsCount(Tb))
Do While Not Tb.EOF
   Total_Display = Total_Display + 1
   SelectionArry(Total_Display).Name = Tb![Table Number] & " " & Tb![Seat Number] & "" & "     ( " & Format(Tb![Check Total], CurrencyPreFix) & " )"
   SelectionArry(Total_Display).CheckNo = Tb![Check Number] & ""
   SelectionArry(Total_Display).SvrName = Tb![Server Name] & ""
 Tb.MoveNext
Loop
DbClose Db, Tb
Exit Sub
ErrorHandler:
Err.Clear
End Sub
Private Sub LoadAllChecks()
On Error GoTo ErrorHandler
Dim Db   As ADODB.Connection
Dim Tb   As ADODB.Recordset
'Dim SBy  As String
'Dim TBG  As String 'tabs by group
'Dim TBGA As String
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select * from [Open Checks] where [Close Out Day] = 0 order by [Table Number], [check number] asc"
'ConnectADORS sql, Db, Tb
Set Tb = OpenRSFast(Db, Sql)
ReDim SelectionArry(RecordsCount(Tb))
Total_Display = -1
Do While Not Tb.EOF
   Total_Display = Total_Display + 1
   SelectionArry(Total_Display).Name = Tb![Table Number] & " " & Tb![Seat Number] & "" & "     ( " & Format(Tb![Check Total], CurrencyPreFix) & " )"
   SelectionArry(Total_Display).CheckNo = Tb![Check Number] & ""
   SelectionArry(Total_Display).SvrName = Tb![Server Name] & ""
 Tb.MoveNext
Loop
DbClose Db, Tb
Exit Sub
ErrorHandler:
Err.Clear
End Sub
Private Sub Form_Activate()
If Total_Display = -1 Then
  ClearRecords
Else
  NRFound.Visible = False
End If
'If Total_Display = -1 Then Exit_Click
'Do Not Code Like This will cause fault Errors
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
ResizeScreen1024 Me
CenterForm Me
ScanCheck = ""
ButtonFore = GetSetting("Micro$ale", "Register Setup", "Tab Fore Color")
ButtonBack = GetSetting("Micro$ale", "Register Setup", "Tab Back Color")
FSize = Val(GetSetting("Micro$ale", "Register Setup", "Tab Font Size"))
If FSize = 0 Then FSize = 11
Dim Db                  As ADODB.Connection
Dim RS                  As ADODB.Recordset
Dim ASdb                As ADODB.Connection
Dim ASrs                As ADODB.Recordset
Dim MYCheckSeries       As Long
Dim CurrentCasherName   As String
Dim TipOwed             As Currency

ReDim SelectionArry(0): Total_Display = -1: TScroll = 0

Get_What.Caption = Find_What
If Find_What = "Get Closed Check" Then SSCommand12.Visible = False
 If Load_What = "Change Sections" Then
    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
     Sql = "Select Distinct [Layout Name] from [Layout]"
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)

      If Not RS.EOF Then
        ReDim SelectionArry(RecordsCount(RS))
        Do While Not RS.EOF
          Total_Display = Total_Display + 1
          SelectionArry(Total_Display).Name = RS![Layout Name] & ""
          RS.MoveNext
        Loop
     End If
    DbClose Db, RS
    Scroller
    Exit Sub
 End If
If Load_What = "ShowAllChecks" Then
  LoadAllChecks
  Scroller
  Exit Sub
End If
If Load_What = "Return Tabs" Then
    ReturnOpenTabs ""
    Scroller
    SSCommand12.Visible = True
    Exit Sub
End If
If Load_What = "Return Server Tabs" Then
    ReturnTabs
    Scroller
    Exit Sub
End If
If Load_What = "List All Checks" Then
    ReturnOpenTabs ""
    Scroller
    Exit Sub
End If

If Load_What = "StartServerAudit" Then '// starting audit from managers screen.
    Sql = "Select * From [General Data] Where ISNULL([Terminated], 0) = 0  order by [Last Name] asc"
    ConnectADODB "MSaleData", Db
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)

    If Not RS.EOF Then
     ReDim SelectionArry(RecordsCount(RS))
     Do While Not RS.EOF
        If Check_System_Access(RS(7)) > 0 Then
           Total_Display = Total_Display + 1
           SelectionArry(Total_Display).Name = RS(0) & " " & RS(2)
        End If
      RS.MoveNext
     Loop
    End If
    DbClose Db, RS
    Scroller
  Exit Sub
End If
If Load_What = "Paid In" Or Load_What = "Misc Item" Or Load_What = "Customer Request" Or Load_What = "Void Reasons" Or Load_What = "Coupon Reasons" Then
   If Load_What = "Void Reasons" Then
      Sql = "Select * from [Void Reasons] Order by [Void Name] asc"
   ElseIf Load_What = "Customer Request" Then
      Sql = "Select * from [Special Requests] Where [Remote 1] = '' Order by [Request Name] asc"
   ElseIf Load_What = "Misc Item" Then
      Sql = "Select * from [Special Requests] Where [Remote 1] = 'Misc Item' Order by [Request Name] asc"
   ElseIf Load_What = "Paid In" Then
      Sql = "Select * from [Special Requests] Where [Remote 1] = 'Paid In' Order by [Request Name] asc"
   Else
      Sql = "Select * from [Coupon Reason] Order by [Name] asc "
   End If
   ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Db
   'ConnectADORS sql, Db, Rs
   Set RS = OpenRSFast(Db, Sql)
    If RS.EOF Then
      Total_Display = 0
      SelectionArry(0).Name = "Keyboard"
      DbClose Db, RS
      Scroller
      Exit Sub
    End If
    ReDim SelectionArry(RecordsCount(RS))
    Do While Not RS.EOF
      If UCase(RS(0) & "") = "KEYBOARD" Then
        Total_Display = Total_Display + 1
        SelectionArry(0).Name = "Keyboard"
        Exit Do
      End If
      RS.MoveNext
    Loop
    If Not RS.BOF Then RS.MoveFirst
    Do While Not RS.EOF
      If UCase(RS(0) & "") <> "KEYBOARD" Then
       Total_Display = Total_Display + 1
       SelectionArry(Total_Display).Name = RS(0) & ""
      End If
      RS.MoveNext
    Loop
    DbClose Db, RS
    Scroller
    Exit Sub
End If
If Load_What = "ServerWork" Then
        ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
         Sql = "Select Distinct [Servers Name] from [Check Numbers]"
        'ConnectADORS sql, Db, Rs
        Set RS = OpenRSFast(Db, Sql)
        If Not RS.EOF Then
         ReDim SelectionArry(RecordsCount(RS))
         Do While Not RS.EOF
            TipOwed = GetUnPaidTips(RS![Servers Name] & "", ReturnMyCashier)
             Total_Display = Total_Display + 1
            If TipOwed > 0 Then
             SelectionArry(Total_Display).Name = RS![Servers Name] & ""
             SelectionArry(Total_Display).SvrName = "Tip Amount Due " & Format(TipOwed, "0.00")
            Else
             SelectionArry(Total_Display).Name = RS![Servers Name] & ""
             SelectionArry(Total_Display).SvrName = "Tips Paid"
            End If
            RS.MoveNext
         Loop
        End If
        DbClose Db, RS
        Scroller
        Exit Sub
End If
If Load_What = "AllTiped" Then
        ConnectADODB "MSaleData", Db
         Sql = "Select [First Name], [Last Name] from [Job Codes] Where [Tip] = 1 "
         Sql = Sql & " Group by [First Name], [Last Name]"
        'ConnectADORS sql, Db, Rs
        Set RS = OpenRSFast(Db, Sql)
        If Not RS.EOF Then
         ReDim SelectionArry(RecordsCount(RS))
         Do While Not RS.EOF
            Total_Display = Total_Display + 1
            SelectionArry(Total_Display).Name = RS![First Name] & " " & RS![Last Name]
            RS.MoveNext
         Loop
        End If
        DbClose Db, RS
        Scroller
        Exit Sub
End If
If Load_What = "Closed Servers" Then
        ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
        Sql = "Select [Servers Name] From [Check Numbers] Where [Audit Status] = 1 Group by [Servers Name] order by [Servers Name] asc"
        Set RS = OpenRSFast(Db, Sql)
        If RS.EOF Then
           See_Msg "No Closed Audits to ReOpen", 5, ""
        Else
         ReDim SelectionArry(RecordsCount(RS))
         Do While Not RS.EOF
          Total_Display = Total_Display + 1
          SelectionArry(Total_Display).Name = RS![Servers Name] & ""
          RS.MoveNext
         Loop
        End If
        DbClose Db, RS
        Scroller
        Exit Sub
End If
If Load_What = "AllServers" Then
       ConnectADODB "MSaleData", Db
       Sql = "Select * From [General Data] Where not isnull(Terminated) and [Terminated] = 0 order by [Last Name] asc"
       'ConnectADORS sql, Db, Rs
       Set RS = OpenRSFast(Db, Sql)
       If Not RS.EOF Then
        ReDim SelectionArry(RecordsCount(RS))
        Do While Not RS.EOF
         Total_Display = Total_Display + 1
         SelectionArry(Total_Display).Name = RS.fields("First Name").Value & " " & RS.fields("Last Name").Value
         RS.MoveNext
        Loop
       End If
       DbClose Db, RS
       Scroller
       Exit Sub
End If
If Load_What = "Work Area" Then
    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
     Sql = "Select Distinct [Layout Name] from [Layout]"
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)
     If Not RS.EOF Then
        ReDim SelectionArry(RecordsCount(RS) + 1)
        Do While Not RS.EOF
          Total_Display = Total_Display + 1
          SelectionArry(Total_Display).Name = RS![Layout Name] & ""
          RS.MoveNext
        Loop
     End If
     If Total_Display > 1 Then
        Total_Display = Total_Display + 1
        SelectionArry(Total_Display).Name = "All Areas"
     End If
     DbClose Db, RS
     Scroller
     Exit Sub
End If
If Load_What = "Drivers" And System_config = ExpressService Then
    If CPath.TerminalFailed Then
      See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
      Exit Sub
    End If
    ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
    'sql = "Select Distinct [employee Name] from [Employee records] where [time in] > '" & "" & "' and [time out] = '" & "" & "'"
    Sql = "SELECT DISTINCT [Employee Name] From [Employee Records] Where [Time In] Is Not Null  AND [Time Out] IS NULL;"

    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)
        If RS.EOF Then
          See_Msg "No Records Found", 5, ""
        Else
         ReDim SelectionArry(RecordsCount(RS))
         Do While Not RS.EOF
          Total_Display = Total_Display + 1
          SelectionArry(Total_Display).Name = RS![employee name] & ""
          RS.MoveNext
         Loop
        End If
        DbClose Db, RS
        Scroller
        Exit Sub
ElseIf Load_What = "Drivers" And System_config = FullService Then
        ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
        Sql = "Select * from [Check Numbers] where [audit status] = 0"
        If GetConfigInfo("Phone Orders", "Drivers Assignment Only Is", "Off", "FUNCTIONS.INI") = "ENABLED" Then
          Sql = Sql & " and [Work Area] = 'Driver'"
        End If
        'ConnectADORS sql, Db, Rs
        Set RS = OpenRSFast(Db, Sql)
        If RS.EOF Then
           See_Msg "No Records Found", 5, ""
        Else
         ReDim SelectionArry(RecordsCount(RS))
         Do While Not RS.EOF
           Total_Display = Total_Display + 1
           SelectionArry(Total_Display).Name = RS![Servers Name]
          RS.MoveNext
         Loop
        End If
        DbClose Db, RS
        Scroller
        Exit Sub
End If
If Load_What = "Set as Drivers" And System_config = FullService Then
       ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
       Sql = "Select * from [Check Numbers] where [audit status] = 0"
       'ConnectADORS sql, Db, Rs
       Set RS = OpenRSFast(Db, Sql)
       If RS.EOF Then
          See_Msg "No Records Found", 5, ""
       Else
        ReDim SelectionArry(RecordsCount(RS))
        Do While Not RS.EOF
          Total_Display = Total_Display + 1
          SelectionArry(Total_Display).Name = RS![Servers Name]
         RS.MoveNext
        Loop
       End If
      DbClose Db, RS
      Scroller
      Exit Sub
End If
If Load_What = "Charge" Or Load_What = "Charge All" Then
    Type_Message = ShowKeyBoard("Enter Name or Inital")
    If Type_Message = "" Then Exit Sub
    If ConnectADODB(HouseAcct, Db) = False Then
      Exit Sub
    End If
    Sql = "Select [Scan Name],[Phone],[EMail],[Acct Number],[Address] from [HouseAccountData] "
    Select Case Load_What
     Case Is = "Charge"
          Sql = Sql & " Where [Suspend] = 0 and ([Scan Name] like '" & Type_Message & "%' or [Acct Number] like '" & Type_Message & "%' or [Spc Interest 1] Like '" & Type_Message & "%')"
         
          If TenderLink > "" Then
           Sql = Sql & " and [Account type] = '" & TenderLink & "'"
          End If
          Sql = Sql & " Order By [Scan Name] asc "
         'If ConnectADORS(sql, Db, Rs) = False Then Exit Sub
          Set RS = OpenRSFast(Db, Sql)
          If RS.EOF Then
            SetNothing RS
            Sql = "Select [Scan Name],[Phone],[Address] from [HouseAccountData] "
            Sql = Sql & " Where [suspend] = 0 and [Scan Name] Like '%" & Type_Message & "%';"
            'If ConnectADORS(sql, Db, Rs) = False Then Exit Sub
            Set RS = OpenRSFast(Db, Sql)
          End If
          If RS.EOF Then
           DbClose Db, RS
           Exit Sub
          End If
     Case Else
        Sql = Sql & " Where [Suspend] = 0 and [Scan Name] like '%" & Type_Message & "%'"
        Sql = Sql & " Or [suspend] = 0 and [Acct Number] like '%" & Type_Message & "%'"
        If TenderLink > "" Then
         Sql = Sql & " and [Account type] = '" & TenderLink & "'"
        End If
        Sql = Sql & " Order By [Scan Name] asc "
        'If ConnectADORS(sql, Db, Rs) = False Then Exit Sub
        Set RS = OpenRSFast(Db, Sql)
        If RS.EOF Then
         DbClose Db, RS
         Exit Sub
        End If
     End Select
    ReDim SelectionArry(RecordsCount(RS))
    Do While Not RS.EOF
      Total_Display = Total_Display + 1
      SelectionArry(Total_Display).Name = RS![scan Name] & ""
      If GetConfigInfo("House Account", "Search By Address", "No", "Functions.INI") = "ENABLED" Then
       SelectionArry(Total_Display).SvrName = RS![Address] & ""
      Else
       SelectionArry(Total_Display).SvrName = RS![scan Name] & ""
      End If
      'If Input_Option(50) = "Account #" Then
      ' SelectionArry(Total_Display).SvrName = RS![Acct Number] & ""
      'ElseIf Input_Option(50) = "Address" Then
      ' SelectionArry(Total_Display).SvrName = RS![Address] & ""
      'ElseIf Input_Option(50) = "EMail" Then
      ' SelectionArry(Total_Display).SvrName = RS![Email] & ""
      'Else
      ' SelectionArry(Total_Display).SvrName = RS![Phone] & ""
      'End If
    RS.MoveNext
    Loop
    DbClose Db, RS
    Scroller
    Exit Sub
End If
If Load_What = "GetTip" Then
    LoadTipList
    Exit Sub
End If
If Load_What = "AllTips" Then
    ConnectADODB SvrPath.TerminalPath & DebitDb, Db
      Sql = "Select * from [Debit] Where [Tender] = 'EMV' and [Close Out day] = 0 order by [time] desc"
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)
    If Not RS.EOF Then
     ReDim SelectionArry(RecordsCount(RS))
     Do While Not RS.EOF
        Total_Display = Total_Display + 1
        If Val(RS![tip] & "") = 0 Then
         SelectionArry(Total_Display).Name = "CHK NO " & RS![Check Number] & " Acct. " & "EMV-" & RS.fields("Account Number").Value & " -No Tip Applied"
        Else
         SelectionArry(Total_Display).Name = "Chk No " & RS![Check Number] & " Acct. " & "EMV-" & RS.fields("Account Number").Value & " -Tip Applied- " & Format(Val(RS![tip] & ""), "0.00")
        End If
        MYCheckSeries = CInt(RS![Check Number] / 1000)
        RS.MoveNext
     Loop
    End If
    DbClose Db, RS
    Sql = "Select * from [HouseCharges] Where [Day of Charge] = 0 and [voided] = 0"
    If ConnectADODB(HouseAcct, Db) = False Then
      Exit Sub
    End If
    'If ConnectADORS(sql, Db, Rs) = False Then Exit Sub
    Set RS = OpenRSFast(Db, Sql)
    If Not RS.EOF Then
    ReDim Preserve SelectionArry(UBound(SelectionArry, 1) + RecordsCount(RS))
    Do While Not RS.EOF
      Total_Display = Total_Display + 1
      If Val(RS![Tip Amount] & "") = 0 Then
      SelectionArry(Total_Display).Name = "CHK NO " & RS![Check Number] & " Acct. " & Trim(RS![Scan Number]) & " -No Tip Applied"
      Else
      SelectionArry(Total_Display).Name = "Chk No " & RS![Check Number] & " Acct. " & Trim(RS![Scan Number]) & " -Tip Applied"
      End If
      RS.MoveNext
    Loop
   End If
   DbClose Db, RS
   Scroller
   Exit Sub
End If
If Load_What = "Find Check" Then
   SSCommand12.Visible = True
   Dim ChkNO As Long
   ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    Sql = "Select * from [check Numbers] where [Servers Name] = '" & SelectedPerson & "'"
    Sql = Sql & " And [Audit Status] = 0"
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)
     If RS.EOF Then
        DbClose Db, RS
        See_Msg "No Audit Open for Server", 5, "Audit Not Found"
        Exit Sub
     Else
        ChkNO = RS(1)
     End If
     DbClose Db, RS
     'If Check_Option(65) = 1 Then ChkNO = 1
     ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
       Sql = "Select * from [Closed Checks] "
       Sql = Sql & " Where [CHeck Number] > " & ChkNO - 1
       Sql = Sql & " and [Close out day] = 0 "
       Sql = Sql & " and [server Name] = '" & SelectedPerson & "' order by [check number] desc"
     'ConnectADORS sql, Db, Rs
     Set RS = OpenRSFast(Db, Sql)
       If Not RS.EOF Then
        ReDim SelectionArry(RecordsCount(RS))
        Do While Not RS.EOF
         Total_Display = Total_Display + 1
         SelectionArry(Total_Display).Name = "Chk No " & RS![Check Number] & " - " & RS![Table Number]
         RS.MoveNext
        Loop
       End If
     DbClose Db, RS
    Scroller
    Exit Sub
End If
If Load_What = "Reopen Phone Order" Then
     SSCommand12.Visible = True '' search by check number buttono
     ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
       Sql = "Select * from [Closed Checks] "
       Sql = Sql & " Where [Order Destination] = 'Delivery' and [Close out day] = 0 or [Order Destination] = 'Pick Up'"
       Sql = Sql & " and [Close out day] = 0 "
       Sql = Sql & " order by [Check Number] desc "
     'ConnectADORS sql, Db, Rs
     Set RS = OpenRSFast(Db, Sql)
       If Not RS.EOF Then
        ReDim SelectionArry(RecordsCount(RS))
        Do While Not RS.EOF
         Total_Display = Total_Display + 1
         SelectionArry(Total_Display).Name = "Chk No " & RS![Check Number] & " - " & RS![Table Number]
         SelectionArry(Total_Display).CheckNo = RS![Check Number]
         RS.MoveNext
        Loop
       End If
     DbClose Db, RS
    Scroller
    Exit Sub
End If

If Load_What = "Frequent" Then
    Type_Message = ShowKeyBoard("Enter Name or Inital")
    If Trim(Type_Message) > "" Then
      Sql = "Select [Scan Name] from [HouseAccountData] where [Scan Name] like '" & Type_Message & "%'"   '& " order by [Scan Name] asc"
      If GetConfigInfo("Frequent Diner", "Account Search By Frequent Dinner Only", "No", "Functions.INI") = "ENABLED" Then
       Sql = Sql & " and [Account Type] = 'Frequent Diner'"
      End If
    Else
      Sql = "Select * from [HouseAccountData]"
      If GetConfigInfo("Frequent Diner", "Account Search By Frequent Dinner Only", "No", "Functions.INI") = "ENABLED" Then
       Sql = Sql & " Where [Account Type] = 'Frequent Diner'"
      End If
    End If
    ConnectADODB HouseAcct, Db
    'ConnectADORS sql, Db, Rs
    Set RS = OpenRSFast(Db, Sql)
    If Not RS.EOF Then
     ReDim SelectionArry(RecordsCount(RS))
     Do While Not RS.EOF
      Total_Display = Total_Display + 1
      SelectionArry(Total_Display).Name = RS![scan Name] & ""
     RS.MoveNext
     Loop
    End If
    DbClose Db, RS
    Scroller
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Request List Form Load " & Load_What)
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Erase SelectionArry
End Sub
Private Sub SelectionButton_Click(Index As Integer)
On Error GoTo ErrorHandler
   If Load_What = "Reopen Phone Order" Or Load_What = "Return Tabs" Or Load_What = "ShowAllChecks" Or Load_What = "Return Server Tabs" Or Load_What = "Add To Tabs" Then
       GetSelected = CStr(SelectionArry(Index + TScroll).CheckNo)
       SelectedPerson = SelectionArry(Index + TScroll).SvrName
   ElseIf Load_What = "Charge" Then    'And Input_Option(50) = "Account #"
       GetSelected = SelectionArry(Index + TScroll).Name
   Else
    If UCase(SelectionButton(Index).Caption) = "KEYBOARD" Then
       GetSelected = ShowKeyBoard("Enter Message")
    Else
       GetSelected = SelectionButton(Index).Caption
    End If
    If Load_What = "GetTip" Then
        Dim GInfo(1)        As String
        If InStr(GetSelected, "EMV") > 0 Then
          GInfo(0) = Mid(GetSelected, 8, InStr(GetSelected, "A") - 9)
          ListCards.LoadEMVChecks Val(GInfo(0))
          ListCards.Show 1
          LoadTipList
          Exit Sub
        End If
    End If
   End If
Call MemoryCheck(Me)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Request List Form Selection Button ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub SSCommand12_Click()
On Error Resume Next
Dim SearchData          As String
Dim ReturnedValue       As Long
If Load_What = "Reopen Phone Order" Then
  GetSelected = ShowNumberPad("Find Check", False, 0, 0)
  If Val(GetSelected) = 0 Then Exit Sub
  Unload Me
  Exit Sub
End If
If Load_What = "Find Check" Then
   'ReturnedValue = ShowNumberPad("Find Check", False, 0, 0)
   SearchData = ShowKeyBoard("Enter Name or Check No")
   If IsNumeric(SearchData) Then
      Sql = "Select * from [Closed Checks] where [Close out Day] = 0 and [Check Number] = " & Val(SearchData)
      ReturnOpenTabs Sql
   ElseIf Not IsEmpty(SearchData) Then
     Sql = "Select * from [Closed Checks] where [Close Out Day] = 0 and [Table Number] like '" & SearchData & "%'"
     ReturnOpenTabs Sql
   Else
     Exit Sub
   End If
Else
 SearchData = ShowKeyBoard("Enter Name")
 If SearchData = "" Then
  ReturnOpenTabs ""
 Else
  Sql = "Select * from [Open Checks] where [Close Out Day] = 0 and [Table Number] like '" & SearchData & "%'"
  ReturnOpenTabs Sql
 End If
End If
Scroller
End Sub
Private Sub UpArrow_Click()
TScroll = TScroll - 28
Scroller
End Sub
Public Sub ClearRecords()
NRFound.Visible = True
For Display = 0 To 27
    SelectionButton(Display).Caption = ""
    SelectionButton(Display).Visible = False
    SelectionButton(Display).ClearButtonCache
Next Display
DwnArrow.Visible = False
UpArrow.Visible = False
End Sub
Public Sub Scroller()
On Error GoTo ErrorHandler
If Total_Display = -1 Then Exit Sub
For Display = 0 To 27
    SelectionButton(Display).Caption = ""
    SelectionButton(Display).Visible = False
    SelectionButton(Display).ClearButtonCache
    If TScroll + Display < 0 Then TScroll = 0
    If TScroll + Display <= Total_Display Then
        SelectionButton(Display).BackColor = &HC46200
        If UCase(SelectionArry(TScroll + Display).Name) = "KEYBOARD" Then
          SelectionButton(Display).BackColor = &HFF&
        Else
         If Left(SelectionArry(TScroll + Display).Name, 6) = "CHK NO" Or Left(SelectionArry(TScroll + Display).SvrName, 14) = "Tip Amount Due" Then
          'SelectionButton(Display).BackColorInterior = &HFF&
          SelectionButton(Display).BackColor = &HFF&
         Else
           If ButtonFore > "" Then
             SelectionButton(Display).BackColor = ButtonFore
           End If
           If ButtonBack > "" Then
             SelectionButton(Display).BackColor = ButtonBack
           Else
             SelectionButton(Display).BackColor = &HC46200 '&H404040         ' &HFF8080
           End If
         End If
        End If
        SelectionButton(Display).FontTextCaption.Bold = True
        SelectionButton(Display).FontTextCaption.Italic = False
        SelectionButton(Display).Caption = SelectionArry(TScroll + Display).Name
        SelectionButton(Display).TextDescrCB.Text = SelectionArry(TScroll + Display).SvrName
        SelectionButton(Display).TextDescrCB.ColorNormal = vbWhite
        SelectionButton(Display).Visible = True
        SelectionButton(Display).FontTextCaption.Size = FSize
        If SelectionArry(TScroll + Display).BtnColor = &H404080 Then
         SelectionButton(Display).BackColor = &H404080
        End If
    End If
    
Next Display
If TScroll <= 0 Then
    UpArrow.Visible = False
    Else
    UpArrow.Visible = True
End If
If Total_Display > (TScroll + Display - 1) Then
    DwnArrow.Visible = True
Else
    DwnArrow.Visible = False
End If
UpArrow.Refresh
DwnArrow.Refresh
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Request List Form Scroller ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
'Me.Visible = True
End Sub

                                                                                      