VERSION 5.00
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.3#0"; "3dabm9u.ocx"
Begin VB.Form Cashier_Audit 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cashier Audit"
   ClientHeight    =   7230
   ClientLeft      =   4710
   ClientTop       =   2340
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7230
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BTNENHLib4.BtnEnh SSCommand3 
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "."
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":0000
      textLT          =   "Cashier_Audit.frx":0062
      textCT          =   "Cashier_Audit.frx":007A
      textRT          =   "Cashier_Audit.frx":0092
      textLM          =   "Cashier_Audit.frx":00AA
      textRM          =   "Cashier_Audit.frx":00C2
      textLB          =   "Cashier_Audit.frx":00DA
      textCB          =   "Cashier_Audit.frx":00F2
      textRB          =   "Cashier_Audit.frx":010A
      colorBack       =   "Cashier_Audit.frx":0122
      colorIntern     =   "Cashier_Audit.frx":014C
      colorMO         =   "Cashier_Audit.frx":0176
      colorFocus      =   "Cashier_Audit.frx":01A0
      colorDisabled   =   "Cashier_Audit.frx":01CA
      colorPressed    =   "Cashier_Audit.frx":01F4
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.ComboBox Cashier 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Cashier"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1935
      End
   End
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "0"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":021E
      textLT          =   "Cashier_Audit.frx":0280
      textCT          =   "Cashier_Audit.frx":0298
      textRT          =   "Cashier_Audit.frx":02B0
      textLM          =   "Cashier_Audit.frx":02C8
      textRM          =   "Cashier_Audit.frx":02E0
      textLB          =   "Cashier_Audit.frx":02F8
      textCB          =   "Cashier_Audit.frx":0310
      textRB          =   "Cashier_Audit.frx":0328
      colorBack       =   "Cashier_Audit.frx":0340
      colorIntern     =   "Cashier_Audit.frx":036A
      colorMO         =   "Cashier_Audit.frx":0394
      colorFocus      =   "Cashier_Audit.frx":03BE
      colorDisabled   =   "Cashier_Audit.frx":03E8
      colorPressed    =   "Cashier_Audit.frx":0412
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":043C
      textLT          =   "Cashier_Audit.frx":049E
      textCT          =   "Cashier_Audit.frx":04B6
      textRT          =   "Cashier_Audit.frx":04CE
      textLM          =   "Cashier_Audit.frx":04E6
      textRM          =   "Cashier_Audit.frx":04FE
      textLB          =   "Cashier_Audit.frx":0516
      textCB          =   "Cashier_Audit.frx":052E
      textRB          =   "Cashier_Audit.frx":0546
      colorBack       =   "Cashier_Audit.frx":055E
      colorIntern     =   "Cashier_Audit.frx":0588
      colorMO         =   "Cashier_Audit.frx":05B2
      colorFocus      =   "Cashier_Audit.frx":05DC
      colorDisabled   =   "Cashier_Audit.frx":0606
      colorPressed    =   "Cashier_Audit.frx":0630
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "2"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":065A
      textLT          =   "Cashier_Audit.frx":06BC
      textCT          =   "Cashier_Audit.frx":06D4
      textRT          =   "Cashier_Audit.frx":06EC
      textLM          =   "Cashier_Audit.frx":0704
      textRM          =   "Cashier_Audit.frx":071C
      textLB          =   "Cashier_Audit.frx":0734
      textCB          =   "Cashier_Audit.frx":074C
      textRB          =   "Cashier_Audit.frx":0764
      colorBack       =   "Cashier_Audit.frx":077C
      colorIntern     =   "Cashier_Audit.frx":07A6
      colorMO         =   "Cashier_Audit.frx":07D0
      colorFocus      =   "Cashier_Audit.frx":07FA
      colorDisabled   =   "Cashier_Audit.frx":0824
      colorPressed    =   "Cashier_Audit.frx":084E
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   3
      Left            =   2880
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "3"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":0878
      textLT          =   "Cashier_Audit.frx":08DA
      textCT          =   "Cashier_Audit.frx":08F2
      textRT          =   "Cashier_Audit.frx":090A
      textLM          =   "Cashier_Audit.frx":0922
      textRM          =   "Cashier_Audit.frx":093A
      textLB          =   "Cashier_Audit.frx":0952
      textCB          =   "Cashier_Audit.frx":096A
      textRB          =   "Cashier_Audit.frx":0982
      colorBack       =   "Cashier_Audit.frx":099A
      colorIntern     =   "Cashier_Audit.frx":09C4
      colorMO         =   "Cashier_Audit.frx":09EE
      colorFocus      =   "Cashier_Audit.frx":0A18
      colorDisabled   =   "Cashier_Audit.frx":0A42
      colorPressed    =   "Cashier_Audit.frx":0A6C
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "4"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":0A96
      textLT          =   "Cashier_Audit.frx":0AF8
      textCT          =   "Cashier_Audit.frx":0B10
      textRT          =   "Cashier_Audit.frx":0B28
      textLM          =   "Cashier_Audit.frx":0B40
      textRM          =   "Cashier_Audit.frx":0B58
      textLB          =   "Cashier_Audit.frx":0B70
      textCB          =   "Cashier_Audit.frx":0B88
      textRB          =   "Cashier_Audit.frx":0BA0
      colorBack       =   "Cashier_Audit.frx":0BB8
      colorIntern     =   "Cashier_Audit.frx":0BE2
      colorMO         =   "Cashier_Audit.frx":0C0C
      colorFocus      =   "Cashier_Audit.frx":0C36
      colorDisabled   =   "Cashier_Audit.frx":0C60
      colorPressed    =   "Cashier_Audit.frx":0C8A
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   5
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "5"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":0CB4
      textLT          =   "Cashier_Audit.frx":0D16
      textCT          =   "Cashier_Audit.frx":0D2E
      textRT          =   "Cashier_Audit.frx":0D46
      textLM          =   "Cashier_Audit.frx":0D5E
      textRM          =   "Cashier_Audit.frx":0D76
      textLB          =   "Cashier_Audit.frx":0D8E
      textCB          =   "Cashier_Audit.frx":0DA6
      textRB          =   "Cashier_Audit.frx":0DBE
      colorBack       =   "Cashier_Audit.frx":0DD6
      colorIntern     =   "Cashier_Audit.frx":0E00
      colorMO         =   "Cashier_Audit.frx":0E2A
      colorFocus      =   "Cashier_Audit.frx":0E54
      colorDisabled   =   "Cashier_Audit.frx":0E7E
      colorPressed    =   "Cashier_Audit.frx":0EA8
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   6
      Left            =   2880
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "6"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":0ED2
      textLT          =   "Cashier_Audit.frx":0F34
      textCT          =   "Cashier_Audit.frx":0F4C
      textRT          =   "Cashier_Audit.frx":0F64
      textLM          =   "Cashier_Audit.frx":0F7C
      textRM          =   "Cashier_Audit.frx":0F94
      textLB          =   "Cashier_Audit.frx":0FAC
      textCB          =   "Cashier_Audit.frx":0FC4
      textRB          =   "Cashier_Audit.frx":0FDC
      colorBack       =   "Cashier_Audit.frx":0FF4
      colorIntern     =   "Cashier_Audit.frx":101E
      colorMO         =   "Cashier_Audit.frx":1048
      colorFocus      =   "Cashier_Audit.frx":1072
      colorDisabled   =   "Cashier_Audit.frx":109C
      colorPressed    =   "Cashier_Audit.frx":10C6
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   7
      Left            =   480
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "7"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":10F0
      textLT          =   "Cashier_Audit.frx":1152
      textCT          =   "Cashier_Audit.frx":116A
      textRT          =   "Cashier_Audit.frx":1182
      textLM          =   "Cashier_Audit.frx":119A
      textRM          =   "Cashier_Audit.frx":11B2
      textLB          =   "Cashier_Audit.frx":11CA
      textCB          =   "Cashier_Audit.frx":11E2
      textRB          =   "Cashier_Audit.frx":11FA
      colorBack       =   "Cashier_Audit.frx":1212
      colorIntern     =   "Cashier_Audit.frx":123C
      colorMO         =   "Cashier_Audit.frx":1266
      colorFocus      =   "Cashier_Audit.frx":1290
      colorDisabled   =   "Cashier_Audit.frx":12BA
      colorPressed    =   "Cashier_Audit.frx":12E4
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   8
      Left            =   1680
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "8"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":130E
      textLT          =   "Cashier_Audit.frx":1370
      textCT          =   "Cashier_Audit.frx":1388
      textRT          =   "Cashier_Audit.frx":13A0
      textLM          =   "Cashier_Audit.frx":13B8
      textRM          =   "Cashier_Audit.frx":13D0
      textLB          =   "Cashier_Audit.frx":13E8
      textCB          =   "Cashier_Audit.frx":1400
      textRB          =   "Cashier_Audit.frx":1418
      colorBack       =   "Cashier_Audit.frx":1430
      colorIntern     =   "Cashier_Audit.frx":145A
      colorMO         =   "Cashier_Audit.frx":1484
      colorFocus      =   "Cashier_Audit.frx":14AE
      colorDisabled   =   "Cashier_Audit.frx":14D8
      colorPressed    =   "Cashier_Audit.frx":1502
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
   Begin BTNENHLib4.BtnEnh Number 
      Height          =   735
      Index           =   9
      Left            =   2880
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "9"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      textCaption     =   "Cashier_Audit.frx":152C
      textLT          =   "Cashier_Audit.frx":158E
      textCT          =   "Cashier_Audit.frx":15A6
      textRT          =   "Cashier_Audit.frx":15BE
      textLM          =   "Cashier_Audit.frx":15D6
      textRM          =   "Cashier_Audit.frx":15EE
      textLB          =   "Cashier_Audit.frx":1606
      textCB          =   "Cashier_Audit.frx":161E
      textRB          =   "Cashier_Audit.frx":1636
      colorBack       =   "Cashier_Audit.frx":164E
      colorIntern     =   "Cashier_Audit.frx":1678
      colorMO         =   "Cashier_Audit.frx":16A2
      colorFocus      =   "Cashier_Audit.frx":16CC
      colorDisabled   =   "Cashier_Audit.frx":16F6
      colorPressed    =   "Cashier_Audit.frx":1720
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
   Begin BTNENHLib4.BtnEnh apply 
      Height          =   735
      Left            =   480
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
      _Version        =   589827
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Clear"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      textCaption     =   "Cashier_Audit.frx":174A
      textLT          =   "Cashier_Audit.frx":17B4
      textCT          =   "Cashier_Audit.frx":17CC
      textRT          =   "Cashier_Audit.frx":17E4
      textLM          =   "Cashier_Audit.frx":17FC
      textRM          =   "Cashier_Audit.frx":1814
      textLB          =   "Cashier_Audit.frx":182C
      textCB          =   "Cashier_Audit.frx":1844
      textRB          =   "Cashier_Audit.frx":185C
      colorBack       =   "Cashier_Audit.frx":1874
      colorIntern     =   "Cashier_Audit.frx":189E
      colorMO         =   "Cashier_Audit.frx":18C8
      colorFocus      =   "Cashier_Audit.frx":18F2
      colorDisabled   =   "Cashier_Audit.frx":191C
      colorPressed    =   "Cashier_Audit.frx":1946
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
      Office2003ColorNorth=   8618883
      Office2003ColorSouth=   5855577
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   8618883
      Office2003ColorSouthMouseOver=   8618883
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   5395026
      Office2003ColorSouthPressed=   5855577
      Office2003ColorBorderPressed=   16777215
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand2 
      Height          =   615
      Left            =   2400
      TabIndex        =   17
      Top             =   6360
      Width           =   1935
      _Version        =   589827
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Cashier_Audit.frx":1970
      textLT          =   "Cashier_Audit.frx":19D8
      textCT          =   "Cashier_Audit.frx":19F0
      textRT          =   "Cashier_Audit.frx":1A08
      textLM          =   "Cashier_Audit.frx":1A20
      textRM          =   "Cashier_Audit.frx":1A38
      textLB          =   "Cashier_Audit.frx":1A50
      textCB          =   "Cashier_Audit.frx":1A68
      textRB          =   "Cashier_Audit.frx":1A80
      colorBack       =   "Cashier_Audit.frx":1A98
      colorIntern     =   "Cashier_Audit.frx":1AC2
      colorMO         =   "Cashier_Audit.frx":1AEC
      colorFocus      =   "Cashier_Audit.frx":1B16
      colorDisabled   =   "Cashier_Audit.frx":1B40
      colorPressed    =   "Cashier_Audit.frx":1B6A
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
   Begin BTNENHLib4.BtnEnh SSCommand1 
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   2055
      _Version        =   589827
      _ExtentX        =   3625
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Start Audit"
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
      textCaption     =   "Cashier_Audit.frx":1B94
      textLT          =   "Cashier_Audit.frx":1C0A
      textCT          =   "Cashier_Audit.frx":1C22
      textRT          =   "Cashier_Audit.frx":1C3A
      textLM          =   "Cashier_Audit.frx":1C52
      textRM          =   "Cashier_Audit.frx":1C6A
      textLB          =   "Cashier_Audit.frx":1C82
      textCB          =   "Cashier_Audit.frx":1C9A
      textRB          =   "Cashier_Audit.frx":1CB2
      colorBack       =   "Cashier_Audit.frx":1CCA
      colorIntern     =   "Cashier_Audit.frx":1CF4
      colorMO         =   "Cashier_Audit.frx":1D1E
      colorFocus      =   "Cashier_Audit.frx":1D48
      colorDisabled   =   "Cashier_Audit.frx":1D72
      colorPressed    =   "Cashier_Audit.frx":1D9C
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
   Begin VB.Shape Shape6 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   360
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label DrawerAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   240
      Top             =   360
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Beginning Drawer Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4215
      Left            =   240
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   6015
      Left            =   120
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Cashier_Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Apply_Click()
DrawerAmt.Caption = ""
End Sub

Private Sub Form_Activate()
If Cashier.Text = "" Then
  See_Msg "Please Exit You Cannot Start an Audit!!", 0, ""
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
ResizeScreen Me, False
CenterForm Me
'MarkRecordsClosed ""
If Val(GetConfigInfo("Cashier Station", "Begin Drawer Amount", 0, "FUNCTIONS.INI")) > 0 Then
 DrawerAmt.Caption = Format(Val(GetConfigInfo("Cashier Station", "Begin Drawer Amount", 0, "FUNCTIONS.INI")), "0.00")
Else
 If AuditCalculator(11) > 0 Then
   DrawerAmt.Caption = Format(AuditCalculator(11), "0.00")
 End If
End If
AuditCalculator(11) = 0
Find_What = "Cashier Audit"
Load_What = "Servers"
Dim ServersName As String
If GetSelected = "" Then Select_What.Show 1
If GetSelected = "" Then Exit Sub
Cashier.Text = GetSelected
GetSelected = ""
Dim Db As ADODB.Connection
Dim Id As ADODB.Recordset
OpenDb MyRegisters(0).RegLoc & "Employee.mdb", Db
 Sql = "Select * from [General Data] Where ISNULL(TRY_CONVERT(int,[Terminated]),0) = 0 order by [Last Name] asc"
ConnectADORS Sql, Db, Id
 Do While Not Id.EOF
    ServersName = Id(0) & " " & Id(2)
    LoggedOnAs(0).EmpAccessType = ServersName 'id(7)
    LoggedOnAs(0).EmpAccessLevel = Check_System_Access(CStr(Id(7)))
    If LoggedOnAs(0).EmpAccessLevel > 0 Then
     Cashier.AddItem ServersName '
    End If
 Id.MoveNext
 Loop
DbClose Db, Id
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form
  For Each Form In Forms
     If Form Is Me Then
       Set Form = Nothing
       Exit For
     End If
  Next Form
End Sub
Private Sub Number_Click(Index As Integer)
DrawerAmt.Caption = DrawerAmt.Caption & Trim(Str(Index))
End Sub
Private Sub SSCommand1_Click()
On Error Resume Next
If AuditStart(Cashier.Text, Val(DrawerAmt.Caption)) = True Then
 SSCommand2_Click
End If
Exit Sub
End Sub
Private Sub SSCommand2_Click()
Call MemoryCheck(Cashier_Audit)
End Sub
Private Sub SSCommand3_Click()
DrawerAmt.Caption = DrawerAmt.Caption & "."
End Sub
