VERSION 5.00
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Server_Report_Menu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Server Report Menu"
   ClientHeight    =   10815
   ClientLeft      =   210
   ClientTop       =   450
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh SSCommand8 
      Height          =   855
      Left            =   4320
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Server Audits"
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
      textCaption     =   "Server_Report_Menu.frx":0000
      textLT          =   "Server_Report_Menu.frx":007A
      textCT          =   "Server_Report_Menu.frx":0092
      textRT          =   "Server_Report_Menu.frx":00AA
      textLM          =   "Server_Report_Menu.frx":00C2
      textRM          =   "Server_Report_Menu.frx":00DA
      textLB          =   "Server_Report_Menu.frx":00F2
      textCB          =   "Server_Report_Menu.frx":010A
      textRB          =   "Server_Report_Menu.frx":0122
      colorBack       =   "Server_Report_Menu.frx":013A
      colorIntern     =   "Server_Report_Menu.frx":0164
      colorMO         =   "Server_Report_Menu.frx":018E
      colorFocus      =   "Server_Report_Menu.frx":01B8
      colorDisabled   =   "Server_Report_Menu.frx":01E2
      colorPressed    =   "Server_Report_Menu.frx":020C
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
   Begin BTNENHLib4.BtnEnh exit 
      Height          =   840
      Left            =   5880
      TabIndex        =   2
      Top             =   7200
      Width           =   3240
      _Version        =   589828
      _ExtentX        =   5715
      _ExtentY        =   1482
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
      PictureCustom   =   "C:\Users\ben\Downloads\flat_button_no_text.png"
      BackColorContainer=   12632256
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   2
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPictureCustom=   "Server_Report_Menu.frx":0236
      textCaption     =   "Server_Report_Menu.frx":02E7
      textLT          =   "Server_Report_Menu.frx":034F
      textCT          =   "Server_Report_Menu.frx":0367
      textRT          =   "Server_Report_Menu.frx":037F
      textLM          =   "Server_Report_Menu.frx":0397
      textRM          =   "Server_Report_Menu.frx":03AF
      textLB          =   "Server_Report_Menu.frx":03C7
      textCB          =   "Server_Report_Menu.frx":03DF
      textRB          =   "Server_Report_Menu.frx":03F7
      colorBack       =   "Server_Report_Menu.frx":040F
      colorIntern     =   "Server_Report_Menu.frx":0439
      colorMO         =   "Server_Report_Menu.frx":0463
      colorFocus      =   "Server_Report_Menu.frx":048D
      colorDisabled   =   "Server_Report_Menu.frx":04B7
      colorPressed    =   "Server_Report_Menu.frx":04E1
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
      Office2003ColorNorth=   128
      Office2003ColorSouth=   4210816
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   4210816
      Office2003ColorSouthMouseOver=   4210816
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   12870144
      Office2003ColorSouthPressed=   8404992
      Office2003ColorBorderPressed=   16777215
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand2 
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   $"Server_Report_Menu.frx":050B
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
      textCaption     =   "Server_Report_Menu.frx":0528
      textLT          =   "Server_Report_Menu.frx":05BA
      textCT          =   "Server_Report_Menu.frx":05D2
      textRT          =   "Server_Report_Menu.frx":05EA
      textLM          =   "Server_Report_Menu.frx":0602
      textRM          =   "Server_Report_Menu.frx":061A
      textLB          =   "Server_Report_Menu.frx":0632
      textCB          =   "Server_Report_Menu.frx":064A
      textRB          =   "Server_Report_Menu.frx":0662
      colorBack       =   "Server_Report_Menu.frx":067A
      colorIntern     =   "Server_Report_Menu.frx":06A4
      colorMO         =   "Server_Report_Menu.frx":06CE
      colorFocus      =   "Server_Report_Menu.frx":06F8
      colorDisabled   =   "Server_Report_Menu.frx":0722
      colorPressed    =   "Server_Report_Menu.frx":074C
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
   Begin BTNENHLib4.BtnEnh SSCommand25 
      Height          =   855
      Left            =   4320
      TabIndex        =   5
      Top             =   5160
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   $"Server_Report_Menu.frx":0776
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
      textCaption     =   "Server_Report_Menu.frx":078E
      textLT          =   "Server_Report_Menu.frx":0816
      textCT          =   "Server_Report_Menu.frx":082E
      textRT          =   "Server_Report_Menu.frx":0846
      textLM          =   "Server_Report_Menu.frx":085E
      textRM          =   "Server_Report_Menu.frx":0876
      textLB          =   "Server_Report_Menu.frx":088E
      textCB          =   "Server_Report_Menu.frx":08A6
      textRB          =   "Server_Report_Menu.frx":08BE
      colorBack       =   "Server_Report_Menu.frx":08D6
      colorIntern     =   "Server_Report_Menu.frx":0900
      colorMO         =   "Server_Report_Menu.frx":092A
      colorFocus      =   "Server_Report_Menu.frx":0954
      colorDisabled   =   "Server_Report_Menu.frx":097E
      colorPressed    =   "Server_Report_Menu.frx":09A8
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
   Begin BTNENHLib4.BtnEnh SSCommand6 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   $"Server_Report_Menu.frx":09D2
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
      textCaption     =   "Server_Report_Menu.frx":09E9
      textLT          =   "Server_Report_Menu.frx":0A6F
      textCT          =   "Server_Report_Menu.frx":0A87
      textRT          =   "Server_Report_Menu.frx":0A9F
      textLM          =   "Server_Report_Menu.frx":0AB7
      textRM          =   "Server_Report_Menu.frx":0ACF
      textLB          =   "Server_Report_Menu.frx":0AE7
      textCB          =   "Server_Report_Menu.frx":0AFF
      textRB          =   "Server_Report_Menu.frx":0B17
      colorBack       =   "Server_Report_Menu.frx":0B2F
      colorIntern     =   "Server_Report_Menu.frx":0B59
      colorMO         =   "Server_Report_Menu.frx":0B83
      colorFocus      =   "Server_Report_Menu.frx":0BAD
      colorDisabled   =   "Server_Report_Menu.frx":0BD7
      colorPressed    =   "Server_Report_Menu.frx":0C01
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
   Begin BTNENHLib4.BtnEnh ChangeSections 
      Height          =   855
      Left            =   7560
      TabIndex        =   1
      Top             =   5160
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   $"Server_Report_Menu.frx":0C2B
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
      textCaption     =   "Server_Report_Menu.frx":0C4C
      textLT          =   "Server_Report_Menu.frx":0CE6
      textCT          =   "Server_Report_Menu.frx":0CFE
      textRT          =   "Server_Report_Menu.frx":0D16
      textLM          =   "Server_Report_Menu.frx":0D2E
      textRM          =   "Server_Report_Menu.frx":0D46
      textLB          =   "Server_Report_Menu.frx":0D5E
      textCB          =   "Server_Report_Menu.frx":0D76
      textRB          =   "Server_Report_Menu.frx":0D8E
      colorBack       =   "Server_Report_Menu.frx":0DA6
      colorIntern     =   "Server_Report_Menu.frx":0DD0
      colorMO         =   "Server_Report_Menu.frx":0DFA
      colorFocus      =   "Server_Report_Menu.frx":0E24
      colorDisabled   =   "Server_Report_Menu.frx":0E4E
      colorPressed    =   "Server_Report_Menu.frx":0E78
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
      Left            =   7560
      TabIndex        =   6
      Top             =   2280
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Cash Deposit"
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
      textCaption     =   "Server_Report_Menu.frx":0EA2
      textLT          =   "Server_Report_Menu.frx":0F1A
      textCT          =   "Server_Report_Menu.frx":0F32
      textRT          =   "Server_Report_Menu.frx":0F4A
      textLM          =   "Server_Report_Menu.frx":0F62
      textRM          =   "Server_Report_Menu.frx":0F7A
      textLB          =   "Server_Report_Menu.frx":0F92
      textCB          =   "Server_Report_Menu.frx":0FAA
      textRB          =   "Server_Report_Menu.frx":0FC2
      colorBack       =   "Server_Report_Menu.frx":0FDA
      colorIntern     =   "Server_Report_Menu.frx":1004
      colorMO         =   "Server_Report_Menu.frx":102E
      colorFocus      =   "Server_Report_Menu.frx":1058
      colorDisabled   =   "Server_Report_Menu.frx":1082
      colorPressed    =   "Server_Report_Menu.frx":10AC
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
   Begin BTNENHLib4.BtnEnh SSCommand3 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   $"Server_Report_Menu.frx":10D6
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
      textCaption     =   "Server_Report_Menu.frx":10F4
      textLT          =   "Server_Report_Menu.frx":1188
      textCT          =   "Server_Report_Menu.frx":11A0
      textRT          =   "Server_Report_Menu.frx":11B8
      textLM          =   "Server_Report_Menu.frx":11D0
      textRM          =   "Server_Report_Menu.frx":11E8
      textLB          =   "Server_Report_Menu.frx":1200
      textCB          =   "Server_Report_Menu.frx":1218
      textRB          =   "Server_Report_Menu.frx":1230
      colorBack       =   "Server_Report_Menu.frx":1248
      colorIntern     =   "Server_Report_Menu.frx":1272
      colorMO         =   "Server_Report_Menu.frx":129C
      colorFocus      =   "Server_Report_Menu.frx":12C6
      colorDisabled   =   "Server_Report_Menu.frx":12F0
      colorPressed    =   "Server_Report_Menu.frx":131A
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
   Begin BTNENHLib4.BtnEnh SSCommand15 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Server Income"
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
      textCaption     =   "Server_Report_Menu.frx":1344
      textLT          =   "Server_Report_Menu.frx":13BE
      textCT          =   "Server_Report_Menu.frx":13D6
      textRT          =   "Server_Report_Menu.frx":13EE
      textLM          =   "Server_Report_Menu.frx":1406
      textRM          =   "Server_Report_Menu.frx":141E
      textLB          =   "Server_Report_Menu.frx":1436
      textCB          =   "Server_Report_Menu.frx":144E
      textRB          =   "Server_Report_Menu.frx":1466
      colorBack       =   "Server_Report_Menu.frx":147E
      colorIntern     =   "Server_Report_Menu.frx":14A8
      colorMO         =   "Server_Report_Menu.frx":14D2
      colorFocus      =   "Server_Report_Menu.frx":14FC
      colorDisabled   =   "Server_Report_Menu.frx":1526
      colorPressed    =   "Server_Report_Menu.frx":1550
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
   Begin BTNENHLib4.BtnEnh SSCommand7 
      Height          =   855
      Left            =   6720
      TabIndex        =   9
      Top             =   120
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
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
      textCaption     =   "Server_Report_Menu.frx":157A
      textLT          =   "Server_Report_Menu.frx":15F6
      textCT          =   "Server_Report_Menu.frx":160E
      textRT          =   "Server_Report_Menu.frx":1626
      textLM          =   "Server_Report_Menu.frx":163E
      textRM          =   "Server_Report_Menu.frx":1656
      textLB          =   "Server_Report_Menu.frx":166E
      textCB          =   "Server_Report_Menu.frx":1686
      textRB          =   "Server_Report_Menu.frx":169E
      colorBack       =   "Server_Report_Menu.frx":16B6
      colorIntern     =   "Server_Report_Menu.frx":16E0
      colorMO         =   "Server_Report_Menu.frx":170A
      colorFocus      =   "Server_Report_Menu.frx":1734
      colorDisabled   =   "Server_Report_Menu.frx":175E
      colorPressed    =   "Server_Report_Menu.frx":1788
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
   Begin BTNENHLib4.BtnEnh SSCommand1 
      Height          =   855
      Left            =   9960
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Payment Report"
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
      textCaption     =   "Server_Report_Menu.frx":17B2
      textLT          =   "Server_Report_Menu.frx":182E
      textCT          =   "Server_Report_Menu.frx":1846
      textRT          =   "Server_Report_Menu.frx":185E
      textLM          =   "Server_Report_Menu.frx":1876
      textRM          =   "Server_Report_Menu.frx":188E
      textLB          =   "Server_Report_Menu.frx":18A6
      textCB          =   "Server_Report_Menu.frx":18BE
      textRB          =   "Server_Report_Menu.frx":18D6
      colorBack       =   "Server_Report_Menu.frx":18EE
      colorIntern     =   "Server_Report_Menu.frx":1918
      colorMO         =   "Server_Report_Menu.frx":1942
      colorFocus      =   "Server_Report_Menu.frx":196C
      colorDisabled   =   "Server_Report_Menu.frx":1996
      colorPressed    =   "Server_Report_Menu.frx":19C0
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
   Begin BTNENHLib4.BtnEnh SSCommand4 
      Height          =   855
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Department Sales"
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
      textCaption     =   "Server_Report_Menu.frx":19EA
      textLT          =   "Server_Report_Menu.frx":1A6A
      textCT          =   "Server_Report_Menu.frx":1A82
      textRT          =   "Server_Report_Menu.frx":1A9A
      textLM          =   "Server_Report_Menu.frx":1AB2
      textRM          =   "Server_Report_Menu.frx":1ACA
      textLB          =   "Server_Report_Menu.frx":1AE2
      textCB          =   "Server_Report_Menu.frx":1AFA
      textRB          =   "Server_Report_Menu.frx":1B12
      colorBack       =   "Server_Report_Menu.frx":1B2A
      colorIntern     =   "Server_Report_Menu.frx":1B54
      colorMO         =   "Server_Report_Menu.frx":1B7E
      colorFocus      =   "Server_Report_Menu.frx":1BA8
      colorDisabled   =   "Server_Report_Menu.frx":1BD2
      colorPressed    =   "Server_Report_Menu.frx":1BFC
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
   Begin BTNENHLib4.BtnEnh SSCommand5 
      Height          =   855
      Left            =   4320
      TabIndex        =   12
      Top             =   4200
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Drawer Audits"
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
      textCaption     =   "Server_Report_Menu.frx":1C26
      textLT          =   "Server_Report_Menu.frx":1CA0
      textCT          =   "Server_Report_Menu.frx":1CB8
      textRT          =   "Server_Report_Menu.frx":1CD0
      textLM          =   "Server_Report_Menu.frx":1CE8
      textRM          =   "Server_Report_Menu.frx":1D00
      textLB          =   "Server_Report_Menu.frx":1D18
      textCB          =   "Server_Report_Menu.frx":1D30
      textRB          =   "Server_Report_Menu.frx":1D48
      colorBack       =   "Server_Report_Menu.frx":1D60
      colorIntern     =   "Server_Report_Menu.frx":1D8A
      colorMO         =   "Server_Report_Menu.frx":1DB4
      colorFocus      =   "Server_Report_Menu.frx":1DDE
      colorDisabled   =   "Server_Report_Menu.frx":1E08
      colorPressed    =   "Server_Report_Menu.frx":1E32
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
   Begin BTNENHLib4.BtnEnh SSCommand13 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Voids Today by Department"
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
      textCaption     =   "Server_Report_Menu.frx":1E5C
      textLT          =   "Server_Report_Menu.frx":1EEE
      textCT          =   "Server_Report_Menu.frx":1F06
      textRT          =   "Server_Report_Menu.frx":1F1E
      textLM          =   "Server_Report_Menu.frx":1F36
      textRM          =   "Server_Report_Menu.frx":1F4E
      textLB          =   "Server_Report_Menu.frx":1F66
      textCB          =   "Server_Report_Menu.frx":1F7E
      textRB          =   "Server_Report_Menu.frx":1F96
      colorBack       =   "Server_Report_Menu.frx":1FAE
      colorIntern     =   "Server_Report_Menu.frx":1FD8
      colorMO         =   "Server_Report_Menu.frx":2002
      colorFocus      =   "Server_Report_Menu.frx":202C
      colorDisabled   =   "Server_Report_Menu.frx":2056
      colorPressed    =   "Server_Report_Menu.frx":2080
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
      Left            =   7560
      TabIndex        =   14
      Top             =   3240
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Tips Withheld"
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
      textCaption     =   "Server_Report_Menu.frx":20AA
      textLT          =   "Server_Report_Menu.frx":2124
      textCT          =   "Server_Report_Menu.frx":213C
      textRT          =   "Server_Report_Menu.frx":2154
      textLM          =   "Server_Report_Menu.frx":216C
      textRM          =   "Server_Report_Menu.frx":2184
      textLB          =   "Server_Report_Menu.frx":219C
      textCB          =   "Server_Report_Menu.frx":21B4
      textRB          =   "Server_Report_Menu.frx":21CC
      colorBack       =   "Server_Report_Menu.frx":21E4
      colorIntern     =   "Server_Report_Menu.frx":220E
      colorMO         =   "Server_Report_Menu.frx":2238
      colorFocus      =   "Server_Report_Menu.frx":2262
      colorDisabled   =   "Server_Report_Menu.frx":228C
      colorPressed    =   "Server_Report_Menu.frx":22B6
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
      Left            =   7560
      TabIndex        =   16
      Top             =   4200
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Tip Share"
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
      textCaption     =   "Server_Report_Menu.frx":22E0
      textLT          =   "Server_Report_Menu.frx":2352
      textCT          =   "Server_Report_Menu.frx":236A
      textRT          =   "Server_Report_Menu.frx":2382
      textLM          =   "Server_Report_Menu.frx":239A
      textRM          =   "Server_Report_Menu.frx":23B2
      textLB          =   "Server_Report_Menu.frx":23CA
      textCB          =   "Server_Report_Menu.frx":23E2
      textRB          =   "Server_Report_Menu.frx":23FA
      colorBack       =   "Server_Report_Menu.frx":2412
      colorIntern     =   "Server_Report_Menu.frx":243C
      colorMO         =   "Server_Report_Menu.frx":2466
      colorFocus      =   "Server_Report_Menu.frx":2490
      colorDisabled   =   "Server_Report_Menu.frx":24BA
      colorPressed    =   "Server_Report_Menu.frx":24E4
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
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Voids Today By Detail"
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
      textCaption     =   "Server_Report_Menu.frx":250E
      textLT          =   "Server_Report_Menu.frx":2598
      textCT          =   "Server_Report_Menu.frx":25B0
      textRT          =   "Server_Report_Menu.frx":25C8
      textLM          =   "Server_Report_Menu.frx":25E0
      textRM          =   "Server_Report_Menu.frx":25F8
      textLB          =   "Server_Report_Menu.frx":2610
      textCB          =   "Server_Report_Menu.frx":2628
      textRB          =   "Server_Report_Menu.frx":2640
      colorBack       =   "Server_Report_Menu.frx":2658
      colorIntern     =   "Server_Report_Menu.frx":2682
      colorMO         =   "Server_Report_Menu.frx":26AC
      colorFocus      =   "Server_Report_Menu.frx":26D6
      colorDisabled   =   "Server_Report_Menu.frx":2700
      colorPressed    =   "Server_Report_Menu.frx":272A
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
   Begin BTNENHLib4.BtnEnh BtnEnh5 
      Height          =   855
      Left            =   4320
      TabIndex        =   18
      Top             =   6120
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Reprint Server Audit"
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
      textCaption     =   "Server_Report_Menu.frx":2754
      textLT          =   "Server_Report_Menu.frx":27DC
      textCT          =   "Server_Report_Menu.frx":27F4
      textRT          =   "Server_Report_Menu.frx":280C
      textLM          =   "Server_Report_Menu.frx":2824
      textRM          =   "Server_Report_Menu.frx":283C
      textLB          =   "Server_Report_Menu.frx":2854
      textCB          =   "Server_Report_Menu.frx":286C
      textRB          =   "Server_Report_Menu.frx":2884
      colorBack       =   "Server_Report_Menu.frx":289C
      colorIntern     =   "Server_Report_Menu.frx":28C6
      colorMO         =   "Server_Report_Menu.frx":28F0
      colorFocus      =   "Server_Report_Menu.frx":291A
      colorDisabled   =   "Server_Report_Menu.frx":2944
      colorPressed    =   "Server_Report_Menu.frx":296E
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
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Sever Table Report"
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
      textCaption     =   "Server_Report_Menu.frx":2998
      textLT          =   "Server_Report_Menu.frx":2A1C
      textCT          =   "Server_Report_Menu.frx":2A34
      textRT          =   "Server_Report_Menu.frx":2A4C
      textLM          =   "Server_Report_Menu.frx":2A64
      textRM          =   "Server_Report_Menu.frx":2A7C
      textLB          =   "Server_Report_Menu.frx":2A94
      textCB          =   "Server_Report_Menu.frx":2AAC
      textRB          =   "Server_Report_Menu.frx":2AC4
      colorBack       =   "Server_Report_Menu.frx":2ADC
      colorIntern     =   "Server_Report_Menu.frx":2B06
      colorMO         =   "Server_Report_Menu.frx":2B30
      colorFocus      =   "Server_Report_Menu.frx":2B5A
      colorDisabled   =   "Server_Report_Menu.frx":2B84
      colorPressed    =   "Server_Report_Menu.frx":2BAE
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
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   4080
      Top             =   2040
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Server Reports Menu"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   6375
      Left            =   3840
      Top             =   1800
      Width           =   7335
   End
End
Attribute VB_Name = "Server_Report_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipHold
   Revenue_Center   As String
   Tip_Amount       As Currency
   Day_Of_Week      As String
End Type
Public Enum RptName
    TipOutWithHeld
    TipWithheld
End Enum
Private HTip        As TipHold
Private SVRRPTS     As RptName
Private PaymentType As String
Private Sub GetCreditCardReport()
On Error GoTo ErrorHandler
Dim TipSales As Currency
Dim ChgSales As Currency
Dim TipDecl  As Currency
Dim ChgDecl  As Currency
Dim GratDecl As Currency
Dim TipOut   As Currency
Dim TShare   As Currency
Dim SL       As ADODB.Connection
Dim SRd      As ADODB.Recordset
Dim Data     As String
      TGrat = 0: TSale = 0: Tpayments = 0: SSales = 0
      ConnectADODB "MSaleData", SL
        'Sql = "delete from [Tip Income] where [Date] = ''"
        'SL.Execute Sql
        Sql = "Select Distinct [Server Name], "
        Sql = Sql & " Sum([Tip Sales]) as TSales,"
        Sql = Sql & " Sum([Declared Tips]) as TDeclared,"
        Sql = Sql & " Sum([Charged Tips]) as TCharged,"
        Sql = Sql & " Sum([Gratuity]) as TGratuity,"
        Sql = Sql & " Sum([Tips Paid Out]) as TipsPaid,"
        Sql = Sql & " Sum([Charge Sales]) as CSales,"
        Sql = Sql & " Sum([Tip Share]) as TShare"
        Sql = Sql & " from [Tip Income] "
        Sql = Sql & " Where TRY_CAST([Date] AS date)  Between '" & CDate(SearchDates.Item(1).StartDate) & "'"
        Sql = Sql & " and '" & CDate(SearchDates.Item(1).EndDate) & "'"
        Sql = Sql & " Group By [Server Name]"
      ConnectADORS Sql, SL, SRd
      If SRd.EOF Then
       DbClose SL, SRd
       See_Msg "No Data Found", 5, ""
       Exit Sub
      End If
      SetPage ReportView.View, "Server Tip /Sales Income " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
      Dim F$
      F = "+<2500|+>1100|+>1000|+>1000|+>1000|+>1000|+>1000|+>1100|+>1000;"
      Dim H$
      H = "Server|Tip Sales|Chg Sales|Cash Tips|Chg Tips|Tip Pd Out|Tip Share|Total Tips|Chg Tip%;"
      Dim i%
      TipSales = 0: ChgSales = 0: TipDecl = 0: ChgDecl = 0: TipOut = 0: Data = ""
      Do While Not SRd.EOF
       If SRd![Server Name] & "" > "" Then
            Data = Data & SRd![Server Name] & "|"
            Data = Data & Format(Val(SRd![TSales] & ""), "0.00") & "|"
            Data = Data & Format(Val(SRd![CSales] & ""), "0.00") & "|"
            Data = Data & Format(Val(SRd![TDeclared] & ""), "0.00") & "|"
            Data = Data & Format(Val(SRd![TCharged] & "") + Val(SRd![TGratuity] & ""), "0.00") & "|"
            Data = Data & Format(Val(SRd![TipsPaid] & ""), "0.00") & "|"
            Data = Data & Format(Val(SRd![TShare] & ""), "0.00") & "|"
            Data = Data & Format((((SRd![TDeclared] + SRd![TShare]) + SRd![TCharged]) + SRd![TGratuity]) - Val(SRd![TipsPaid] & ""), "0.00") & "|"
            If Val(SRd![CSales] & "") > 0 And Val(SRd![TCharged] & "") > 0 Then
             Data = Data & Format((Val(SRd![TCharged] & "") / Val(SRd![CSales] & "") * 100), "0.00") & ";"
            Else
             Data = Data & "" & ";"
            End If
            TipSales = TipSales + Val(SRd![TSales] & "")
            ChgSales = ChgSales + Val(SRd![CSales] & "")
            TipDecl = TipDecl + Val(SRd![TDeclared] & "") + Val(SRd![TShare] & "")
            ChgDecl = ChgDecl + (Val(SRd![TCharged] & "") + Val(SRd![TGratuity] & ""))
            TipOut = TipOut + Val(SRd![TipsPaid] & "")
            TShare = TShare + Val(SRd![TShare] & "")
        End If
        SRd.MoveNext
       Loop
       Data = Data & "Totals" & "|"
       Data = Data & Format(TipSales, "0.00") & "|"
       Data = Data & Format(ChgSales, "0.00") & "|"
       Data = Data & Format(TipDecl, "0.00") & "|"
       Data = Data & Format(ChgDecl, "0.00") & "|"
       Data = Data & Format(TipOut, "0.00") & "|"
       Data = Data & Format(TShare, "0.00") & "|"
       Data = Data & Format(((TipDecl + ChgDecl)) - TipOut, "0.00") & "|"
       If ChgSales > 0 And ChgDecl > 0 Then
       Data = Data & Format(((ChgDecl / ChgSales)) * 100, "0.00") & ";"
       Else
       Data = Data & 0 & ";"
       End If
       DbClose SL, SRd
       ReportView.View.TableBorder = tbTopBottom
       ReportView.View.addTable F, H, Data, vbBlack, vbWhite
       FinishJob 1
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Credit Card Report ")
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
 LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
 If SearchDates Is Nothing Then
  Exit Sub
 End If
 ShowTipsWithheld SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, TipOutWithHeld
End Sub
Private Sub BtnEnh2_Click()
If Can_Continue(BtnEnh2.Caption) = False Then Exit Sub
If Check_Option(237) = 1 Then
    LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
    If SearchDates Is Nothing Then
     Exit Sub
    End If
    ShowTipsWithheld SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, TipWithheld
Else
    CenterObject Me, TipPooling
    TipPooling.Visible = True
    DOW(0).Clear: DOW(1).Clear
    DOW(0).AddItem "Today": DOW(0).Text = "Today"
    DOW(1).AddItem "0", DOW(0).NewIndex
    Dim Db As ADODB.Connection
    Dim Tb As ADODB.Recordset
     ConnectADODB CheckStatDb, Db
       Sql = "Select Distinct [Date],[Close Out Day] from [Payments] "
       Sql = Sql & "Group By [date],[Close Out Day]"
       ConnectADORS Sql, Db, Tb
       Do While Not Tb.EOF
        If IsCombo(DOW(0), Format(Tb![Date], "dddd")) = False Then
         DOW(0).AddItem Format(Tb![Date], "dddd")
         DOW(1).AddItem Tb![Close out day], DOW(0).NewIndex
        End If
        Tb.MoveNext
       Loop
    DbClose Db, Tb
End If
End Sub
Private Sub BtnEnh3_Click()
End Sub
Private Sub BtnEnh4_Click()
  Dim ScreenReport As Boolean
  ScreenReport = True
  If Check_Option(96) = 0 Then
     If See_Msg("Print Report", 4, "") = 6 Then
        ScreenReport = False
     End If
  End If
Load_What = "AllServers"
Find_What = "Select a Server"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
Void_Report "Day", GetSelected, ScreenReport, False
End Sub
Private Sub BtnEnh5_Click()
LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
Load_What = "Reprint Audits"
Find_What = "Select a Server"
Select_What.Show 1
If GetSelected = "" Then Exit Sub
AuditHistory GetSelected, 0, SearchDates.Item(1).StartDate, ""
End Sub
Private Sub BtnEnh6_Click()
End Sub
Private Sub BtnEnh8_Click()
 LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
 If SearchDates Is Nothing Then
  Exit Sub
 End If
'On Error GoTo errorhandler
LoadMealPeriods
Set SvrReports = New AllMealPeriods
Dim MP              As Integer
Dim Db              As ADODB.Connection
Dim TDb             As ADODB.Connection
Dim Tb              As ADODB.Recordset
Dim TotalTables     As Integer
Dim TotalCustomer   As Integer
   'Sql = "Select * from MSaleData.dbo.[Closed Checks] as DB1 Inner Join [Table LayoutSQL].dbo.[Layout] as DB2"
   'Sql = Sql & " On DB1.[Table Number] = DB2.[Table Name];"
   'Sql = "Select * from [Closed Checks] "
   'Sql = Sql & " Where TRY_CAST([Order Date] AS date) + TRY_CAST([Open Time] AS datetime)  Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
   'Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
   'Sql = Sql & " and [Table Number] In( Select [Table Name] from [Table LayoutSQL].dbo.[Layout]) "
   'Sql = Sql & " Order by [Server Name] asc"
 
  ConnectADODB "MSaleData", Db
   Sql = "Select * from [Closed Checks] "
   Sql = Sql & " Where TRY_CAST([Order Date] AS date) + TRY_CAST([Open Time] AS datetime)  Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
   Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
   Sql = Sql & " and [Table Number] In( Select [Table Name] from [MSaleData].dbo.[Layout]) "
   Sql = Sql & " Order by [Server Name] asc"
  ConnectADORS Sql, Db, Tb
   If Tb.EOF Then
    DbClose Db, Tb
    See_Msg "No Sale Found", 5, ""
    DbClose Db, Tb
    Exit Sub
   End If
  Do While Not Tb.EOF
      'If IsDate(Tb.Fields("Close Time").Value) Then
       If SvrReports.Item(Tb.fields("Server Name").Value & "") Is Nothing Then
         SvrReports.add Tb.fields("Server Name").Value & ""
       End If
       SvrReports.Item(Tb.fields("Server Name").Value & "").Server = Tb.fields("Server Name").Value & ""
       If Hour(Tb.fields("Close Time").Value) >= BusinessPeriods(0).PeriodStartHour And Hour(Tb.fields("Close Time").Value) <= BusinessPeriods(0).PeriodEndHour Then
         SvrReports.Item(Tb.fields("Server Name").Value & "").Breakfast = "BreakFast  " & Format(BusinessPeriods(0).PeriodStart, "HH:MM AMPM") & " " & Format(BusinessPeriods(0).PeriodEnd, "HH:MM ampm")
         SvrReports.Item(Tb.fields("Server Name").Value & "").BreakfastCustomers = SvrReports.Item(Tb.fields("Server Name").Value & "").BreakfastCustomers + Tb.fields("Number In Party").Value
         SvrReports.Item(Tb.fields("Server Name").Value & "").BreadFastTables = SvrReports.Item(Tb.fields("Server Name").Value & "").BreadFastTables + 1
       
       ElseIf Hour(Tb.fields("Close Time").Value) >= BusinessPeriods(1).PeriodStartHour And Hour(Tb.fields("Close Time").Value) <= BusinessPeriods(1).PeriodEndHour Then
         SvrReports.Item(Tb.fields("Server Name").Value & "").Lunch = "Lunch " & Format(BusinessPeriods(1).PeriodStart, "HH:MM AMPM") & " " & Format(BusinessPeriods(1).PeriodEnd, "HH:MM ampm")
         SvrReports.Item(Tb.fields("Server Name").Value & "").LunchCustomers = SvrReports.Item(Tb.fields("Server Name").Value & "").LunchCustomers + Tb.fields("Number In Party").Value
         SvrReports.Item(Tb.fields("Server Name").Value & "").LunchTables = SvrReports.Item(Tb.fields("Server Name").Value & "").LunchTables + 1
       
       ElseIf Hour(Tb.fields("Close Time").Value) >= BusinessPeriods(2).PeriodStartHour And Hour(Tb.fields("Close Time").Value) <= BusinessPeriods(2).PeriodEndHour Then
         SvrReports.Item(Tb.fields("Server Name").Value & "").Mid = "Mid " & Format(BusinessPeriods(2).PeriodStart, "HH:MM AMPM") & " " & Format(BusinessPeriods(2).PeriodEnd, "HH:MM ampm")
         SvrReports.Item(Tb.fields("Server Name").Value & "").MidCustomers = SvrReports.Item(Tb.fields("Server Name").Value & "").MidCustomers + Tb.fields("Number In Party").Value
         SvrReports.Item(Tb.fields("Server Name").Value & "").MidTables = SvrReports.Item(Tb.fields("Server Name").Value & "").MidTables + 1
       
       ElseIf Hour(Tb.fields("Close Time").Value) >= BusinessPeriods(3).PeriodStartHour And Hour(Tb.fields("Close Time").Value) <= BusinessPeriods(3).PeriodEndHour Then
         SvrReports.Item(Tb.fields("Server Name").Value & "").Dinner = "Dinner " & Format(BusinessPeriods(3).PeriodStart, "HH:MM AMPM") & " " & Format(BusinessPeriods(3).PeriodEnd, "HH:MM ampm")
         SvrReports.Item(Tb.fields("Server Name").Value & "").DinnerCustomers = SvrReports.Item(Tb.fields("Server Name").Value & "").DinnerCustomers + Tb.fields("Number In Party").Value
         SvrReports.Item(Tb.fields("Server Name").Value & "").DinnerTables = SvrReports.Item(Tb.fields("Server Name").Value & "").DinnerTables + 1
       
       ElseIf Hour(Tb.fields("Close Time").Value) >= BusinessPeriods(4).PeriodStartHour And Hour(Tb.fields("Close Time").Value) <= BusinessPeriods(4).PeriodEndHour Then
         SvrReports.Item(Tb.fields("Server Name").Value & "").Late = "Late " & Format(BusinessPeriods(4).PeriodStart, "HH:MM AMPM") & " " & Format(BusinessPeriods(4).PeriodEnd, "HH:MM ampm")
         SvrReports.Item(Tb.fields("Server Name").Value & "").LateCustomers = SvrReports.Item(Tb.fields("Server Name").Value & "").LateCustomers + Tb.fields("Number In Party").Value
         SvrReports.Item(Tb.fields("Server Name").Value & "").LateTables = SvrReports.Item(Tb.fields("Server Name").Value & "").LateTables + 1
       End If
    'End If
    Trans_Log " Name " & Tb.fields("Server Name").Value & " Table Number " & Tb.fields("Table Number").Value
   Tb.MoveNext
  Loop
DbClose Db, Tb
Dim F$
F = "+<3000|+<4000|+^1500|+^1500;"
Dim H$
Dim Data As String
H = "Employee|Period Name|Total Tables|Total Customers;"
SetPage ReportView.View, "Table Report" & vbCrLf & " From " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
Data = ""
For MP = 1 To SvrReports.count
   TotalCustomer = 0: TotalTables = 0
   Data = ""
   Data = Data & SvrReports.Item(MP).Server & "|"
   If SvrReports.Item(MP).BreadFastTables = 0 Then
    Data = Data & "" & "|"
    Data = Data & "" & "|"
    Data = Data & "" & ";"
   Else
    Data = Data & SvrReports.Item(MP).Breakfast & "|"
    Data = Data & SvrReports.Item(MP).BreadFastTables & "|"
    Data = Data & SvrReports.Item(MP).BreakfastCustomers & ";"
   End If
   If SvrReports.Item(MP).LunchTables = 0 Then
    Data = Data & "" & "|"
    Data = Data & "" & "|"
    Data = Data & "" & ";"
   Else
    Data = Data & "" & "|"
    Data = Data & SvrReports.Item(MP).Lunch & "|"
    Data = Data & SvrReports.Item(MP).LunchTables & "|"
    Data = Data & SvrReports.Item(MP).LunchCustomers & ";"
   End If
   If SvrReports.Item(MP).MidTables = 0 Then
    Data = Data & "" & "|"
    Data = Data & "" & "|"
    Data = Data & "" & ";"
   Else
    Data = Data & "" & "|"
    Data = Data & SvrReports.Item(MP).Mid & "|"
    Data = Data & SvrReports.Item(MP).MidTables & "|"
    Data = Data & SvrReports.Item(MP).MidCustomers & ";"
   End If
   If SvrReports.Item(MP).DinnerTables = 0 Then
    Data = Data & "" & "|"
    Data = Data & "" & "|"
    Data = Data & "" & ";"
   Else
    Data = Data & "" & "|"
    Data = Data & SvrReports.Item(MP).Dinner & "|"
    Data = Data & SvrReports.Item(MP).DinnerTables & "|"
    Data = Data & SvrReports.Item(MP).DinnerCustomers & ";"
   End If
    If SvrReports.Item(MP).LateTables = 0 Then
    Data = Data & "" & "|"
    Data = Data & "" & "|"
    Data = Data & "" & ";"
   Else
    Data = Data & "" & "|"
    Data = Data & SvrReports.Item(MP).Late & "|"
    Data = Data & SvrReports.Item(MP).LateTables & "|"
    Data = Data & SvrReports.Item(MP).LateCustomers & ";"
   End If
   TotalTables = TotalTables + SvrReports.Item(MP).BreadFastTables + SvrReports.Item(MP).LunchTables + SvrReports.Item(MP).MidTables + SvrReports.Item(MP).DinnerTables + SvrReports.Item(MP).LateTables
   TotalCustomer = TotalCustomer + SvrReports.Item(MP).BreakfastCustomers + SvrReports.Item(MP).LunchCustomers + SvrReports.Item(MP).MidCustomers + SvrReports.Item(MP).DinnerCustomers + SvrReports.Item(MP).LateCustomers
   Data = Data & "Totals" & "|"
   Data = Data & "" & "|"
   Data = Data & TotalTables & "|"
   Data = Data & TotalCustomer & ";"
   ReportView.View.TableBorder = tbBoxColumns
   ReportView.View.addTable F, H, Data, vbBlack, vbWhite
Next MP
ReportView.View.EndDoc
FinishJob 1
Exit Sub
End Sub

Private Sub BtnEnh9_Click()
End Sub
Private Sub ChangeSections_Click()
If Can_Continue(ChangeSections.Caption) = False Then Exit Sub
Dim Db      As ADODB.Connection
Dim Records As ADODB.Recordset
Dim GEmp    As String
Load_What = "Drivers"
Find_What = "Select a Server"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
GEmp = GetSelected
GetSelected = ""
Load_What = "Work Area"
Find_What = "Select a Station"
'DoEvents
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
GV.WorkArea = GetSelected
GetSelected = ""
Sql = "Select * from [Check Numbers] where [Servers Name] = '" & GEmp & "'"
ConnectADODB "MSaleData", Db
ConnectADORS Sql, Db, Records
If Not Records.EOF Then
   Records![Work Area] = GV.WorkArea & ""
   Records.Update
End If
DbClose Db, Records
End Sub
Private Sub DOW_Click(Index As Integer)
TAmount.Text = "0.00"
TAmount.Refresh
If IsDate(STime.Text) And IsDate(ETime.Text) Then

Else
 Exit Sub
End If
Dim dw As Integer
If DOW(0).Text = "Today" Then
  dw = 0
Else
  dw = Val(DOW(1).List(DOW(0).ListIndex))
End If
TAmount.Text = Format(GetAllPoolTips(STime.Text, ETime.Text, dw), "0.00")
TAmount.Refresh
End Sub
Private Sub ETime_Click()
Dim InTime  As String
Dim PKT     As DateOrTime
Set PKT = New DateOrTime
InTime = PKT.ReturnTime
If IsDate(InTime) Then
  ETime.Text = Format(InTime, "H:mm ampm")
Else
  ETime.Text = ""
End If
Set PKT = Nothing
End Sub
Private Sub Exit_Click()
DontUnLoad = False
Call MemoryCheck(Update)
Call MemoryCheck(Server_Report_Menu)
End Sub
Private Sub Form_Load()
ResizeScreen1024 Me
If Check_Option(237) = 1 Then
  BtnEnh2.Caption = "Tips Withheld"
Else
  BtnEnh2.Caption = "Non Cash Tip Pool"
End If
If Check_Option(89) = 1 Then ChangeSections.Enabled = True
End Sub
Private Sub GetServerReport(WhichServer As String)
On Error GoTo ErrorHandler
Dim SL          As ADODB.Connection
Dim SRd         As ADODB.Recordset
Dim TGrat       As Currency
Dim TSale       As Currency
Dim Tpayments   As Currency
Dim Data        As String
Dim TCash       As Currency
OpenDb "MSaleData", SL
    Sql = "Select * from [Server Payments] "
    Sql = Sql & " Where [Server Name] = '" & WhichServer & "' and  TRY_CAST([Date] AS date)  Between '" & CDate(SearchDates.Item(1).StartDate) & "'"
    Sql = Sql & " and '" & CDate(SearchDates.Item(1).EndDate) & "'"
ConnectADORS Sql, SL, SRd
  If SRd.EOF Then
    See_Msg "No Payment Found", 5, ""
    DbClose SL, SRd
    Exit Sub
 End If
 SetPage ReportView.View, "Payment Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
 Dim H$
 H = WhichServer & "|Payment Type|Amount|Tips|Tip %;"
 Dim F$
 F = "+<1500|+<2000|+>1500|+>1000|+>1000;"
 TCash = 0
 Do While Not SRd.EOF
    Data = Data & "" & "|"
    Data = Data & SRd![Payment Type] & "|"
    If SRd.fields("Payment Type").Value = "Cash" Then
      TCash = TCash + Val(SRd![Payment Amount] & "")
    End If
    Data = Data & Format(SRd![Payment Amount], "0.00") & "|"
    Data = Data & Format(SRd![Tip Amount], "0.00") & "|"
    If SRd![Tip Amount] > 0 And SRd![Payment Amount] > 0 Then
     Data = Data & Format((SRd![Tip Amount] / SRd![Payment Amount]) * 100, "0.00") & ";"
    Else
     Data = Data & "" & ";"
    End If
    Tpayments = Tpayments + SRd![Payment Amount]
    TGrat = TGrat + SRd![Tip Amount]
    SRd.MoveNext
 Loop
 DbClose SL, SRd
 Data = Data & "" & "|"
 Data = Data & "Total Payments" & "|"
 Data = Data & Format(Tpayments, "0.00") & "|"
 Data = Data & Format(TGrat, "0.00") & ";"
 Data = Data & "" & "|"
 Data = Data & "Total Cash" & "|"
 Data = Data & Format(TCash, "0.00") & "|"
 Data = Data & ";"
 Data = Data & "" & "|"
 Data = Data & "Total Non Cash" & "|"
 Data = Data & Format(Tpayments - TCash, "0.00") & "|"
 Data = Data & ";"
 
 ReportView.View.TableBorder = tbColTopBottom
 ReportView.View.addTable F, H, Data, vbBlack, vbWhite
 FinishJob 1
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Server Report")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub ConstidateServerPayments()
On Error GoTo ErrorHandler
Dim SL          As ADODB.Connection
Dim SRd         As ADODB.Recordset
Dim Prs         As ADODB.Recordset
Dim TGrat       As Currency
Dim TSale       As Currency
Dim Tpayments   As Currency
Dim Data        As String
Dim FP          As Boolean
OpenDb "MSaleData", SL
    Sql = "Select Distinct [Server Name] from [Server Payments] "
    Sql = Sql & " Where TRY_CAST([Date] AS date)  Between '" & CDate(SearchDates.Item(1).StartDate) & "'"
    Sql = Sql & " and '" & CDate(SearchDates.Item(1).EndDate) & "'"
    Sql = Sql & " Group by [Server Name]"
ConnectADORS Sql, SL, SRd
  If SRd.EOF Then
    See_Msg "No Data Found", 5, ""
    DbClose SL, SRd
    Exit Sub
 End If
 SetPage ReportView.View, "Payment Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
 Dim H$
 H = "Server Name|Payment Type|Amount|Tips|Tip %;"
 Dim F$
 F = "+<3000|+<2000|+>1500|+>1000|+>1000;"
 Do While Not SRd.EOF
   Sql = "Select Distinct [Payment Type], Sum([Payment Amount]) as [PAmount], Sum([Tip Amount]) as [TAmount] from [Server Payments] "
   Sql = Sql & " Where [Server Name] = '" & SRd.fields("Server Name").Value & "' and  TRY_CAST([Date] AS date)  Between '" & CDate(SearchDates.Item(1).StartDate) & "'"
   Sql = Sql & " and '" & CDate(SearchDates.Item(1).EndDate) & "'"
   Sql = Sql & " Group by [Payment Type]"
   ConnectADORS Sql, SL, Prs
    FP = True
    Tpayments = 0: TGrat = 0
    Do While Not Prs.EOF
     If FP Then
      Data = Data & SRd.fields("Server Name").Value & "|"
      FP = False
     Else
      Data = Data & "|"
     End If
     Data = Data & Prs![Payment Type] & "|"
     Data = Data & Format(Prs![PAmount], "0.00") & "|"
     Data = Data & Format(Prs![TAmount], "0.00") & "|"
     If Prs![TAmount] > 0 And Prs![PAmount] > 0 Then
      Data = Data & Format((Prs![TAmount] / Prs![PAmount]) * 100, "0.00") & ";"
     Else
      Data = Data & "" & ";"
     End If
     Tpayments = Tpayments + Prs![PAmount]
     TGrat = TGrat + Prs![TAmount]
     Prs.MoveNext
     Loop
    SetNothing Prs
    Data = Data & "" & "|"
    Data = Data & "*******" & "|"
    Data = Data & "*******" & "|"
    Data = Data & "*******" & "|"
    Data = Data & "*******" & ";"
    Data = Data & "" & "|"
    Data = Data & "Total" & "|"
    Data = Data & Format(Tpayments, "0.00") & "|"
    Data = Data & Format(TGrat, "0.00") & "|"
    Data = Data & "" & "|"
    Data = Data & "" & ";"
    Tpayments = 0: TGrat = 0
    Data = Data & "" & "|"
    'Data = Data & "****************" & "|"
    Data = Data & "*******" & "|"
    Data = Data & "*******" & "|"
    Data = Data & "*******" & "|"
    Data = Data & "*******" & ";"
    SRd.MoveNext
 Loop
 DbClose SL, SRd
 ReportView.View.TableBorder = tbTopBottom
 ReportView.View.addTable F, H, Data, vbBlack, vbWhite
 FinishJob 1
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Server Report")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub ServerPayments(ServerName As String, BegCheck As Long)
On Error GoTo ErrorHandler
Printer_Drivers "Receipt"
Dim TotalChecks  As Currency
Dim TotalPayment As Currency
Dim TotalTip     As Currency
Dim PrinterPort  As String
Dim Chk          As Long
Dim TipPerc      As Single
Dim PDB          As ADODB.Connection
Dim Prs          As ADODB.Recordset
Dim TCash        As Currency
Dim Data         As String
Dim GratCharge   As Currency
TCash = 0
ConnectADODB FinancialDb, PDB
 Sql = "Select * from [dayPayments] Where [Server Name] = '" & ServerName & "' and [Check Number] > " & BegCheck - 1
 Sql = Sql & " Order by [Check Number] asc"
ConnectADORS Sql, PDB, Prs
  If Check_Option(96) = 0 Or AllPrinter.Item("Receipt").PrinterName <> "None" Then
    PF = FreeFile
    PrinterPort = "Server Payments"
    ReDim Seat_ttl(4, Number_In_Party)
tryagain:
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    Print #PF, EC(0) 'Int printer
    Print #PF, EC(1); EC(3); CenterText; "Payment Report"; EC(2); EC(7)
    Print #PF, LeftText; String(40, 45); EC(7)
    Print #PF, "Server " & ServerName; EC(7)
    Print #PF, Format(Time, "HH:MM AmPm");
    Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Payment"; Tab(16); "Pmt"; Tab(23); "Tip"; Tab(30); "Tot"; Tab(36); "Tip %"; EC(7)
    Print #PF, String(40, 45); EC(7)
  Else
    SetPage ReportView.View, "Today Server Payment Report For " & ServerName
    Dim H$
    H = "Check|Payment Type|Amount|Tips|Tip %|Gratuity;"
    Dim F$
    F = "+<2000|+<2500|+>1500|+>1000|+>1000|+>1000;"
  End If
  Chk = 0: TotalChecks = 0: TotalPayment = 0: TotalTip = 0
  If Not Prs.BOF Then Prs.MoveFirst
  Data = ""
  Do While Not Prs.EOF
    Chk = Prs![Check Number]
    TotalChecks = TotalChecks + Prs![Check Total]
    GratCharge = GratuityCollected(ServerName, Chk)
    Do While Prs![Check Number] = Chk
        If TotalChecks > 0 And Val(Check_Null(Prs![Tip Amount])) > 0 Then
         TipPerc = Format((TotalChecks / Val(Prs![Tip Amount] & "")) * 100, "0.00")
        Else
         TipPerc = 0
        End If
        If Check_Option(96) = 1 And AllPrinter.Item("Receipt").PrinterName = "None" Then
          Data = Data & CStr(Chk) & "|"
          Data = Data & Prs![Payment Type] & "|"
          Data = Data & Format(Prs![Payment Amount], "0.00") & "|"
          Data = Data & Format(Prs![Tip Amount], "0.00") & "|"
          Data = Data & Format(TipPerc, "0.00") & "|"
          Data = Data & Format(GratCharge, "0.00") & ";"
        Else
          Print #PF, CStr(Chk) & " " & Prs![Payment Type]; Tab(14); Format(Prs![Payment Amount], "0.00"); Tab(23); Format(Prs![Tip Amount], "0.00"); Tab(30); Format(Prs![Payment Amount] + Prs![Tip Amount], "0.00"); Tab(37); Format(TipPerc, "0.00"); EC(7)
          If GratCharge > 0 Then
           Print #PF, CStr(Chk) & " Gratuity Charged "; Tab(23); Format(GratCharge, "0.00"); EC(7)
          End If
        End If
        If Prs![Payment Type] = "Cash" Then
          TCash = TCash + Prs![Payment Amount]
        End If
        TotalPayment = TotalPayment + Val(Check_Null(Prs![Payment Amount]))
        TotalTip = TotalTip + Val(Check_Null(Prs![Tip Amount]))
        Chk = Val(Check_Null(Prs![Check Number]))
        Prs.MoveNext
        If Prs.EOF Then Exit Do
     Loop
     If Prs.EOF Then Exit Do
   Loop
  DbClose PDB, Prs
   Dim xmls As MPxmlParse
   Set xmls = New MPxmlParse
   ConnectADODB DebitDb, PDB
      If Trim$(ServerName) = "" Then
        Sql = "Select * from [Debit] Where [Tender] = 'EMV' and [Close Out day] = 0 or [Tender] = 'QRPayment' and [Close out day] = 0"
      Else
       Sql = "Select * from [Debit] Where [Tender] = 'EMV' and [Close Out day] = 0 and [Server Name] = '" & ServerName & "' and TRY_CONVERT(int, [Check Number]) >= " & BegCheck
       Sql = Sql & " or [Tender] = 'QRPayment' and [Close Out day] = 0 and [Server Name] = '" & ServerName & "' and TRY_CONVERT(int, [Check Number]) >= " & BegCheck
      End If
    Sql = Sql & " Order by [Check Number],[Server Name] asc"
    ConnectADORS Sql, PDB, Prs
    If Not Prs.BOF Then Prs.MoveFirst
    Do While Not Prs.EOF
        Chk = Prs![Check Number]
        GratCharge = GratuityCollected(ServerName, Chk)
        TotalChecks = TotalChecks + Prs![Amount]
        If Prs![Amount] > 0 And Val(Prs![tip] & "") > 0 Then
        TipPerc = Format((Prs![tip] / Val(Prs![Amount] & "")) * 100, "0.00")
        Else
        TipPerc = 0
        End If
        Do While Val(Prs![Check Number]) = Chk
        xmls.xml = Prs.fields("XMLString").Value & ""
        If Check_Option(96) = 1 And AllPrinter.Item("Receipt").PrinterName = "None" Then
          Data = Data & CStr(Chk) & "|"
          Data = Data & Prs.fields("Account Name").Value & "" & "|"
          Data = Data & Format(Prs![Amount], "0.00") & "|"
          Data = Data & Format(Prs![tip], "0.00") & "|"
          Data = Data & Format(TipPerc, "0.00") & "|"
          Data = Data & Format(GratCharge, "0.00") & ";"
        Else
         Print #PF, CStr(Chk) & " " & Prs.fields("Account Name").Value & "-" & Right$(xmls.token("AcctNo"), 4); Tab(14); Format(Prs![Amount], "0.00"); Tab(23); Format(Prs![tip], "0.00"); Tab(30); Format(Prs![Amount] + Prs![tip], "0.00"); Tab(37); Format(TipPerc, "0.00"); EC(7)
         If GratCharge > 0 Then
          Print #PF, CStr(Chk) & " Gratuity Charged "; Tab(23); Format(GratCharge, "0.00"); EC(7)
         End If
        End If
        TotalPayment = TotalPayment + Prs![Amount]
        TotalTip = TotalTip + Prs![tip]
        Chk = Val(Prs![Check Number])
        Prs.MoveNext
        If Prs.EOF Then Exit Do
        Loop
        If Prs.EOF Then Exit Do
    Loop
    Set xmls = Nothing
DbClose PDB, Prs
    TotalChecks = TotalPayment  'problem with split payment doubles sales
    If Check_Option(96) = 1 And AllPrinter.Item("Receipt").PrinterName = "None" Then
        Data = Data & "|"
        Data = Data & "Total Payments " & "|"
        Data = Data & Format(TotalPayment, "0.00") & "|"
        Data = Data & Format(TotalTip, "0.00") & "|"
        If TotalTip > 0 And TotalPayment > 0 Then
          Data = Data & Format((TotalTip / TotalPayment) * 100, "0.00") & ";"
        Else
          Data = Data & "0.00;"
        End If
        If TCash > 0 Then
        Data = Data & "|"
        Data = Data & "Total Cash" & "|"
        Data = Data & Format(TCash, "0.00") & "|"
        Data = Data & ";"
        Data = Data & "|"
        Data = Data & "Total Non Cash " & "|"
        Data = Data & Format(TotalPayment - TCash, "0.00") & "|"
        Data = Data & ";"
        End If
        ReportView.View.TableBorder = tbColTopBottom
        ReportView.View.addTable F, H, Data, vbBlack, vbWhite
        FinishJob 1
    Else
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Total Checks   " & Format(TotalChecks, "0.00")
        If TCash > 0 Then
        Print #PF, "Total Cash     " & Format(TCash, "0.00")
        Print #PF, "Total Non Cash " & Format(TotalPayment - TCash, "0.00")
        End If
        Print #PF, "Total Payments " & Format(TotalPayment, "0.00")
        Print #PF, "Total Tips $   " & Format(TotalTip, "0.00")
        If TotalTip > 0 And TotalPayment > 0 Then
        Print #PF, "Total Tip %    " & Format((TotalTip / TotalPayment) * 100, "0.00")
    End If
    If GetConfigInfo("Functions", "Cut Payment Report", "No", "FUNCTIONS.INI") = "ENABLED" Then
       Print #PF, EC(6) & EC(5)
    End If
    Close
    SendJobToPrinter PrinterPort, "Receipt", False
End If
Exit Sub
ErrorHandler:
If Err.Number = 9 Then Resume Next
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Payments")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub RunReport_Click()
If Not IsDate(STime.Text) Or Not IsDate(ETime.Text) Then
  See_Msg "Select Time Range", 5, ""
  Exit Sub
End If
Dim dw As Integer
If DOW(0).Text = "Today" Then
  dw = 0
Else
  dw = DOW(1).List(DOW(0).ListIndex)
End If
TAmount.Text = "0.00"
TAmount.Text = Format(GetAllPoolTips(STime.Text, ETime.Text, dw), "0.00")
If Check_Option(96) = 1 Then
    SetPage ReportView.View, "Non Cash Tip Total"
    PrintLine "Day " & DOW(0).Text, -1
    PrintLine "Starting Time", 2500: PrintLine "Ending Time", 5000: PrintLine "Total", -1
    DrawLine
    PrintLine STime.Text, 2500: PrintLine ETime.Text, 5000: PrintLine TAmount.Text, -1
    FinishJob 1
Else
    Dim PF As Integer
    Dim PrinterPort As String
    PF = FreeFile
    PrinterPort = "Server Payments"
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    Print #PF, EC(0)
    Print #PF, EC(1); EC(3); CenterText; "Tip Pool Report"; EC(2); EC(7)
    Print #PF, LeftText; String(40, 45); EC(7)
    Print #PF, Format(Time, "HH:MM AmPm");
    Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Day of Week " & DOW(0).Text & EC(7)
    Print #PF, "Start Time"; Tab(15); "End Time"; Tab(30); "Tip Total"; EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, STime.Text; Tab(15); ETime.Text; Tab(33); Format(TAmount.Text, "$0.00"); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6) & EC(5)
    Close
    SendJobToPrinter PrinterPort, "Receipt", False
End If
End Sub
Private Sub SList_Click()
End Sub
Public Sub MixReports(ServerName As String, ReportType As String, IndReport As String)
On Error GoTo ErrorHandler
Dim SL          As ADODB.Connection
Dim SRd         As ADODB.Recordset
Dim Data        As String
Dim MyServer    As String
Dim TSales      As Currency
Dim TTSales     As Currency
Dim QSales      As Long
Dim TQty        As Long
Dim TSvr        As String
Dim MyPerc      As Double
TTSales = GetNetSales(SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate)
ConnectADODB ProductHistory, SL

        If ReportType = "Sales Mix" Then
           SetPage ReportView.View, "Mix Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
           Sql = "Select Distinct [Menu Item],[server Name],sum([Quantity]) as QSold, Sum([Item Price]) as TSold  from [Detail] "
           Sql = Sql & " Where [Department Name] <> 'Gratuity' and TRY_CAST([Date] AS datetime) Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
           Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
           If ServerName > "" Then Sql = Sql & " and [Server Name] = '" & ServerName & "'"
           Sql = Sql & " Group By  [Menu Item],[server Name]"
        ElseIf ReportType = "Category" Then
           SetPage ReportView.View, "Category Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
           Sql = "Select Distinct [Category Name],[server Name],sum([Quantity]) as QSold, Sum([Item Price]) as TSold from [Detail] "
           Sql = Sql & " Where [Department Name] <> 'Gratuity' and TRY_CAST([Date] AS datetime) Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
           Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
           If ServerName > "" Then Sql = Sql & " and [Server Name] = '" & ServerName & "'"
           Sql = Sql & " Group By  [Category Name],[server Name]"
        Else
           SetPage ReportView.View, "Department Report " & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate
           Sql = "Select Distinct [Department Name],[server Name],sum([Quantity]) as QSold, Sum([Item Price]) as TSold  from [Detail] "
           Sql = Sql & " Where [Department Name] <> 'Gratuity' and TRY_CAST([Date] AS datetime) Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
           Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
           If ServerName > "" Then Sql = Sql & " and [Server Name] = '" & ServerName & "'"
           Sql = Sql & " Group By [Department Name],[server Name]"
        End If
If Sql = "" Then Exit Sub   'if no match of crieteria then exit sub
Dim H$
H = "Server|Item Name|Quantity|Amount|Percentage;"
Dim F$
F = "+<2500|+<2500|+>1000|+>1500|+>1200;"
Data = ""
TSales = 0:  QSales = 0: TQty = 0
ConnectADORS Sql, SL, SRd
  Do While Not SRd.EOF
        MyPerc = 0
        If Val(SRd!TSold & "") <> 0 And TTSales <> 0 Then
          MyPerc = (SRd!TSold / TTSales) * 100
        End If
        If MyServer <> SRd![Server Name] Then
         MyServer = SRd![Server Name]
         Data = Data & SRd![Server Name] & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & "|"
         Data = Data & "" & ";"
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
          TQty = TQty + SRd!QSold
        SRd.MoveNext
        If Not SRd.EOF Then
         If MyServer <> SRd![Server Name] Then
            Data = Data & "" & "|"
            Data = Data & "** Total **" & "|"
            Data = Data & Format(QSales, "###") & "|"
            Data = Data & Format(TSales, "0.00") & "|"
           If TSales <> 0 And TTSales <> 0 Then
                MyPerc = (TSales / TTSales) * 100
                Data = Data & Format(MyPerc, "#.00") & ";"
            Else
                Data = Data & "" & ";"
            End If
            TSales = 0:  QSales = 0: TQty = 0
         End If
        End If
     Loop
DbClose SL, SRd
Data = Data & "" & "|"
Data = Data & "** Total **" & "|"
Data = Data & Format(QSales, "###") & "|"
Data = Data & Format(TSales, "0.00") & "|"
If TSales <> 0 And TTSales <> 0 Then
  MyPerc = (TSales / TTSales) * 100
  Data = Data & Format(MyPerc, "#.00") & ";"
Else
  Data = Data & "" & ";"
End If
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
Private Sub SSCommand1_Click()
LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
    If SearchDates Is Nothing Then
     Exit Sub
    End If
    ReportType.Caption = "Payments"
    BtnEnh6.Visible = True
    BtnEnh7.Visible = True
    BtnEnh9.Visible = False
    LoadPaymentList "Payments"
    PaymentType = "Payments"
End Sub
Private Sub SSCommand13_Click()
Load_What = "AllServers"
Find_What = "Select a Server"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
DeptVoid GetSelected, True, True
End Sub
Private Sub SSCommand14_Click()
If Can_Continue(SSCommand14.Caption) = False Then Exit Sub
Dim CV9 As VER11SystemMaintenance.MaintenanceMenus
Set CV9 = New VER11SystemMaintenance.MaintenanceMenus
 If GetConfigInfo("Misc", "Skims Are On", "No", "Functions.INI") = "ENABLED" Then
   CV9.LoadSkims LoggedOnAs(0).EmpLoggedOnAs, LoggedOnAs(0).EmpAccessLevel, OpenDrawerNo
 Else
  Dim Empdata(1, 0) As String
  Empdata(0, 0) = LoggedOnAs(0).EmpLoggedOnAs
  Empdata(1, 0) = LoggedOnAs(0).EmpAccessLevel
  'CV9.SystemSetup Empdata(), ReturnSystemType, "AuditDeposit"
 End If
Set CV9 = Nothing
End Sub
Public Function SvrDeposit(SvrName As String) As Currency
Dim GA As Integer, ff As Integer
ff = FreeFile
SvrDeposit = 0: GA = 0
Open SvrPath.TerminalPath & "DAT Files\" & "CashDeposits.Dat" For Random Shared As ff Len = Len(Sdeposit)
     Do
       GA = GA + 1
       Get #ff, GA, Sdeposit
           If Trim(SvrName) = Trim(Sdeposit.SvrName) Then
              SvrDeposit = SvrDeposit + Sdeposit.DepositAmt
           End If
     Loop Until GA >= LOF(ff) / Len(Sdeposit)
Close ff
End Function
Private Sub SSCommand15_Click()
RemoveNullDates
LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
GetCreditCardReport
End Sub
Public Sub LoadPaymentList(WhichList As String)
On Error GoTo ErrorHandler
Dim SL  As ADODB.Connection
Dim SRd As ADODB.Recordset
If WhichList = "Sales Mix" Or WhichList = "Department" Or WhichList = "Category" Then
  ConnectADODB ProductHistory, SL
            Sql = "Select Distinct [Server Name] from [Detail] "
            Sql = Sql & " Where TRY_CAST([Date] AS datetime) Between #" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "#"
            Sql = Sql & " and #" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "#"
            ConnectADORS ReturnDateString(Sql), SL, SRd
            SList.Clear
                If Not SRd.EOF Then
                    CenterObject Me, ServerList
                    ServerList.Visible = True
                    Sleep 0
                Else
                    DbClose SL, SRd
                     See_Msg "No Names Found", 5, ""
                    Exit Sub
                End If
                   Do While Not SRd.EOF
                     SList.AddItem SRd![Server Name]
                     SRd.MoveNext
                   Loop
                DbClose SL, SRd
                Exit Sub
End If
RemoveNullDates
  ConnectADODB "MSaleData", SL
  If WhichList = "Payments" Then
      Sql = "Select Distinct [Server Name] from [Server Payments] "
      Sql = Sql & " Where TRY_CAST([Date] AS date) Between '" & CDate(SearchDates.Item(1).StartDate) & "'"
      Sql = Sql & " and '" & CDate(SearchDates.Item(1).EndDate) & "'"
  Else
      Sql = "Select Distinct [Server Name] from [Tip Income] "
      Sql = Sql & " Where TRY_CAST([Date] AS date) Between '" & CDate(SearchDates.Item(1).StartDate & " " & GetStartTime) & "'"
      Sql = Sql & " and '" & CDate(SetEndingDate(SearchDates.Item(1).EndDate) & " " & SetEndingTime) & "'"
  End If
ConnectADORS (Sql), SL, SRd
SList.Clear
     If SRd.EOF Then
       DbClose SL, SRd
       See_Msg "No Names Found", 5, ""
       Exit Sub
     End If
     CenterObject Me, ServerList
     ServerList.Visible = True
     Sleep 0
     Do While Not SRd.EOF
        SList.AddItem SRd![Server Name] & ""
        SRd.MoveNext
     Loop
  DbClose SL, SRd
  Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Payment List")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand2_Click()
If See_Msg("Are You Sure ?", 4, "Closing All Audit") <> 6 Then Exit Sub
If Can_Continue(SSCommand2.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Heading = "Server Report"
Dim sDb             As ADODB.Connection
Dim sTb             As ADODB.Recordset
Dim ChKN0           As Long
Dim RegDrawer       As String
Dim DrawerSales     As Currency
Dim Caudits         As Boolean
Dim ReportType      As String
Dim S_Name          As String
Dim CDrawer         As Long
Dim StartBank       As Currency
Dim StartTime       As String
Dim StartDate       As String
Dim StartBal        As Currency
Dim DepositPaidIn   As Currency
If See_Msg("Print Audit Slip", 4, "Closing Audits") = 6 Then
    Caudits = False
Else
    Caudits = True
End If
ConnectADODB "MSaleData", sDb
Sql = "select * from [Check Numbers] where [audit status] = 0"
ConnectADORS Sql, sDb, sTb
Do While Not sTb.EOF
   SSCommand2.Caption = "Searching.... " & sTb![Servers Name]
   SSCommand2.Refresh
   StartTime = TimeValue(sTb.fields("Time").Value)
   StartDate = DateValue(sTb.fields("Start Date").Value)
   ChKN0 = sTb.fields("Check No").Value
   RegDrawer = sTb![Register Name] & ""
   CDrawer = Val(sTb![Drawer No] & "")
   StartBank = Val(sTb![Starting Bank] & "")
   S_Name = sTb![Servers Name]
   StartBal = Val(sTb![Starting Balance] & "")
   DepositPaidIn = Val(sTb![Deposits Paid In] & "")
   If ChKN0 = 0 Then ChKN0 = 1
ReportType = "Shift"
DontUnLoad = True
GetServerAudit S_Name, ChKN0, StartTime, StartDate, RegDrawer, Caudits, ReportType, CDrawer, True, StartBank, 0, StartBal, DepositPaidIn
NextServer:
sTb.MoveNext
Loop
DbClose sDb, sTb
DontUnLoad = False
Call MemoryCheck(Server_Reports)
SSCommand2.Caption = "Close All Audits"
Exit Sub
Printerr:
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Server Reports SS2")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand23_Click()
End Sub
Private Sub SSCommand25_Click()
If Can_Continue(SSCommand25.Caption) = False Then Exit Sub
Load_What = "Closed Servers"
Find_What = "Select a Server"
Dim StartingNo  As Long
Dim server_t    As String
Dim PrinterPort As String
Dim AuditDate   As String
Dim AuditTime   As String
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
server_t = GetSelected  'getselected is the server
GetSelected = ""
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB "MSaleData", Db
 Sql = "Select * From [Check Numbers] where [Servers Name] = '" & server_t & "'"
 Sql = Sql & " Order by TRY_CAST([start date] AS datetime) + TRY_CAST([Time] AS datetime) desc"
ConnectADORS Sql, Db, Tb
If Not Tb.EOF Then
  If Tb.fields("Audit Status").Value = 0 Then
    See_Msg "Cannot ReOpen Audit", 5, ""
    DbClose Db, Tb
    Exit Sub
  End If
  Tb![Audit Status] = 0
  AuditDate = Tb![Start date] & ""
  AuditTime = Tb![Time] & ""
  StartingNo = Val(Tb![Check No])
  Tb.Update
  See_Msg "Audit ReOpened", 5, server_t
  Trans_Log "Reopen Server Audit " & server_t
  ReOpenSvrAudit server_t, AuditDate, AuditTime
  ReOpenPaidOut server_t
End If
DbClose Db, Tb
'If StartingNo = 0 Then Exit Sub
End Sub
Private Sub SSCommand3_Click()
If Can_Continue(SSCommand3.Caption) = False Then Exit Sub
Load_What = "AllServers"
Find_What = "Select a Server"
RequestListForm.Show 1
If GetSelected = "" Then
   Exit Sub
End If
ServerPayments GetSelected, 0
End Sub
Private Sub SSCommand4_Click()
If See_Msg("Current Day ?", 4, "Department Sale for Today") = 6 Then
 CurrentDepartmentSales
Else
 LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
 If SearchDates Is Nothing Then
  Exit Sub
 End If
 ReportType.Caption = "Department Sales"
 BtnEnh6.Visible = False
 BtnEnh7.Visible = False
 BtnEnh9.Visible = True
 LoadPaymentList "Department"
 PaymentType = "Department"
End If
End Sub
Private Sub SSCommand5_Click()
If Can_Continue(SSCommand5.Caption) = False Then Exit Sub
ShowAllAudits
End Sub
Private Sub SSCommand7_Click()
LoadDates "Select [Date] from [Sales] order by [Date] desc", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
ReportType.Caption = "Category Sales"
BtnEnh6.Visible = False
BtnEnh7.Visible = False
BtnEnh9.Visible = True
LoadPaymentList "Category"
PaymentType = "Category"
End Sub
Private Sub SSCommand8_Click()
If Can_Continue(SSCommand8.Caption) = False Then Exit Sub
Server_Reports.Show
End Sub
Private Sub STime_Click()
Dim InTime  As String
Dim PKT     As DateOrTime
Set PKT = New DateOrTime
InTime = PKT.ReturnTime
If IsDate(InTime) Then
  STime.Text = Format(InTime, "H:mm ampm")
Else
  STime.Text = ""
End If
Set PKT = Nothing
End Sub

Private Sub TipPooling_Click()

End Sub


