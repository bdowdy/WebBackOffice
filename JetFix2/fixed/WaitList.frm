VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form WaitList 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Customer Wait List"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15150
   Icon            =   "WaitList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   675
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1191
      _StockProps     =   66
      Caption         =   "List By # In Party"
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
      textCaption     =   "WaitList.frx":000C
      textLT          =   "WaitList.frx":0090
      textCT          =   "WaitList.frx":00A8
      textRT          =   "WaitList.frx":00C0
      textLM          =   "WaitList.frx":00D8
      textRM          =   "WaitList.frx":00F0
      textLB          =   "WaitList.frx":0108
      textCB          =   "WaitList.frx":0120
      textRB          =   "WaitList.frx":0138
      colorBack       =   "WaitList.frx":0150
      colorIntern     =   "WaitList.frx":017A
      colorMO         =   "WaitList.frx":01A4
      colorFocus      =   "WaitList.frx":01CE
      colorDisabled   =   "WaitList.frx":01F8
      colorPressed    =   "WaitList.frx":0222
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
      Left            =   9240
      TabIndex        =   11
      Top             =   9840
      Width           =   3255
      _Version        =   589828
      _ExtentX        =   5741
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   $"WaitList.frx":024C
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
      textCaption     =   "WaitList.frx":025D
      textLT          =   "WaitList.frx":02D7
      textCT          =   "WaitList.frx":02EF
      textRT          =   "WaitList.frx":0307
      textLM          =   "WaitList.frx":031F
      textRM          =   "WaitList.frx":0337
      textLB          =   "WaitList.frx":034F
      textCB          =   "WaitList.frx":0367
      textRB          =   "WaitList.frx":037F
      colorBack       =   "WaitList.frx":0397
      colorIntern     =   "WaitList.frx":03C1
      colorMO         =   "WaitList.frx":03EB
      colorFocus      =   "WaitList.frx":0415
      colorDisabled   =   "WaitList.frx":043F
      colorPressed    =   "WaitList.frx":0469
      HollowFrame     =   -1  'True
      LightDirection  =   1
      VistaColorGlossyEffectLower=   13677994
      Office2007ColorOuterBorder=   16777215
      Office2007ColorNorthTop=   128
      Office2007ColorNorthBottom=   128
      Office2007ColorSouthTop=   128
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
      Office2003ColorSouth=   255
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   8421504
      Office2003ColorSouthMouseOver=   0
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   5395026
      Office2003ColorSouthPressed=   2302755
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand8 
      Height          =   855
      Left            =   12840
      TabIndex        =   9
      Top             =   9840
      Visible         =   0   'False
      Width           =   2175
      _Version        =   589828
      _ExtentX        =   3836
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
      PictureYOffset  =   -2
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPicture      =   "WaitList.frx":0493
      textCaption     =   "WaitList.frx":0530
      textLT          =   "WaitList.frx":0548
      textCT          =   "WaitList.frx":0560
      textRT          =   "WaitList.frx":0578
      textLM          =   "WaitList.frx":0590
      textRM          =   "WaitList.frx":05A8
      textLB          =   "WaitList.frx":05C0
      textCB          =   "WaitList.frx":05D8
      textRB          =   "WaitList.frx":05F0
      colorBack       =   "WaitList.frx":0608
      colorIntern     =   "WaitList.frx":0632
      colorMO         =   "WaitList.frx":065C
      colorFocus      =   "WaitList.frx":0686
      colorDisabled   =   "WaitList.frx":06B0
      colorPressed    =   "WaitList.frx":06DA
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
      Office2003ColorSouth=   15461355
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   15790320
      Office2003ColorSouthPressed=   15461355
      Office2003ColorBorderPressed=   12632256
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand7 
      Height          =   735
      Left            =   12600
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      _Version        =   589828
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      BackColorContainer=   4210752
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPicture      =   "WaitList.frx":0704
      textCaption     =   "WaitList.frx":07AD
      textLT          =   "WaitList.frx":07C5
      textCT          =   "WaitList.frx":07DD
      textRT          =   "WaitList.frx":07F5
      textLM          =   "WaitList.frx":080D
      textRM          =   "WaitList.frx":0825
      textLB          =   "WaitList.frx":083D
      textCB          =   "WaitList.frx":0855
      textRB          =   "WaitList.frx":086D
      colorBack       =   "WaitList.frx":0885
      colorIntern     =   "WaitList.frx":08AF
      colorMO         =   "WaitList.frx":08D9
      colorFocus      =   "WaitList.frx":0903
      colorDisabled   =   "WaitList.frx":092D
      colorPressed    =   "WaitList.frx":0957
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
      Office2003ColorSouth=   15461355
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   15790320
      Office2003ColorSouthMouseOver=   15461355
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   15790320
      Office2003ColorSouthPressed=   15461355
      Office2003ColorBorderPressed=   12632256
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand3 
      Height          =   855
      Left            =   6480
      TabIndex        =   5
      Top             =   9840
      Width           =   2655
      _Version        =   589828
      _ExtentX        =   4683
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Add Name"
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
      textCaption     =   "WaitList.frx":0981
      textLT          =   "WaitList.frx":09F1
      textCT          =   "WaitList.frx":0A09
      textRT          =   "WaitList.frx":0A21
      textLM          =   "WaitList.frx":0A39
      textRM          =   "WaitList.frx":0A51
      textLB          =   "WaitList.frx":0A69
      textCB          =   "WaitList.frx":0A81
      textRB          =   "WaitList.frx":0A99
      colorBack       =   "WaitList.frx":0AB1
      colorIntern     =   "WaitList.frx":0ADB
      colorMO         =   "WaitList.frx":0B05
      colorFocus      =   "WaitList.frx":0B2F
      colorDisabled   =   "WaitList.frx":0B59
      colorPressed    =   "WaitList.frx":0B83
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
   Begin BTNENHLib4.BtnEnh SSCommand1 
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   9840
      Width           =   1095
      _Version        =   589828
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Show All"
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
      textCaption     =   "WaitList.frx":0BAD
      textLT          =   "WaitList.frx":0C1D
      textCT          =   "WaitList.frx":0C35
      textRT          =   "WaitList.frx":0C4D
      textLM          =   "WaitList.frx":0C65
      textRM          =   "WaitList.frx":0C7D
      textLB          =   "WaitList.frx":0C95
      textCB          =   "WaitList.frx":0CAD
      textRB          =   "WaitList.frx":0CC5
      colorBack       =   "WaitList.frx":0CDD
      colorIntern     =   "WaitList.frx":0D07
      colorMO         =   "WaitList.frx":0D31
      colorFocus      =   "WaitList.frx":0D5B
      colorDisabled   =   "WaitList.frx":0D85
      colorPressed    =   "WaitList.frx":0DAF
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
      Office2003ColorNorth=   12615680
      Office2003ColorSouth=   12615680
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   12615680
      Office2003ColorSouthMouseOver=   12615680
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   5395026
      Office2003ColorSouthPressed=   2302755
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   240
      Top             =   8280
   End
   Begin VSFlex8DAOCtl.VSFlexGrid CList 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   14895
      _cx             =   1992451745
      _cy             =   1992439892
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8454143
      ForeColorSel    =   0
      BackColorBkg    =   14737632
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
      GridLines       =   5
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
   Begin BTNENHLib4.BtnEnh SSCommand2 
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   9840
      Width           =   1095
      _Version        =   589828
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Show Waiting"
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
      CornerFactor    =   20
      Surface         =   15
      BackColorContainer=   12632256
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "WaitList.frx":0DD9
      textLT          =   "WaitList.frx":0E51
      textCT          =   "WaitList.frx":0E69
      textRT          =   "WaitList.frx":0E81
      textLM          =   "WaitList.frx":0E99
      textRM          =   "WaitList.frx":0EB1
      textLB          =   "WaitList.frx":0EC9
      textCB          =   "WaitList.frx":0EE1
      textRB          =   "WaitList.frx":0EF9
      colorBack       =   "WaitList.frx":0F11
      colorIntern     =   "WaitList.frx":0F3B
      colorMO         =   "WaitList.frx":0F65
      colorFocus      =   "WaitList.frx":0F8F
      colorDisabled   =   "WaitList.frx":0FB9
      colorPressed    =   "WaitList.frx":0FE3
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
      Office2003ColorNorth=   12615680
      Office2003ColorSouth=   12615680
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   12615680
      Office2003ColorSouthMouseOver=   12615680
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
      Left            =   2640
      TabIndex        =   6
      Top             =   9840
      Width           =   1095
      _Version        =   589828
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Find Name"
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
      textCaption     =   "WaitList.frx":100D
      textLT          =   "WaitList.frx":107F
      textCT          =   "WaitList.frx":1097
      textRT          =   "WaitList.frx":10AF
      textLM          =   "WaitList.frx":10C7
      textRM          =   "WaitList.frx":10DF
      textLB          =   "WaitList.frx":10F7
      textCB          =   "WaitList.frx":110F
      textRB          =   "WaitList.frx":1127
      colorBack       =   "WaitList.frx":113F
      colorIntern     =   "WaitList.frx":1169
      colorMO         =   "WaitList.frx":1193
      colorFocus      =   "WaitList.frx":11BD
      colorDisabled   =   "WaitList.frx":11E7
      colorPressed    =   "WaitList.frx":1211
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
      Office2003ColorNorth=   12615680
      Office2003ColorSouth=   12615680
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   12615680
      Office2003ColorSouthMouseOver=   12615680
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
      Left            =   3840
      TabIndex        =   7
      Top             =   9840
      Width           =   1095
      _Version        =   589828
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Show All Lists"
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
      CornerFactor    =   20
      Surface         =   15
      BackColorContainer=   12632256
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "WaitList.frx":123B
      textLT          =   "WaitList.frx":12B7
      textCT          =   "WaitList.frx":12CF
      textRT          =   "WaitList.frx":12E7
      textLM          =   "WaitList.frx":12FF
      textRM          =   "WaitList.frx":1317
      textLB          =   "WaitList.frx":132F
      textCB          =   "WaitList.frx":1347
      textRB          =   "WaitList.frx":135F
      colorBack       =   "WaitList.frx":1377
      colorIntern     =   "WaitList.frx":13A1
      colorMO         =   "WaitList.frx":13CB
      colorFocus      =   "WaitList.frx":13F5
      colorDisabled   =   "WaitList.frx":141F
      colorPressed    =   "WaitList.frx":1449
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
      Office2003ColorNorth=   12615680
      Office2003ColorSouth=   12615680
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   12615680
      Office2003ColorSouthMouseOver=   12615680
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   5395026
      Office2003ColorSouthPressed=   2302755
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh SSCommand9 
      Height          =   855
      Left            =   5040
      TabIndex        =   8
      Top             =   9840
      Width           =   1095
      _Version        =   589828
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Clear List"
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
      textCaption     =   "WaitList.frx":1473
      textLT          =   "WaitList.frx":14E7
      textCT          =   "WaitList.frx":14FF
      textRT          =   "WaitList.frx":1517
      textLM          =   "WaitList.frx":152F
      textRM          =   "WaitList.frx":1547
      textLB          =   "WaitList.frx":155F
      textCB          =   "WaitList.frx":1577
      textRB          =   "WaitList.frx":158F
      colorBack       =   "WaitList.frx":15A7
      colorIntern     =   "WaitList.frx":15D1
      colorMO         =   "WaitList.frx":15FB
      colorFocus      =   "WaitList.frx":1625
      colorDisabled   =   "WaitList.frx":164F
      colorPressed    =   "WaitList.frx":1679
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
      Office2003ColorNorth=   12615680
      Office2003ColorSouth=   12615680
      Office2003ColorBorder=   12632256
      Office2003ColorNorthMouseOver=   12615680
      Office2003ColorSouthMouseOver=   12615680
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   5395026
      Office2003ColorSouthPressed=   2302755
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh2 
      Height          =   675
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1191
      _StockProps     =   66
      Caption         =   "List By Time In"
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
      textCaption     =   "WaitList.frx":16A3
      textLT          =   "WaitList.frx":1721
      textCT          =   "WaitList.frx":1739
      textRT          =   "WaitList.frx":1751
      textLM          =   "WaitList.frx":1769
      textRM          =   "WaitList.frx":1781
      textLB          =   "WaitList.frx":1799
      textCB          =   "WaitList.frx":17B1
      textRB          =   "WaitList.frx":17C9
      colorBack       =   "WaitList.frx":17E1
      colorIntern     =   "WaitList.frx":180B
      colorMO         =   "WaitList.frx":1835
      colorFocus      =   "WaitList.frx":185F
      colorDisabled   =   "WaitList.frx":1889
      colorPressed    =   "WaitList.frx":18B3
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
      Height          =   675
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1191
      _StockProps     =   66
      Caption         =   "List By Preference"
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
      textCaption     =   "WaitList.frx":18DD
      textLT          =   "WaitList.frx":1961
      textCT          =   "WaitList.frx":1979
      textRT          =   "WaitList.frx":1991
      textLM          =   "WaitList.frx":19A9
      textRM          =   "WaitList.frx":19C1
      textLB          =   "WaitList.frx":19D9
      textCB          =   "WaitList.frx":19F1
      textRB          =   "WaitList.frx":1A09
      colorBack       =   "WaitList.frx":1A21
      colorIntern     =   "WaitList.frx":1A4B
      colorMO         =   "WaitList.frx":1A75
      colorFocus      =   "WaitList.frx":1A9F
      colorDisabled   =   "WaitList.frx":1AC9
      colorPressed    =   "WaitList.frx":1AF3
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
      Height          =   675
      Left            =   8760
      TabIndex        =   16
      Top             =   120
      Width           =   2775
      _Version        =   589828
      _ExtentX        =   4895
      _ExtentY        =   1191
      _StockProps     =   66
      Caption         =   "Print Wait List"
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
      textCaption     =   "WaitList.frx":1B1D
      textLT          =   "WaitList.frx":1B9B
      textCT          =   "WaitList.frx":1BB3
      textRT          =   "WaitList.frx":1BCB
      textLM          =   "WaitList.frx":1BE3
      textRM          =   "WaitList.frx":1BFB
      textLB          =   "WaitList.frx":1C13
      textCB          =   "WaitList.frx":1C2B
      textRB          =   "WaitList.frx":1C43
      colorBack       =   "WaitList.frx":1C5B
      colorIntern     =   "WaitList.frx":1C85
      colorMO         =   "WaitList.frx":1CAF
      colorFocus      =   "WaitList.frx":1CD9
      colorDisabled   =   "WaitList.frx":1D03
      colorPressed    =   "WaitList.frx":1D2D
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
      Office2003ColorNorth=   4227200
      Office2003ColorSouth=   24158
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   4227200
      Office2003ColorSouthMouseOver=   24158
      Office2003ColorBorderMouseOver=   65408
      Office2003ColorNorthPressed=   4227200
      Office2003ColorSouthPressed=   24158
      Office2003ColorBorderPressed=   65408
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VB.Label Parties 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parties Waiting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   9240
      Width           =   4455
   End
   Begin VB.Label AverageWait 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Average Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   9240
      Width           =   4455
   End
   Begin VB.Label GuestWaiting 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guest Waiting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   9240
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Top             =   9720
      Width           =   6135
   End
End
Attribute VB_Name = "WaitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyGroup  As String
Private CMessage As String
Private TReady   As String
Private Sub BlinkNextTable()
On Error GoTo ERH
Dim LFOT As Integer
With TableLayout
For LFOT = 0 To .MyTable.count - 1
  If .MyTable(LFOT).ForeColor = vbGreen Then
   If .MyTable(LFOT).BackColor <> vbBlack Then
    .MyTable(LFOT).MarqueeStyle = ssBlinkingMarquee
   End If
  End If
Next LFOT
End With
Exit Sub
ERH:
Err.Clear
End Sub
Private Sub BtnEnh1_Click()
SaveList
Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime),[Number In Party]"
LoadList True, Sql

End Sub

Private Sub BtnEnh2_Click()
SaveList
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql

End Sub

Private Sub BtnEnh3_Click()
SaveList
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime),[Preference]"
LoadList True, Sql
End Sub

Private Sub BtnEnh4_Click()
SaveList
ReportView.View.Orientation = orPortrait
SetPage ReportView.View, Format(Date, "MM/DD/YYYY")
ReportView.View.RenderControl = CList.hwnd
ReportView.View.EndDoc
ReportView.Caption = "Customer Wait List"
ReportView.Show 1
Call MemoryCheck(ReportView)
End Sub
Private Sub CList_Click()
Dim PHONE_NO    As String
Dim TxtResponse As String
With CList
  If .Row < 0 Then Exit Sub
  If .Col = 0 Then
   If .Cell(flexcpChecked, .Row, 0) = flexUnchecked Then
      .Cell(flexcpChecked, .Row, 0) = flexChecked
      .TextMatrix(.Row, 8) = Format(Time, "H:mm ampm")
      SaveList
      If Check_Option(89) = 1 Then
        BlinkNextTable
        SSCommand4_Click
        Exit Sub
      End If
   Else
     .Cell(flexcpChecked, .Row, 0) = flexUnchecked
      .TextMatrix(.Row, 8) = ""
   End If
  End If
If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = Format(Time, "H:MM ampm")
Select Case .Col
   Case Is = 0
   Case Is = 1
      .TextMatrix(.Row, .Col) = GetCovers
   Case Is = 2
      If .TextMatrix(.Row, .Col) > "" Then
      
      Else
       .TextMatrix(.Row, .Col) = ShowKeyBoard("Enter Customer Name")  'Enter_Name(True) '
      End If
   Case Is = 3
     If .TextMatrix(.Row, .Col) > "" Then
   
     Else
        Dim rTime As String
        Dim PKT As DateOrTime
        Set PKT = New DateOrTime
        rTime = PKT.ReturnTime
        If rTime > "" Then
          .TextMatrix(.Row, .Col) = Format(rTime, "H:MM ampm")
        End If
        Set PKT = Nothing
     End If
   Case Is = 5
      Load_What = "Preferences"
      Find_What = "Select Preference"
      Select_What.Show 1
      If GetSelected = "" Then GetSelected = ShowKeyBoard("Enter Preference")
      .TextMatrix(.Row, .Col) = GetSelected
   Case Is = 6
   
   
   Case Is = 9
     If .TextMatrix(.Row, .Col) > "" Then
       If See_Msg("Edit Cell Number", 4, "") = 6 Then
        .TextMatrix(.Row, .Col) = ""
        GoTo RedoPhone
       End If
     Else
RedoPhone:
       .TextMatrix(.Row, .Col) = ShowNumberPad("Cell", False, 0, 0)
       SaveList
       SendTextMessage "Text From " & StoreName & vbCrLf & " Estimated Wait Is " & .TextMatrix(.Row, 6) & " Mins" & vbCrLf & CMessage, .TextMatrix(.Row, 9)
       LoadAllToList
     End If
   Case Is = 10
      If .TextMatrix(.Row, .Col) > "" Then
      
      Else
       If Check_Option(153) = 0 Then
       .TextMatrix(.Row, .Col) = "ALL"
       Else
        Load_What = "Areas"
        Find_What = "Select Area"
        Select_What.Show 1
        If GetSelected = "" Then Exit Sub
        .TextMatrix(.Row, .Col) = GetSelected
       End If
      End If
  Case Is = 11
       If .TextMatrix(.Row, 9) > "" Then
       If SiteLic.Twilio = 0 Then
          See_Msg "No License To Text", 5, "Contact dealer for texting support"
          Exit Sub
       End If
       If See_Msg("Send Message To " & .TextMatrix(.Row, 2) & " ?", 4, "Table is Ready") = 6 Then
         If SendTextMessage("Text From " & StoreName & vbCrLf & " Your Table Is Ready" & vbCrLf & TReady, .TextMatrix(.Row, 9)) = "Success" Then
           .TextMatrix(.Row, 11) = Format(Time, "HH:MM ampm")
           SaveList
           Exit Sub
         End If
       End If
     End If
End Select
End With
End Sub

Private Function GetNextRow()
Dim TR As Integer
For TR = 1 To CList.rows - 1
  If Val(CList.TextMatrix(TR, 7)) = 0 Then
       If CList.TextMatrix(TR, 2) = "" And CList.TextMatrix(TR, 1) = "" And CList.TextMatrix(TR, 5) = "" And CList.TextMatrix(TR, 6) = "" Then
         GetNextRow = TR
         Exit Function
       End If
  End If
Next TR
End Function
Private Sub Form_Load()
ResizeScreen1024 Me
CenterForm Me
If Check_Option(153) = 1 Then
 MyGroup = SvrPath.TerminalGroupName
Else
 MyGroup = "ALL"
 SSCommand6.Enabled = False
End If
CMessage = ReturnTxtFromFile(SvrPath.TerminalPath & "DAT Files\" & "CMessage.dat")
'// CMessage is the confirmation text message
TReady = ReturnTxtFromFile(SvrPath.TerminalPath & "DAT Files\" & "TReady.dat")
'// TReady is the message your table is ready
Dim dw As Integer
With CList
     .cols = 12
     .rows = 2
     .ColWidth(0) = 1000 * WidScale1024  'Seated
     .ColWidth(1) = 1200 * WidScale1024    'Number in Party
     .ColWidth(2) = 2800 * WidScale1024  'Customer Name
     .ColWidth(3) = 1200 * WidScale1024  'Time In
     .ColWidth(4) = 1100 * WidScale1024 'Elapsed Time
     .ColWidth(5) = 2000 * WidScale1024  'Preference
     .ColWidth(6) = 1200 * WidScale1024  'Quoted Time
     .ColWidth(7) = 0      'stored
     .ColWidth(8) = 0
     .ColWidth(9) = 1600 * WidScale1024
     .ColWidth(10) = 1400 * WidScale1024
     .ColWidth(11) = 1600 * WidScale1024
     .FixedCols = 0
     .FixedRows = 1
     .TextMatrix(0, 0) = "Seated"
     .TextMatrix(0, 1) = "Party #"
     .TextMatrix(0, 2) = "Customer Name"
     .TextMatrix(0, 3) = "Time In"
     .TextMatrix(0, 4) = "Waiting"
     .TextMatrix(0, 5) = "Preference"
     .TextMatrix(0, 6) = "Quoted"
     .TextMatrix(0, 9) = "Cell"
     .TextMatrix(0, 10) = "Area"
     .TextMatrix(0, 11) = "Send Text"
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(6) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignCenterCenter
     .Cell(flexcpChecked, 1, 0, 1, 0) = 2
     .ColAlignment(0) = flexAlignCenterCenter
     .width = 0
     For dw = 0 To 11
     .width = .width + .ColWidth(dw)
     Next dw
     .width = .width + 300
End With
LoadReservation Date
LoadAllToList
End Sub
Private Sub LoadAllToList()
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql

End Sub

Private Sub SSCommand1_Click()
SaveList
 Sql = "Select * from [Wait List] Where [Group] = '" & MyGroup & "' and [Day] = 0 order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList False, Sql
End Sub
Private Sub SSCommand2_Click()
SaveList
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql
End Sub
Private Sub SSCommand3_Click()
Timer1.Enabled = False
CList.rows = CList.rows + 1
CList.Cell(flexcpChecked, CList.rows - 1, 0) = flexUnchecked
With CList
      .Row = .rows - 1
      .RowHeight(.Row) = 500
      .TextMatrix(.Row, 1) = GetCovers
      If Val(.TextMatrix(.Row, 1)) = 0 Then GoTo ClearRow
      .TextMatrix(.Row, 2) = ShowKeyBoard("Enter Customer Name")  'Enter_Name(True)  '
      If .TextMatrix(.Row, 2) = "" Then GoTo ClearRow
      .TextMatrix(.Row, 3) = Format(Time, "H:MM ampm")
      Load_What = "Preferences"
      Find_What = "Select Preference"
      Select_What.Show 1
      If GetSelected = "" Then GetSelected = ShowKeyBoard("Enter Preference")
      .TextMatrix(.Row, 5) = GetSelected
      .TextMatrix(.Row, 6) = ShowNumberPad("Enter Quote Mins", False, 0, 0)
      .TextMatrix(.Row, 9) = ShowNumberPad("Enter Cell Number", True, 0, 0)
      SendTextMessage "Text From " & StoreName & vbCrLf & " Estimated Wait Is " & .TextMatrix(.Row, 6) & " Mins" & vbCrLf & CMessage, .TextMatrix(.Row, 9)
      If Val(.TextMatrix(.Row, 6)) = 0 Then GoTo ClearRow
      If Check_Option(153) = 0 Then
       .TextMatrix(.Row, 10) = "ALL"
      Else
        Load_What = "Areas"
        Find_What = "Select Area"
        Select_What.Show 1
        If GetSelected = "" Then Exit Sub
        .TextMatrix(.Row, 10) = GetSelected
      End If
End With
SaveList
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql
Timer1.Enabled = True
Exit Sub
ClearRow:
CList.RemoveItem CList.Row
End Sub
Private Sub SSCommand4_Click()
SaveList
Call MemoryCheck(WaitList)
End Sub
Private Sub LoadReservation(InDate As String)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim WRs As ADODB.Recordset
On Error GoTo ErrorHandler
If TerminalIsPresent(SvrPath.TerminalName) = False Then Exit Sub
Dim MinsOut As Integer
MinsOut = Val(GetConfigInfo("Reservations", "Transfer To Wait List Minutes", 60, "FUNCTIONS.INI"))
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
 Sql = "Select * from [Reservations] Where TRY_CAST([Date] AS datetime) = '" & InDate & "'"
 Sql = Sql & " and TRY_CAST([Time] AS datetime) < '" & DateAdd("n", MinsOut, Time) & "'"
If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
  Do While Not RS.EOF
   Sql = "Select * from [Wait List] Where [Day] = 0 and [Name] = '" & RS![Name] & "" & "'"
   Sql = Sql & " and [Preference] = '" & RS![Seating] & "" & "'"
   Sql = Sql & " and [Pager] = '" & RS![Phone] & "'"
   Sql = Sql & " and [Time In] = '" & RS![Time] & "'"
   If ConnectADORS(Sql, Db, WRs) = False Then Exit Sub
   If WRs.EOF Then
     WRs.AddNew
     WRs![Day] = 0
     WRs![Date] = Format(InDate, "mm/dd/yyyy")
     WRs![Seated] = flexUnchecked
     WRs![Name] = RS![Name] & ""
     WRs![Number in Party] = Val(RS![Number in Party] & "")
     WRs![Time In] = RS![Time] & ""
     WRs![Preference] = RS![Seating]  '& ""
     WRs![Quoted] = 0
     WRs![Time Seated] = ""
     WRs![Pager] = RS![Phone]
     WRs![Group] = "ALL"
     WRs![TextTime] = ""
     WRs.Update
     RS![Name] = ""
     RS.Update
   End If
  RS.MoveNext
Loop
Sql = "Delete from [Reservations] Where [Name] = ''"
Db.Execute Sql
DbClose Db, RS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "LOAD RESERVATION ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub LoadList(WaitingOnly As Boolean, Sql As String)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim dw As Integer
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
dw = 0
CList.Clear flexClearScrollable, flexClearEverything
CList.rows = 1
Dim Gw As Integer
Gw = 0
Do While Not RS.EOF
  dw = dw + 1
  With CList
    .rows = .rows + 1
    .Cell(flexcpChecked, dw, 0, dw, 0) = 2
    .Cell(flexcpChecked, dw, 0) = RS![Seated]
    .TextMatrix(dw, 1) = RS![Number in Party]
    If .Cell(flexcpChecked, dw, 0) = flexUnchecked Then
       Gw = Gw + Val(.TextMatrix(dw, 1))
    End If
    .TextMatrix(dw, 2) = RS![Name] & ""
    .TextMatrix(dw, 3) = RS![Time In] & ""
    If IsDate(RS![Time Seated]) Then
      .TextMatrix(dw, 4) = DateDiff("n", RS![Time In], RS![Time Seated])
    Else
      .TextMatrix(dw, 4) = DateDiff("n", RS![Time In], Time)
    End If
    .TextMatrix(dw, 5) = RS![Preference] & ""
    .TextMatrix(dw, 6) = RS![Quoted] & ""
    .TextMatrix(dw, 7) = 1
    .TextMatrix(dw, 8) = RS![Time Seated] & ""
    .TextMatrix(dw, 9) = RS![Pager] & ""
    .TextMatrix(dw, 10) = RS![Group] & ""
    .TextMatrix(dw, 11) = RS![TextTime] & ""
    If .Cell(flexcpChecked, dw, 0) = flexChecked Then
       .Cell(flexcpForeColor, dw, 1, dw, 10) = vbBlack
       .Cell(flexcpFontStrikethru, dw, 1, dw, 10) = True
    Else
      If Val(.TextMatrix(dw, 4)) >= Val(.TextMatrix(dw, 6)) Then
       .Cell(flexcpForeColor, dw, 1, dw, 10) = vbRed
      Else
       .Cell(flexcpForeColor, dw, 1, dw, 10) = vbBlue
      End If
    End If
    .Cell(flexcpBackColor, dw, 11, dw, 11) = vbYellow
    .RowHeight(dw) = 500
  End With
 RS.MoveNext
Loop
DbClose Db, RS
Parties.Caption = "Number of Parties Waiting " & dw
AverageWait.Caption = GetAverageWait
GuestWaiting.Caption = "Number of Guests Waiting " & Gw
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "LOAD List ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Function GetAverageWait() As String
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim dw As Integer
Dim WaitTime As Long

  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexChecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Function
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
 dw = 0: WaitTime = 0
Do While Not RS.EOF
 WaitTime = WaitTime + Abs(DateDiff("n", CDate(RS![Time In]), CDate(RS![Time Seated])))
 dw = dw + 1
 RS.MoveNext
Loop
DbClose Db, RS
If WaitTime > 0 And dw > 0 Then
  GetAverageWait = "Average Minutes for Seating " & CStr(WaitTime \ dw)
Else
  GetAverageWait = ""
End If
ERH:
Err.Clear
End Function
Private Sub SaveList()
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim dw As Integer

If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
dw = 0
For dw = 1 To CList.rows - 1
  With CList
   If .TextMatrix(dw, 2) = "" Then Exit For
   Sql = "Select * from [Wait List] Where [Day] = 0 and [Name] = '" & .TextMatrix(dw, 2) & "'"
   Sql = Sql & " and [Time In] = '" & .TextMatrix(dw, 3) & "'"
   If ConnectADORS(Sql, Db, RS) = False Then Exit Sub
   If RS.EOF Then
     RS.AddNew
     RS![Day] = 0
     RS![Date] = Format(Date, "mm/dd/yyyy")
   End If
     RS![Seated] = .Cell(flexcpChecked, dw, 0, dw, 0)
     RS![Number in Party] = Val(.TextMatrix(dw, 1))
     RS![Name] = .TextMatrix(dw, 2)
     RS![Time In] = .TextMatrix(dw, 3)
     RS![Preference] = .TextMatrix(dw, 5) & ""
     RS![Quoted] = .TextMatrix(dw, 6) & ""
     RS![Time Seated] = .TextMatrix(dw, 8) & ""
     RS![Pager] = .TextMatrix(dw, 9) & ""
     If .TextMatrix(dw, 10) & "" = "" Then
       RS![Group] = MyGroup
     Else
       RS![Group] = .TextMatrix(dw, 10) & ""
     End If
     RS.fields("TextTime").Value = .TextMatrix(dw, 11) & ""
     RS.Update
     SetNothing RS
   End With
Next dw
DbClose Db, RS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Save List ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand5_Click()
     Dim FindNames As String
     FindNames = ShowKeyBoard("Enter Customer Name") 'Enter_Name(False) '
     If FindNames = "" Then Exit Sub
     SaveList
       Sql = "Select * from [Wait List] Where [Name] like '" & FindNames & "%' and [Day] = 0 order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
     LoadList False, Sql
End Sub
Private Sub SSCommand6_Click()
SaveList
 Sql = "Select * from [Wait List] Where [Day] = 0 and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql
End Sub
Private Sub SSCommand7_Click()
CList.SetFocus
With CList
  If .Row - 10 < 0 Then
   .TopRow = 0
  Else
   .TopRow = .Row - 10
   .Row = .TopRow
  End If
End With
End Sub
Private Sub SSCommand8_Click()
LoadAllToList
With CList
  If .Row + 10 >= .rows Then Exit Sub
   .TopRow = .Row + 10
   .Row = .TopRow
End With

End Sub
Private Sub SSCommand9_Click()
Dim Db As ADODB.Connection
If See_Msg("Reset Customer Waiting List", 4, "Erases all seated customer on the List") = 6 Then
     ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
     Sql = "Update [Wait List] Set [day] = " & GetDateToNo(CDate(Date))
     Sql = Sql & " where [day] = 0 and [Seated] = 1;"
     Db.Execute Sql
     SetNothing Db
     See_Msg "Update Complete", 5, ""
End If
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql
End Sub
Private Sub Timer1_Timer()
Dim DA As Integer
Dim Gw As Integer
Gw = 0
  Sql = "Select * from [Wait List] Where [Day] = 0 and [Group] = '" & MyGroup & "' and [Seated] = " & flexUnchecked & " order by TRY_CAST([date] AS datetime),TRY_CAST([Time In] AS datetime)"
LoadList True, Sql

With CList
    For DA = 1 To .rows - 1
          If .TextMatrix(DA, 3) = "" Then Exit For
          If .Cell(flexcpChecked, DA, 0, DA, 0) = flexUnchecked Then
            If .Cell(flexcpChecked, DA, 0) = flexUnchecked Then
               Gw = Gw + Val(.TextMatrix(DA, 1))
            End If
            If IsDate(.TextMatrix(DA, 8)) Then
            .TextMatrix(DA, 4) = Val(DateDiff("n", CDate(.TextMatrix(DA, 3)), .TextMatrix(DA, 8)))
            Else
              .TextMatrix(DA, 4) = Val(DateDiff("n", CDate(.TextMatrix(DA, 3)), Time))
            End If
            If Val(.TextMatrix(DA, 4)) > Val(.TextMatrix(DA, 6)) Then
             .Cell(flexcpForeColor, DA, 1, DA, 6) = vbRed
            Else
             .Cell(flexcpForeColor, DA, 1, DA, 6) = vbBlack
            End If
          End If
    Next DA
End With
Parties.Caption = "Number of Parties Waiting " & DA - 1
GuestWaiting.Caption = "Number of Guests Waiting " & Gw
AverageWait.Caption = GetAverageWait
End Sub

