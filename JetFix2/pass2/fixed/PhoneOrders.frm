VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form PhoneOrders 
   Appearance      =   0  'Flat
   BackColor       =   &H002B1B09&
   BorderStyle     =   0  'None
   Caption         =   "Phone Orders"
   ClientHeight    =   11055
   ClientLeft      =   375
   ClientTop       =   495
   ClientWidth     =   15285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H002B1B09&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SearchBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00453A32&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00453A32&
      Height          =   9375
      Left            =   5400
      ScaleHeight     =   9345
      ScaleWidth      =   8505
      TabIndex        =   4
      Top             =   9720
      Visible         =   0   'False
      Width           =   8535
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   615
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   4560
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Phone Number"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":0000
         textLT          =   "PhoneOrders.frx":0078
         textCT          =   "PhoneOrders.frx":0090
         textRT          =   "PhoneOrders.frx":00A8
         textLM          =   "PhoneOrders.frx":00C0
         textRM          =   "PhoneOrders.frx":00D8
         textLB          =   "PhoneOrders.frx":00F0
         textCB          =   "PhoneOrders.frx":0108
         textRB          =   "PhoneOrders.frx":0120
         colorBack       =   "PhoneOrders.frx":0138
         colorIntern     =   "PhoneOrders.frx":0162
         colorMO         =   "PhoneOrders.frx":018C
         colorFocus      =   "PhoneOrders.frx":01B6
         colorDisabled   =   "PhoneOrders.frx":01E0
         colorPressed    =   "PhoneOrders.frx":020A
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   615
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   5280
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Street Address"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":0234
         textLT          =   "PhoneOrders.frx":02B0
         textCT          =   "PhoneOrders.frx":02C8
         textRT          =   "PhoneOrders.frx":02E0
         textLM          =   "PhoneOrders.frx":02F8
         textRM          =   "PhoneOrders.frx":0310
         textLB          =   "PhoneOrders.frx":0328
         textCB          =   "PhoneOrders.frx":0340
         textRB          =   "PhoneOrders.frx":0358
         colorBack       =   "PhoneOrders.frx":0370
         colorIntern     =   "PhoneOrders.frx":039A
         colorMO         =   "PhoneOrders.frx":03C4
         colorFocus      =   "PhoneOrders.frx":03EE
         colorDisabled   =   "PhoneOrders.frx":0418
         colorPressed    =   "PhoneOrders.frx":0442
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   615
         Index           =   3
         Left            =   600
         TabIndex        =   7
         Top             =   6000
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Driver Name"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":046C
         textLT          =   "PhoneOrders.frx":04E2
         textCT          =   "PhoneOrders.frx":04FA
         textRT          =   "PhoneOrders.frx":0512
         textLM          =   "PhoneOrders.frx":052A
         textRM          =   "PhoneOrders.frx":0542
         textLB          =   "PhoneOrders.frx":055A
         textCB          =   "PhoneOrders.frx":0572
         textRB          =   "PhoneOrders.frx":058A
         colorBack       =   "PhoneOrders.frx":05A2
         colorIntern     =   "PhoneOrders.frx":05CC
         colorMO         =   "PhoneOrders.frx":05F6
         colorFocus      =   "PhoneOrders.frx":0620
         colorDisabled   =   "PhoneOrders.frx":064A
         colorPressed    =   "PhoneOrders.frx":0674
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   6720
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Customer Name"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":069E
         textLT          =   "PhoneOrders.frx":0718
         textCT          =   "PhoneOrders.frx":0730
         textRT          =   "PhoneOrders.frx":0748
         textLM          =   "PhoneOrders.frx":0760
         textRM          =   "PhoneOrders.frx":0778
         textLB          =   "PhoneOrders.frx":0790
         textCB          =   "PhoneOrders.frx":07A8
         textRB          =   "PhoneOrders.frx":07C0
         colorBack       =   "PhoneOrders.frx":07D8
         colorIntern     =   "PhoneOrders.frx":0802
         colorMO         =   "PhoneOrders.frx":082C
         colorFocus      =   "PhoneOrders.frx":0856
         colorDisabled   =   "PhoneOrders.frx":0880
         colorPressed    =   "PhoneOrders.frx":08AA
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   735
         Index           =   4
         Left            =   4680
         TabIndex        =   9
         Top             =   5520
         Width           =   3255
         _Version        =   589828
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Open Time"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   12648447
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":08D4
         textLT          =   "PhoneOrders.frx":0946
         textCT          =   "PhoneOrders.frx":095E
         textRT          =   "PhoneOrders.frx":0976
         textLM          =   "PhoneOrders.frx":098E
         textRM          =   "PhoneOrders.frx":09A6
         textLB          =   "PhoneOrders.frx":09BE
         textCB          =   "PhoneOrders.frx":09D6
         textRB          =   "PhoneOrders.frx":09EE
         colorBack       =   "PhoneOrders.frx":0A06
         colorIntern     =   "PhoneOrders.frx":0A30
         colorMO         =   "PhoneOrders.frx":0A5A
         colorFocus      =   "PhoneOrders.frx":0A84
         colorDisabled   =   "PhoneOrders.frx":0AAE
         colorPressed    =   "PhoneOrders.frx":0AD8
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh FilterBy 
         Height          =   735
         Left            =   4680
         TabIndex        =   10
         Top             =   6360
         Width           =   3255
         _Version        =   589828
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Sort By Delivery/Pickup"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   12648447
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":0B02
         textLT          =   "PhoneOrders.frx":0B90
         textCT          =   "PhoneOrders.frx":0BA8
         textRT          =   "PhoneOrders.frx":0BC0
         textLM          =   "PhoneOrders.frx":0BD8
         textRM          =   "PhoneOrders.frx":0BF0
         textLB          =   "PhoneOrders.frx":0C08
         textCB          =   "PhoneOrders.frx":0C20
         textRB          =   "PhoneOrders.frx":0C38
         colorBack       =   "PhoneOrders.frx":0C50
         colorIntern     =   "PhoneOrders.frx":0C7A
         colorMO         =   "PhoneOrders.frx":0CA4
         colorFocus      =   "PhoneOrders.frx":0CCE
         colorDisabled   =   "PhoneOrders.frx":0CF8
         colorPressed    =   "PhoneOrders.frx":0D22
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh9 
         Height          =   975
         Left            =   360
         TabIndex        =   13
         Top             =   7920
         Width           =   7815
         _Version        =   589828
         _ExtentX        =   13785
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Cancel"
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
         textCaption     =   "PhoneOrders.frx":0D4C
         textLT          =   "PhoneOrders.frx":0DB8
         textCT          =   "PhoneOrders.frx":0DD0
         textRT          =   "PhoneOrders.frx":0DE8
         textLM          =   "PhoneOrders.frx":0E00
         textRM          =   "PhoneOrders.frx":0E18
         textLB          =   "PhoneOrders.frx":0E30
         textCB          =   "PhoneOrders.frx":0E48
         textRB          =   "PhoneOrders.frx":0E60
         colorBack       =   "PhoneOrders.frx":0E78
         colorIntern     =   "PhoneOrders.frx":0EA2
         colorMO         =   "PhoneOrders.frx":0ECC
         colorFocus      =   "PhoneOrders.frx":0EF6
         colorDisabled   =   "PhoneOrders.frx":0F20
         colorPressed    =   "PhoneOrders.frx":0F4A
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
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand3 
         Height          =   735
         Left            =   4680
         TabIndex        =   15
         Top             =   4680
         Width           =   3255
         _Version        =   589828
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Pick Up By Time"
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777152
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":0F74
         textLT          =   "PhoneOrders.frx":0FF2
         textCT          =   "PhoneOrders.frx":100A
         textRT          =   "PhoneOrders.frx":1022
         textLM          =   "PhoneOrders.frx":103A
         textRM          =   "PhoneOrders.frx":1052
         textLB          =   "PhoneOrders.frx":106A
         textCB          =   "PhoneOrders.frx":1082
         textRB          =   "PhoneOrders.frx":109A
         colorBack       =   "PhoneOrders.frx":10B2
         colorIntern     =   "PhoneOrders.frx":10DC
         colorMO         =   "PhoneOrders.frx":1106
         colorFocus      =   "PhoneOrders.frx":1130
         colorDisabled   =   "PhoneOrders.frx":115A
         colorPressed    =   "PhoneOrders.frx":1184
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
         Office2003ColorNorth=   16052204
         Office2003ColorSouth=   16052204
         Office2003ColorBorder=   14008501
         Office2003ColorNorthMouseOver=   16052204
         Office2003ColorSouthMouseOver=   16052204
         Office2003ColorBorderMouseOver=   16735744
         Office2003ColorNorthPressed=   14008501
         Office2003ColorSouthPressed=   14008501
         Office2003ColorBorderPressed=   16735744
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   975
         Left            =   4440
         TabIndex        =   16
         Top             =   120
         Width           =   3735
         _Version        =   589828
         _ExtentX        =   6588
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Transfer To Dining Room"
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
         textCaption     =   "PhoneOrders.frx":11AE
         textLT          =   "PhoneOrders.frx":123C
         textCT          =   "PhoneOrders.frx":1254
         textRT          =   "PhoneOrders.frx":126C
         textLM          =   "PhoneOrders.frx":1284
         textRM          =   "PhoneOrders.frx":129C
         textLB          =   "PhoneOrders.frx":12B4
         textCB          =   "PhoneOrders.frx":12CC
         textRB          =   "PhoneOrders.frx":12E4
         colorBack       =   "PhoneOrders.frx":12FC
         colorIntern     =   "PhoneOrders.frx":1326
         colorMO         =   "PhoneOrders.frx":1350
         colorFocus      =   "PhoneOrders.frx":137A
         colorDisabled   =   "PhoneOrders.frx":13A4
         colorPressed    =   "PhoneOrders.frx":13CE
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
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh11 
         Height          =   975
         Left            =   480
         TabIndex        =   18
         Top             =   1920
         Width           =   3615
         _Version        =   589828
         _ExtentX        =   6376
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Pending Summary"
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
         textCaption     =   "PhoneOrders.frx":13F8
         textLT          =   "PhoneOrders.frx":1476
         textCT          =   "PhoneOrders.frx":148E
         textRT          =   "PhoneOrders.frx":14A6
         textLM          =   "PhoneOrders.frx":14BE
         textRM          =   "PhoneOrders.frx":14D6
         textLB          =   "PhoneOrders.frx":14EE
         textCB          =   "PhoneOrders.frx":1506
         textRB          =   "PhoneOrders.frx":151E
         colorBack       =   "PhoneOrders.frx":1536
         colorIntern     =   "PhoneOrders.frx":1560
         colorMO         =   "PhoneOrders.frx":158A
         colorFocus      =   "PhoneOrders.frx":15B4
         colorDisabled   =   "PhoneOrders.frx":15DE
         colorPressed    =   "PhoneOrders.frx":1608
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
         Office2003ColorNorth=   4194432
         Office2003ColorSouth=   2424906
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   4194432
         Office2003ColorSouthMouseOver=   4194432
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   4194432
         Office2003ColorSouthPressed=   2424906
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh12 
         Height          =   975
         Left            =   4200
         TabIndex        =   19
         Top             =   1920
         Width           =   3615
         _Version        =   589828
         _ExtentX        =   6376
         _ExtentY        =   1720
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
         Surface         =   15
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":1632
         textLT          =   "PhoneOrders.frx":16AE
         textCT          =   "PhoneOrders.frx":16C6
         textRT          =   "PhoneOrders.frx":16DE
         textLM          =   "PhoneOrders.frx":16F6
         textRM          =   "PhoneOrders.frx":170E
         textLB          =   "PhoneOrders.frx":1726
         textCB          =   "PhoneOrders.frx":173E
         textRB          =   "PhoneOrders.frx":1756
         colorBack       =   "PhoneOrders.frx":176E
         colorIntern     =   "PhoneOrders.frx":1798
         colorMO         =   "PhoneOrders.frx":17C2
         colorFocus      =   "PhoneOrders.frx":17EC
         colorDisabled   =   "PhoneOrders.frx":1816
         colorPressed    =   "PhoneOrders.frx":1840
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
         Office2003ColorNorth=   4194432
         Office2003ColorSouth=   2424906
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   4194432
         Office2003ColorSouthMouseOver=   4194432
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   4194432
         Office2003ColorSouthPressed=   2424906
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh ReOpenCheck 
         Height          =   975
         Left            =   360
         TabIndex        =   57
         Top             =   120
         Width           =   3735
         _Version        =   589828
         _ExtentX        =   6588
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Re-open Phone Order"
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
         textCaption     =   "PhoneOrders.frx":186A
         textLT          =   "PhoneOrders.frx":18F0
         textCT          =   "PhoneOrders.frx":1908
         textRT          =   "PhoneOrders.frx":1920
         textLM          =   "PhoneOrders.frx":1938
         textRM          =   "PhoneOrders.frx":1950
         textLB          =   "PhoneOrders.frx":1968
         textCB          =   "PhoneOrders.frx":1980
         textRB          =   "PhoneOrders.frx":1998
         colorBack       =   "PhoneOrders.frx":19B0
         colorIntern     =   "PhoneOrders.frx":19DA
         colorMO         =   "PhoneOrders.frx":1A04
         colorFocus      =   "PhoneOrders.frx":1A2E
         colorDisabled   =   "PhoneOrders.frx":1A58
         colorPressed    =   "PhoneOrders.frx":1A82
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
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh18 
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   9120
         Visible         =   0   'False
         Width           =   2175
         _Version        =   589828
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   66
         Caption         =   "Change To Driver"
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
         textCaption     =   "PhoneOrders.frx":1AAC
         textLT          =   "PhoneOrders.frx":1B2C
         textCT          =   "PhoneOrders.frx":1B44
         textRT          =   "PhoneOrders.frx":1B5C
         textLM          =   "PhoneOrders.frx":1B74
         textRM          =   "PhoneOrders.frx":1B8C
         textLB          =   "PhoneOrders.frx":1BA4
         textCB          =   "PhoneOrders.frx":1BBC
         textRB          =   "PhoneOrders.frx":1BD4
         colorBack       =   "PhoneOrders.frx":1BEC
         colorIntern     =   "PhoneOrders.frx":1C16
         colorMO         =   "PhoneOrders.frx":1C40
         colorFocus      =   "PhoneOrders.frx":1C6A
         colorDisabled   =   "PhoneOrders.frx":1C94
         colorPressed    =   "PhoneOrders.frx":1CBE
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
         Office2003ColorNorth=   16512
         Office2003ColorSouth=   12642
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   16512
         Office2003ColorSouthMouseOver=   16512
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   16512
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Print Pending"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   7815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "List By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   9000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   360
         TabIndex        =   12
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sort By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   4440
         TabIndex        =   11
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   6015
         Left            =   840
         Top             =   4320
         Width           =   6735
      End
   End
   Begin VB.PictureBox CustomerINfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00453A32&
      FillColor       =   &H00453A32&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00453A32&
      Height          =   10815
      Left            =   -6600
      ScaleHeight     =   10785
      ScaleWidth      =   7365
      TabIndex        =   25
      Top             =   -240
      Visible         =   0   'False
      Width           =   7400
      Begin VB.TextBox PaymentAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00564940&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   4800
         TabIndex        =   99
         Text            =   "0.00"
         Top             =   6120
         Width           =   1815
      End
      Begin VB.TextBox CInfo 
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
         Height          =   1155
         Index           =   6
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   7920
         Width           =   5895
      End
      Begin VB.TextBox CInfo 
         Appearance      =   0  'Flat
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
         Height          =   400
         Index           =   5
         Left            =   720
         TabIndex        =   35
         Top             =   2160
         Width           =   5895
      End
      Begin VB.TextBox CInfo 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   400
         Index           =   4
         Left            =   5520
         TabIndex        =   34
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox CInfo 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   400
         Index           =   3
         Left            =   3120
         TabIndex        =   33
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox CInfo 
         Appearance      =   0  'Flat
         Height          =   400
         Index           =   2
         Left            =   720
         TabIndex        =   32
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox CInfo 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   400
         Index           =   1
         Left            =   720
         TabIndex        =   31
         Top             =   3240
         Width           =   5895
      End
      Begin VB.TextBox CInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   400
         Index           =   0
         Left            =   720
         TabIndex        =   30
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox CInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B72900&
         Height          =   600
         Index           =   7
         Left            =   720
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox CheckAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00564940&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   4800
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox PUTime 
         Appearance      =   0  'Flat
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
         Height          =   400
         Left            =   720
         TabIndex        =   27
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox PUDATE 
         Appearance      =   0  'Flat
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
         Height          =   400
         Left            =   2640
         TabIndex        =   26
         Top             =   6600
         Width           =   1695
      End
      Begin BTNENHLib4.BtnEnh sscommand6 
         Height          =   855
         Left            =   4680
         TabIndex        =   37
         Top             =   9480
         Width           =   1965
         _Version        =   589828
         _ExtentX        =   3466
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
         Surface         =   15
         BackColorContainer=   8421504
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":1CE8
         textLT          =   "PhoneOrders.frx":1D50
         textCT          =   "PhoneOrders.frx":1D68
         textRT          =   "PhoneOrders.frx":1D80
         textLM          =   "PhoneOrders.frx":1D98
         textRM          =   "PhoneOrders.frx":1DB0
         textLB          =   "PhoneOrders.frx":1DC8
         textCB          =   "PhoneOrders.frx":1DE0
         textRB          =   "PhoneOrders.frx":1DF8
         colorBack       =   "PhoneOrders.frx":1E10
         colorIntern     =   "PhoneOrders.frx":1E3A
         colorMO         =   "PhoneOrders.frx":1E64
         colorFocus      =   "PhoneOrders.frx":1E8E
         colorDisabled   =   "PhoneOrders.frx":1EB8
         colorPressed    =   "PhoneOrders.frx":1EE2
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
         Office2003ColorNorth=   16731678
         Office2003ColorSouth=   14102528
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   16731678
         Office2003ColorSouthMouseOver=   14102528
         Office2003ColorBorderMouseOver=   16740422
         Office2003ColorNorthPressed=   12003584
         Office2003ColorSouthPressed=   12003584
         Office2003ColorBorderPressed=   16740422
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   615
         Left            =   2880
         TabIndex        =   38
         Top             =   5280
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Pay Deposit or Partial Amount"
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
         BackColorContainer=   12632256
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":1F0C
         textLT          =   "PhoneOrders.frx":1FA6
         textCT          =   "PhoneOrders.frx":1FBE
         textRT          =   "PhoneOrders.frx":1FD6
         textLM          =   "PhoneOrders.frx":1FEE
         textRM          =   "PhoneOrders.frx":2006
         textLB          =   "PhoneOrders.frx":201E
         textCB          =   "PhoneOrders.frx":2036
         textRB          =   "PhoneOrders.frx":204E
         colorBack       =   "PhoneOrders.frx":2066
         colorIntern     =   "PhoneOrders.frx":2090
         colorMO         =   "PhoneOrders.frx":20BA
         colorFocus      =   "PhoneOrders.frx":20E4
         colorDisabled   =   "PhoneOrders.frx":210E
         colorPressed    =   "PhoneOrders.frx":2138
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
         Office2003ColorNorth=   16731678
         Office2003ColorSouth=   14102528
         Office2003ColorBorder=   2562065
         Office2003ColorNorthMouseOver=   16731678
         Office2003ColorSouthMouseOver=   14102528
         Office2003ColorBorderMouseOver=   16740422
         Office2003ColorNorthPressed=   12003584
         Office2003ColorSouthPressed=   12003584
         Office2003ColorBorderPressed=   16740422
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh5 
         Height          =   855
         Left            =   2640
         TabIndex        =   39
         Top             =   9480
         Width           =   1965
         _Version        =   589828
         _ExtentX        =   3466
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Prepay / Change Payment"
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
         CornerFactor    =   100
         Surface         =   15
         BackColorContainer=   12632256
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":2162
         textLT          =   "PhoneOrders.frx":21F0
         textCT          =   "PhoneOrders.frx":2208
         textRT          =   "PhoneOrders.frx":2220
         textLM          =   "PhoneOrders.frx":2238
         textRM          =   "PhoneOrders.frx":2250
         textLB          =   "PhoneOrders.frx":2268
         textCB          =   "PhoneOrders.frx":2280
         textRB          =   "PhoneOrders.frx":2298
         colorBack       =   "PhoneOrders.frx":22B0
         colorIntern     =   "PhoneOrders.frx":22DA
         colorMO         =   "PhoneOrders.frx":2304
         colorFocus      =   "PhoneOrders.frx":232E
         colorDisabled   =   "PhoneOrders.frx":2358
         colorPressed    =   "PhoneOrders.frx":2382
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
         Office2003ColorNorth=   4467486
         Office2003ColorSouth=   3547928
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   4467486
         Office2003ColorSouthMouseOver=   3547928
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   5189922
         Office2003ColorSouthPressed=   5189922
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   615
         Left            =   4680
         TabIndex        =   40
         Top             =   7200
         Visible         =   0   'False
         Width           =   1965
         _Version        =   589828
         _ExtentX        =   3466
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Cancel Order"
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
         CornerFactor    =   10
         Surface         =   2
         BackColorContainer=   12632256
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":23AC
         textLT          =   "PhoneOrders.frx":2424
         textCT          =   "PhoneOrders.frx":243C
         textRT          =   "PhoneOrders.frx":2454
         textLM          =   "PhoneOrders.frx":246C
         textRM          =   "PhoneOrders.frx":2484
         textLB          =   "PhoneOrders.frx":249C
         textCB          =   "PhoneOrders.frx":24B4
         textRB          =   "PhoneOrders.frx":24CC
         colorBack       =   "PhoneOrders.frx":24E4
         colorIntern     =   "PhoneOrders.frx":250E
         colorMO         =   "PhoneOrders.frx":2538
         colorFocus      =   "PhoneOrders.frx":2562
         colorDisabled   =   "PhoneOrders.frx":258C
         colorPressed    =   "PhoneOrders.frx":25B6
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
      Begin BTNENHLib4.BtnEnh BtnEnh14 
         Height          =   855
         Index           =   0
         Left            =   720
         TabIndex        =   41
         Top             =   9480
         Width           =   1845
         _Version        =   589828
         _ExtentX        =   3254
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Receipt"
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
         CornerFactor    =   100
         Surface         =   15
         BackColorContainer=   12632256
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":25E0
         textLT          =   "PhoneOrders.frx":265A
         textCT          =   "PhoneOrders.frx":2672
         textRT          =   "PhoneOrders.frx":268A
         textLM          =   "PhoneOrders.frx":26A2
         textRM          =   "PhoneOrders.frx":26BA
         textLB          =   "PhoneOrders.frx":26D2
         textCB          =   "PhoneOrders.frx":26EA
         textRB          =   "PhoneOrders.frx":2702
         colorBack       =   "PhoneOrders.frx":271A
         colorIntern     =   "PhoneOrders.frx":2744
         colorMO         =   "PhoneOrders.frx":276E
         colorFocus      =   "PhoneOrders.frx":2798
         colorDisabled   =   "PhoneOrders.frx":27C2
         colorPressed    =   "PhoneOrders.frx":27EC
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
         Office2003ColorNorth=   4467486
         Office2003ColorSouth=   3547928
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   4467486
         Office2003ColorSouthMouseOver=   3547928
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   5189922
         Office2003ColorSouthPressed=   5189922
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label WorkingCheck 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   2280
         TabIndex        =   100
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   4800
         TabIndex        =   98
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label OrderIsFor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   255
         Left            =   4920
         TabIndex        =   95
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Information"
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
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Directions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   54
         Top             =   7560
         Width           =   3375
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   53
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   52
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   51
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   50
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label CData 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   49
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   48
         Top             =   960
         Width           =   2565
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click information box for map and directions."
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
         Left            =   720
         TabIndex        =   47
         Tag             =   "0"
         Top             =   9120
         Width           =   5895
      End
      Begin VB.Label CData 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deposits Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F4EFEC&
         Height          =   345
         Index           =   7
         Left            =   720
         TabIndex        =   46
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label CTotal 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   4800
         TabIndex        =   45
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ready Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ready Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   2640
         TabIndex        =   43
         Top             =   6240
         Width           =   2415
      End
      Begin VB.Label PickUpInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Time Or Date To Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   7200
         Width           =   3615
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00453A32&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   12720
         Top             =   6720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00271811&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   800
      Left            =   120
      TabIndex        =   72
      Top             =   10560
      Width           =   16000
      Begin VB.Timer WatchingforNewDeliveries 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   2520
         Top             =   120
      End
      Begin BTNENHLib4.BtnEnh ShowEdit 
         Height          =   675
         Left            =   6360
         TabIndex        =   73
         Top             =   0
         Width           =   1860
         _Version        =   589828
         _ExtentX        =   3281
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Change-Cancel Order/Edit Customer"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         textCaption     =   "PhoneOrders.frx":2816
         textLT          =   "PhoneOrders.frx":28B8
         textCT          =   "PhoneOrders.frx":28D0
         textRT          =   "PhoneOrders.frx":28E8
         textLM          =   "PhoneOrders.frx":2900
         textRM          =   "PhoneOrders.frx":2918
         textLB          =   "PhoneOrders.frx":2930
         textCB          =   "PhoneOrders.frx":2948
         textRB          =   "PhoneOrders.frx":2960
         colorBack       =   "PhoneOrders.frx":2978
         colorIntern     =   "PhoneOrders.frx":29A2
         colorMO         =   "PhoneOrders.frx":29CC
         colorFocus      =   "PhoneOrders.frx":29F6
         colorDisabled   =   "PhoneOrders.frx":2A20
         colorPressed    =   "PhoneOrders.frx":2A4A
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
         Office2003ColorNorth=   0
         Office2003ColorSouth=   0
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   0
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   675
         Left            =   3000
         TabIndex        =   74
         Top             =   0
         Width           =   1620
         _Version        =   589828
         _ExtentX        =   2857
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Release Order"
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
         textCaption     =   "PhoneOrders.frx":2A74
         textLT          =   "PhoneOrders.frx":2AEE
         textCT          =   "PhoneOrders.frx":2B06
         textRT          =   "PhoneOrders.frx":2B1E
         textLM          =   "PhoneOrders.frx":2B36
         textRM          =   "PhoneOrders.frx":2B4E
         textLB          =   "PhoneOrders.frx":2B66
         textCB          =   "PhoneOrders.frx":2B7E
         textRB          =   "PhoneOrders.frx":2B96
         colorBack       =   "PhoneOrders.frx":2BAE
         colorIntern     =   "PhoneOrders.frx":2BD8
         colorMO         =   "PhoneOrders.frx":2C02
         colorFocus      =   "PhoneOrders.frx":2C2C
         colorDisabled   =   "PhoneOrders.frx":2C56
         colorPressed    =   "PhoneOrders.frx":2C80
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
         Office2003ColorNorth=   1908097
         Office2003ColorSouth=   1579120
         Office2003ColorBorder=   1579120
         Office2003ColorNorthMouseOver=   1908097
         Office2003ColorSouthMouseOver=   1579120
         Office2003ColorBorderMouseOver=   16740422
         Office2003ColorNorthPressed=   255
         Office2003ColorSouthPressed=   255
         Office2003ColorBorderPressed=   16740422
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand4 
         Height          =   675
         Left            =   8280
         TabIndex        =   75
         Top             =   0
         Width           =   1860
         _Version        =   589828
         _ExtentX        =   3281
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Assign Driver / Get Directions"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":2CAA
         textLT          =   "PhoneOrders.frx":2D46
         textCT          =   "PhoneOrders.frx":2D5E
         textRT          =   "PhoneOrders.frx":2D76
         textLM          =   "PhoneOrders.frx":2D8E
         textRM          =   "PhoneOrders.frx":2DA6
         textLB          =   "PhoneOrders.frx":2DBE
         textCB          =   "PhoneOrders.frx":2DD6
         textRB          =   "PhoneOrders.frx":2DEE
         colorBack       =   "PhoneOrders.frx":2E06
         colorIntern     =   "PhoneOrders.frx":2E30
         colorMO         =   "PhoneOrders.frx":2E5A
         colorFocus      =   "PhoneOrders.frx":2E84
         colorDisabled   =   "PhoneOrders.frx":2EAE
         colorPressed    =   "PhoneOrders.frx":2ED8
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
         Office2003ColorNorth=   0
         Office2003ColorSouth=   0
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   0
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand5 
         Height          =   675
         Left            =   11880
         TabIndex        =   76
         Top             =   0
         Width           =   1620
         _Version        =   589828
         _ExtentX        =   2857
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Close All"
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
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   2
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":2F02
         textLT          =   "PhoneOrders.frx":2F74
         textCT          =   "PhoneOrders.frx":2F8C
         textRT          =   "PhoneOrders.frx":2FA4
         textLM          =   "PhoneOrders.frx":2FBC
         textRM          =   "PhoneOrders.frx":2FD4
         textLB          =   "PhoneOrders.frx":2FEC
         textCB          =   "PhoneOrders.frx":3004
         textRB          =   "PhoneOrders.frx":301C
         colorBack       =   "PhoneOrders.frx":3034
         colorIntern     =   "PhoneOrders.frx":305E
         colorMO         =   "PhoneOrders.frx":3088
         colorFocus      =   "PhoneOrders.frx":30B2
         colorDisabled   =   "PhoneOrders.frx":30DC
         colorPressed    =   "PhoneOrders.frx":3106
         Orientation     =   3
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
         Office2003ColorSouth=   0
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   12870144
         Office2003ColorSouthMouseOver=   12870144
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   12870144
         Office2003ColorSouthPressed=   8404992
         Office2003ColorBorderPressed=   16777215
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh16 
         Height          =   675
         Left            =   10200
         TabIndex        =   77
         Top             =   0
         Width           =   1620
         _Version        =   589828
         _ExtentX        =   2857
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Mark Transaction Paid"
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
         BackColorContainer=   4210752
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":3130
         textLT          =   "PhoneOrders.frx":31BA
         textCT          =   "PhoneOrders.frx":31D2
         textRT          =   "PhoneOrders.frx":31EA
         textLM          =   "PhoneOrders.frx":3202
         textRM          =   "PhoneOrders.frx":321A
         textLB          =   "PhoneOrders.frx":3232
         textCB          =   "PhoneOrders.frx":324A
         textRB          =   "PhoneOrders.frx":3262
         colorBack       =   "PhoneOrders.frx":327A
         colorIntern     =   "PhoneOrders.frx":32A4
         colorMO         =   "PhoneOrders.frx":32CE
         colorFocus      =   "PhoneOrders.frx":32F8
         colorDisabled   =   "PhoneOrders.frx":3322
         colorPressed    =   "PhoneOrders.frx":334C
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
         Office2003ColorNorth=   0
         Office2003ColorSouth=   0
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   0
         Office2003ColorSouthMouseOver=   0
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand10 
         Height          =   675
         Left            =   13560
         TabIndex        =   78
         Top             =   0
         Width           =   1620
         _Version        =   589828
         _ExtentX        =   2857
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Driver Sales Report"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         textCaption     =   "PhoneOrders.frx":3376
         textLT          =   "PhoneOrders.frx":33FC
         textCT          =   "PhoneOrders.frx":3414
         textRT          =   "PhoneOrders.frx":342C
         textLM          =   "PhoneOrders.frx":3444
         textRM          =   "PhoneOrders.frx":345C
         textLB          =   "PhoneOrders.frx":3474
         textCB          =   "PhoneOrders.frx":348C
         textRB          =   "PhoneOrders.frx":34A4
         colorBack       =   "PhoneOrders.frx":34BC
         colorIntern     =   "PhoneOrders.frx":34E6
         colorMO         =   "PhoneOrders.frx":3510
         colorFocus      =   "PhoneOrders.frx":353A
         colorDisabled   =   "PhoneOrders.frx":3564
         colorPressed    =   "PhoneOrders.frx":358E
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
         Office2003ColorNorth=   16117997
         Office2003ColorSouth=   16117997
         Office2003ColorBorder=   16117997
         Office2003ColorNorthMouseOver=   16117997
         Office2003ColorSouthMouseOver=   16117997
         Office2003ColorBorderMouseOver=   16711680
         Office2003ColorNorthPressed=   11441540
         Office2003ColorSouthPressed=   11441540
         Office2003ColorBorderPressed=   16711680
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh10 
         Height          =   675
         Left            =   4680
         TabIndex        =   79
         Top             =   0
         Width           =   1620
         _Version        =   589828
         _ExtentX        =   2857
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Option Menu"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         textCaption     =   "PhoneOrders.frx":35B8
         textLT          =   "PhoneOrders.frx":362E
         textCT          =   "PhoneOrders.frx":3646
         textRT          =   "PhoneOrders.frx":365E
         textLM          =   "PhoneOrders.frx":3676
         textRM          =   "PhoneOrders.frx":368E
         textLB          =   "PhoneOrders.frx":36A6
         textCB          =   "PhoneOrders.frx":36BE
         textRB          =   "PhoneOrders.frx":36D6
         colorBack       =   "PhoneOrders.frx":36EE
         colorIntern     =   "PhoneOrders.frx":3718
         colorMO         =   "PhoneOrders.frx":3742
         colorFocus      =   "PhoneOrders.frx":376C
         colorDisabled   =   "PhoneOrders.frx":3796
         colorPressed    =   "PhoneOrders.frx":37C0
         Orientation     =   1
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
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   0
         Office2003ColorSouthPressed=   0
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand11 
         Height          =   675
         Left            =   50
         TabIndex        =   80
         Top             =   0
         Width           =   1140
         _Version        =   589828
         _ExtentX        =   2011
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Print Open Orders"
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
         CornerFactor    =   40
         Surface         =   15
         BackColorContainer=   2562065
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":37EA
         textLT          =   "PhoneOrders.frx":386C
         textCT          =   "PhoneOrders.frx":3884
         textRT          =   "PhoneOrders.frx":389C
         textLM          =   "PhoneOrders.frx":38B4
         textRM          =   "PhoneOrders.frx":38CC
         textLB          =   "PhoneOrders.frx":38E4
         textCB          =   "PhoneOrders.frx":38FC
         textRB          =   "PhoneOrders.frx":3914
         colorBack       =   "PhoneOrders.frx":392C
         colorIntern     =   "PhoneOrders.frx":3956
         colorMO         =   "PhoneOrders.frx":3980
         colorFocus      =   "PhoneOrders.frx":39AA
         colorDisabled   =   "PhoneOrders.frx":39D4
         colorPressed    =   "PhoneOrders.frx":39FE
         Orientation     =   1
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
         Office2003ColorNorth=   7624534
         Office2003ColorSouth=   7033168
         Office2003ColorBorder=   5324092
         Office2003ColorNorthMouseOver=   7624534
         Office2003ColorSouthMouseOver=   7033168
         Office2003ColorBorderMouseOver=   16740422
         Office2003ColorNorthPressed=   16415549
         Office2003ColorSouthPressed=   16415549
         Office2003ColorBorderPressed=   16740422
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh6 
         Height          =   675
         Left            =   1200
         TabIndex        =   81
         Top             =   0
         Width           =   1260
         _Version        =   589828
         _ExtentX        =   2222
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Print Order Slips"
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
         CornerFactor    =   40
         Surface         =   15
         BackColorContainer=   2562065
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":3A28
         textLT          =   "PhoneOrders.frx":3AAA
         textCT          =   "PhoneOrders.frx":3AC2
         textRT          =   "PhoneOrders.frx":3ADA
         textLM          =   "PhoneOrders.frx":3AF2
         textRM          =   "PhoneOrders.frx":3B0A
         textLB          =   "PhoneOrders.frx":3B22
         textCB          =   "PhoneOrders.frx":3B3A
         textRB          =   "PhoneOrders.frx":3B52
         colorBack       =   "PhoneOrders.frx":3B6A
         colorIntern     =   "PhoneOrders.frx":3B94
         colorMO         =   "PhoneOrders.frx":3BBE
         colorFocus      =   "PhoneOrders.frx":3BE8
         colorDisabled   =   "PhoneOrders.frx":3C12
         colorPressed    =   "PhoneOrders.frx":3C3C
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
         Office2003ColorNorth=   7624534
         Office2003ColorSouth=   7033168
         Office2003ColorBorder=   5324092
         Office2003ColorNorthMouseOver=   7624534
         Office2003ColorSouthMouseOver=   7033168
         Office2003ColorBorderMouseOver=   16740422
         Office2003ColorNorthPressed=   16415549
         Office2003ColorSouthPressed=   16415549
         Office2003ColorBorderPressed=   16740422
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00271811&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00271811&
      Height          =   700
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   16000
      Begin BTNENHLib4.BtnEnh BtnEnh8 
         Height          =   625
         Left            =   2280
         TabIndex        =   60
         Top             =   30
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Find Check"
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
         CornerFactor    =   100
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":3C66
         textLT          =   "PhoneOrders.frx":3CDA
         textCT          =   "PhoneOrders.frx":3CF2
         textRT          =   "PhoneOrders.frx":3D0A
         textLM          =   "PhoneOrders.frx":3D22
         textRM          =   "PhoneOrders.frx":3D3A
         textLB          =   "PhoneOrders.frx":3D52
         textCB          =   "PhoneOrders.frx":3D6A
         textRB          =   "PhoneOrders.frx":3D82
         colorBack       =   "PhoneOrders.frx":3D9A
         colorIntern     =   "PhoneOrders.frx":3DC4
         colorMO         =   "PhoneOrders.frx":3DEE
         colorFocus      =   "PhoneOrders.frx":3E18
         colorDisabled   =   "PhoneOrders.frx":3E42
         colorPressed    =   "PhoneOrders.frx":3E6C
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh17 
         Height          =   625
         Left            =   3465
         TabIndex        =   61
         Top             =   30
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Search By Phone"
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
         CornerFactor    =   100
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":3E96
         textLT          =   "PhoneOrders.frx":3F14
         textCT          =   "PhoneOrders.frx":3F2C
         textRT          =   "PhoneOrders.frx":3F44
         textLM          =   "PhoneOrders.frx":3F5C
         textRM          =   "PhoneOrders.frx":3F74
         textLB          =   "PhoneOrders.frx":3F8C
         textCB          =   "PhoneOrders.frx":3FA4
         textRB          =   "PhoneOrders.frx":3FBC
         colorBack       =   "PhoneOrders.frx":3FD4
         colorIntern     =   "PhoneOrders.frx":3FFE
         colorMO         =   "PhoneOrders.frx":4028
         colorFocus      =   "PhoneOrders.frx":4052
         colorDisabled   =   "PhoneOrders.frx":407C
         colorPressed    =   "PhoneOrders.frx":40A6
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh exit 
         Height          =   625
         Left            =   30
         TabIndex        =   62
         Top             =   30
         Width           =   1035
         _Version        =   589828
         _ExtentX        =   1826
         _ExtentY        =   1102
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":40D0
         textLT          =   "PhoneOrders.frx":4138
         textCT          =   "PhoneOrders.frx":4150
         textRT          =   "PhoneOrders.frx":4168
         textLM          =   "PhoneOrders.frx":4180
         textRM          =   "PhoneOrders.frx":4198
         textLB          =   "PhoneOrders.frx":41B0
         textCB          =   "PhoneOrders.frx":41C8
         textRB          =   "PhoneOrders.frx":41E0
         colorBack       =   "PhoneOrders.frx":41F8
         colorIntern     =   "PhoneOrders.frx":4222
         colorMO         =   "PhoneOrders.frx":424C
         colorFocus      =   "PhoneOrders.frx":4276
         colorDisabled   =   "PhoneOrders.frx":42A0
         colorPressed    =   "PhoneOrders.frx":42CA
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
         Office2003ColorNorth=   467
         Office2003ColorSouth=   202
         Office2003ColorBorder=   467
         Office2003ColorNorthMouseOver=   467
         Office2003ColorSouthMouseOver=   202
         Office2003ColorBorderMouseOver=   9275636
         Office2003ColorNorthPressed=   1182862
         Office2003ColorSouthPressed=   1182862
         Office2003ColorBorderPressed=   4932591
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   625
         Index           =   6
         Left            =   9840
         TabIndex        =   63
         Top             =   30
         Width           =   1605
         _Version        =   589828
         _ExtentX        =   2831
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Show Todays Pending Orders"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         CornerFactor    =   40
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":42F4
         textLT          =   "PhoneOrders.frx":4388
         textCT          =   "PhoneOrders.frx":43A0
         textRT          =   "PhoneOrders.frx":43B8
         textLM          =   "PhoneOrders.frx":43D0
         textRM          =   "PhoneOrders.frx":43E8
         textLB          =   "PhoneOrders.frx":4400
         textCB          =   "PhoneOrders.frx":4418
         textRB          =   "PhoneOrders.frx":4430
         colorBack       =   "PhoneOrders.frx":4448
         colorIntern     =   "PhoneOrders.frx":4472
         colorMO         =   "PhoneOrders.frx":449C
         colorFocus      =   "PhoneOrders.frx":44C6
         colorDisabled   =   "PhoneOrders.frx":44F0
         colorPressed    =   "PhoneOrders.frx":451A
         Orientation     =   1
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
         Office2003ColorNorth=   16773096
         Office2003ColorSouth=   16773096
         Office2003ColorBorder=   16773096
         Office2003ColorNorthMouseOver=   16773096
         Office2003ColorSouthMouseOver=   16773096
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   16764087
         Office2003ColorSouthPressed=   16764087
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   630
         Index           =   7
         Left            =   11450
         TabIndex        =   64
         Top             =   30
         Width           =   1590
         _Version        =   589828
         _ExtentX        =   2805
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Select Date Pending Orders"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         CornerFactor    =   40
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":4544
         textLT          =   "PhoneOrders.frx":45D8
         textCT          =   "PhoneOrders.frx":45F0
         textRT          =   "PhoneOrders.frx":4608
         textLM          =   "PhoneOrders.frx":4620
         textRM          =   "PhoneOrders.frx":4638
         textLB          =   "PhoneOrders.frx":4650
         textCB          =   "PhoneOrders.frx":4668
         textRB          =   "PhoneOrders.frx":4680
         colorBack       =   "PhoneOrders.frx":4698
         colorIntern     =   "PhoneOrders.frx":46C2
         colorMO         =   "PhoneOrders.frx":46EC
         colorFocus      =   "PhoneOrders.frx":4716
         colorDisabled   =   "PhoneOrders.frx":4740
         colorPressed    =   "PhoneOrders.frx":476A
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
         Office2003ColorNorth=   16773096
         Office2003ColorSouth=   16773096
         Office2003ColorBorder=   16773096
         Office2003ColorNorthMouseOver=   16773096
         Office2003ColorSouthMouseOver=   16773096
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   16764087
         Office2003ColorSouthPressed=   16764087
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh13 
         Height          =   625
         Left            =   4650
         TabIndex        =   65
         Top             =   30
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Sort by Driver"
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
         CornerFactor    =   100
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":4794
         textLT          =   "PhoneOrders.frx":4810
         textCT          =   "PhoneOrders.frx":4828
         textRT          =   "PhoneOrders.frx":4840
         textLM          =   "PhoneOrders.frx":4858
         textRM          =   "PhoneOrders.frx":4870
         textLB          =   "PhoneOrders.frx":4888
         textCB          =   "PhoneOrders.frx":48A0
         textRB          =   "PhoneOrders.frx":48B8
         colorBack       =   "PhoneOrders.frx":48D0
         colorIntern     =   "PhoneOrders.frx":48FA
         colorMO         =   "PhoneOrders.frx":4924
         colorFocus      =   "PhoneOrders.frx":494E
         colorDisabled   =   "PhoneOrders.frx":4978
         colorPressed    =   "PhoneOrders.frx":49A2
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SearchHow 
         Height          =   625
         Index           =   5
         Left            =   8205
         TabIndex        =   66
         Top             =   30
         Width           =   1500
         _Version        =   589828
         _ExtentX        =   2646
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Unassigned Deliveries"
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
         CornerFactor    =   100
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":49CC
         textLT          =   "PhoneOrders.frx":4A56
         textCT          =   "PhoneOrders.frx":4A6E
         textRT          =   "PhoneOrders.frx":4A86
         textLM          =   "PhoneOrders.frx":4A9E
         textRM          =   "PhoneOrders.frx":4AB6
         textLB          =   "PhoneOrders.frx":4ACE
         textCB          =   "PhoneOrders.frx":4AE6
         textRB          =   "PhoneOrders.frx":4AFE
         colorBack       =   "PhoneOrders.frx":4B16
         colorIntern     =   "PhoneOrders.frx":4B40
         colorMO         =   "PhoneOrders.frx":4B6A
         colorFocus      =   "PhoneOrders.frx":4B94
         colorDisabled   =   "PhoneOrders.frx":4BBE
         colorPressed    =   "PhoneOrders.frx":4BE8
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand2 
         Height          =   625
         Left            =   7020
         TabIndex        =   67
         Top             =   30
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Show Delivery Only"
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
         CornerFactor    =   100
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":4C12
         textLT          =   "PhoneOrders.frx":4C96
         textCT          =   "PhoneOrders.frx":4CAE
         textRT          =   "PhoneOrders.frx":4CC6
         textLM          =   "PhoneOrders.frx":4CDE
         textRM          =   "PhoneOrders.frx":4CF6
         textLB          =   "PhoneOrders.frx":4D0E
         textCB          =   "PhoneOrders.frx":4D26
         textRB          =   "PhoneOrders.frx":4D3E
         colorBack       =   "PhoneOrders.frx":4D56
         colorIntern     =   "PhoneOrders.frx":4D80
         colorMO         =   "PhoneOrders.frx":4DAA
         colorFocus      =   "PhoneOrders.frx":4DD4
         colorDisabled   =   "PhoneOrders.frx":4DFE
         colorPressed    =   "PhoneOrders.frx":4E28
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand1 
         Height          =   625
         Left            =   5835
         TabIndex        =   68
         Top             =   30
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Show Pickup Only"
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
         CornerFactor    =   100
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":4E52
         textLT          =   "PhoneOrders.frx":4ED2
         textCT          =   "PhoneOrders.frx":4EEA
         textRT          =   "PhoneOrders.frx":4F02
         textLM          =   "PhoneOrders.frx":4F1A
         textRM          =   "PhoneOrders.frx":4F32
         textLB          =   "PhoneOrders.frx":4F4A
         textCB          =   "PhoneOrders.frx":4F62
         textRB          =   "PhoneOrders.frx":4F7A
         colorBack       =   "PhoneOrders.frx":4F92
         colorIntern     =   "PhoneOrders.frx":4FBC
         colorMO         =   "PhoneOrders.frx":4FE6
         colorFocus      =   "PhoneOrders.frx":5010
         colorDisabled   =   "PhoneOrders.frx":503A
         colorPressed    =   "PhoneOrders.frx":5064
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand15 
         Height          =   625
         Left            =   1095
         TabIndex        =   69
         Top             =   30
         Width           =   1155
         _Version        =   589828
         _ExtentX        =   2037
         _ExtentY        =   1102
         _StockProps     =   66
         Caption         =   "Show All"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         CornerFactor    =   100
         Surface         =   15
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":508E
         textLT          =   "PhoneOrders.frx":50FE
         textCT          =   "PhoneOrders.frx":5116
         textRT          =   "PhoneOrders.frx":512E
         textLM          =   "PhoneOrders.frx":5146
         textRM          =   "PhoneOrders.frx":515E
         textLB          =   "PhoneOrders.frx":5176
         textCB          =   "PhoneOrders.frx":518E
         textRB          =   "PhoneOrders.frx":51A6
         colorBack       =   "PhoneOrders.frx":51BE
         colorIntern     =   "PhoneOrders.frx":51E8
         colorMO         =   "PhoneOrders.frx":5212
         colorFocus      =   "PhoneOrders.frx":523C
         colorDisabled   =   "PhoneOrders.frx":5266
         colorPressed    =   "PhoneOrders.frx":5290
         Orientation     =   1
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
         Office2003ColorNorth=   16775413
         Office2003ColorSouth=   16775413
         Office2003ColorBorder=   13023930
         Office2003ColorNorthMouseOver=   16775413
         Office2003ColorSouthMouseOver=   16775413
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   13153965
         Office2003ColorSouthPressed=   13153965
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh Up_Arrow 
         Height          =   630
         Left            =   14280
         TabIndex        =   70
         Top             =   30
         Width           =   1020
         _Version        =   589828
         _ExtentX        =   1799
         _ExtentY        =   1102
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
         Surface         =   15
         Picture         =   "C:\Users\ben\Desktop\arrowup.bmp"
         BackColorContainer=   4210752
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "PhoneOrders.frx":52BA
         textCaption     =   "PhoneOrders.frx":53FE
         textLT          =   "PhoneOrders.frx":5416
         textCT          =   "PhoneOrders.frx":542E
         textRT          =   "PhoneOrders.frx":5446
         textLM          =   "PhoneOrders.frx":545E
         textRM          =   "PhoneOrders.frx":5476
         textLB          =   "PhoneOrders.frx":548E
         textCB          =   "PhoneOrders.frx":54A6
         textRB          =   "PhoneOrders.frx":54BE
         colorBack       =   "PhoneOrders.frx":54D6
         colorIntern     =   "PhoneOrders.frx":5500
         colorMO         =   "PhoneOrders.frx":552A
         colorFocus      =   "PhoneOrders.frx":5554
         colorDisabled   =   "PhoneOrders.frx":557E
         colorPressed    =   "PhoneOrders.frx":55A8
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
         Office2003ColorNorth=   2368806
         Office2003ColorSouth=   1645083
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   2368806
         Office2003ColorSouthMouseOver=   1645083
         Office2003ColorBorderMouseOver=   16731678
         Office2003ColorNorthPressed=   3421752
         Office2003ColorSouthPressed=   3421752
         Office2003ColorBorderPressed=   16731678
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh Dwn_arrow 
         Height          =   630
         Left            =   13200
         TabIndex        =   71
         Top             =   30
         Width           =   1020
         _Version        =   589828
         _ExtentX        =   1799
         _ExtentY        =   1102
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
         Surface         =   15
         Picture         =   "C:\Users\ben\Desktop\arrowdown.bmp"
         BackColorContainer=   4210752
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "PhoneOrders.frx":55D2
         textCaption     =   "PhoneOrders.frx":5709
         textLT          =   "PhoneOrders.frx":5721
         textCT          =   "PhoneOrders.frx":5739
         textRT          =   "PhoneOrders.frx":5751
         textLM          =   "PhoneOrders.frx":5769
         textRM          =   "PhoneOrders.frx":5781
         textLB          =   "PhoneOrders.frx":5799
         textCB          =   "PhoneOrders.frx":57B1
         textRB          =   "PhoneOrders.frx":57C9
         colorBack       =   "PhoneOrders.frx":57E1
         colorIntern     =   "PhoneOrders.frx":580B
         colorMO         =   "PhoneOrders.frx":5835
         colorFocus      =   "PhoneOrders.frx":585F
         colorDisabled   =   "PhoneOrders.frx":5889
         colorPressed    =   "PhoneOrders.frx":58B3
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
         Office2003ColorNorth=   2368806
         Office2003ColorSouth=   1645083
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   2368806
         Office2003ColorSouthMouseOver=   1645083
         Office2003ColorBorderMouseOver=   16731678
         Office2003ColorNorthPressed=   3421752
         Office2003ColorSouthPressed=   3421752
         Office2003ColorBorderPressed=   16731678
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
   End
   Begin VB.Frame ShowOrders 
      Appearance      =   0  'Flat
      BackColor       =   &H00453A32&
      BorderStyle     =   0  'None
      Caption         =   "Phone Orders"
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
      Height          =   8055
      Left            =   13320
      TabIndex        =   20
      Top             =   -2400
      Visible         =   0   'False
      Width           =   5655
      Begin VSFlex8DAOCtl.VSFlexGrid EditCustomer 
         Height          =   6615
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   5415
         _cx             =   2000037199
         _cy             =   2000039316
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   4536882
         BackColorAlternate=   14737632
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
         Cols            =   10
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
      Begin BTNENHLib4.BtnEnh BtnEnh15 
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   7080
         Width           =   5415
         _Version        =   589828
         _ExtentX        =   9551
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Cancel"
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
         CornerFactor    =   100
         Surface         =   15
         BackColorContainer=   4210752
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":58DD
         textLT          =   "PhoneOrders.frx":5949
         textCT          =   "PhoneOrders.frx":5961
         textRT          =   "PhoneOrders.frx":5979
         textLM          =   "PhoneOrders.frx":5991
         textRM          =   "PhoneOrders.frx":59A9
         textLB          =   "PhoneOrders.frx":59C1
         textCB          =   "PhoneOrders.frx":59D9
         textRB          =   "PhoneOrders.frx":59F1
         colorBack       =   "PhoneOrders.frx":5A09
         colorIntern     =   "PhoneOrders.frx":5A33
         colorMO         =   "PhoneOrders.frx":5A5D
         colorFocus      =   "PhoneOrders.frx":5A87
         colorDisabled   =   "PhoneOrders.frx":5AB1
         colorPressed    =   "PhoneOrders.frx":5ADB
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
         Office2003ColorNorth=   2894892
         Office2003ColorSouth=   2302755
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   2894892
         Office2003ColorSouthMouseOver=   2302755
         Office2003ColorBorderMouseOver=   16733952
         Office2003ColorNorthPressed=   1644825
         Office2003ColorSouthPressed=   1644825
         Office2003ColorBorderPressed=   16733952
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
   End
   Begin VB.PictureBox AssignDrv 
      Appearance      =   0  'Flat
      BackColor       =   &H00453A32&
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
      Height          =   8175
      Left            =   4440
      ScaleHeight     =   8145
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   -6840
      Visible         =   0   'False
      Width           =   5715
      Begin VB.ListBox DriverName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5970
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   5175
      End
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   7080
         Width           =   5175
         _Version        =   589828
         _ExtentX        =   9128
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Cancel"
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
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":5B05
         textLT          =   "PhoneOrders.frx":5B71
         textCT          =   "PhoneOrders.frx":5B89
         textRT          =   "PhoneOrders.frx":5BA1
         textLM          =   "PhoneOrders.frx":5BB9
         textRM          =   "PhoneOrders.frx":5BD1
         textLB          =   "PhoneOrders.frx":5BE9
         textCB          =   "PhoneOrders.frx":5C01
         textRB          =   "PhoneOrders.frx":5C19
         colorBack       =   "PhoneOrders.frx":5C31
         colorIntern     =   "PhoneOrders.frx":5C5B
         colorMO         =   "PhoneOrders.frx":5C85
         colorFocus      =   "PhoneOrders.frx":5CAF
         colorDisabled   =   "PhoneOrders.frx":5CD9
         colorPressed    =   "PhoneOrders.frx":5D03
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
         Office2003ColorNorth=   2368806
         Office2003ColorSouth=   1645083
         Office2003ColorBorder=   1645083
         Office2003ColorNorthMouseOver=   2368806
         Office2003ColorSouthMouseOver=   1645083
         Office2003ColorBorderMouseOver=   16731678
         Office2003ColorNorthPressed=   3421752
         Office2003ColorSouthPressed=   3421752
         Office2003ColorBorderPressed=   16731678
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label OrderAddress 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Assignment List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   97
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label AssignmentCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Assignment List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.Frame CurrentPhoneOrders 
      Appearance      =   0  'Flat
      BackColor       =   &H00271811&
      Caption         =   "Open Orders"
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
      Height          =   9650
      Left            =   0
      TabIndex        =   23
      Top             =   680
      Width           =   15800
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   3445
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":5D2D
         textLT          =   "PhoneOrders.frx":5D45
         textCT          =   "PhoneOrders.frx":5D5D
         textRT          =   "PhoneOrders.frx":5D75
         textLM          =   "PhoneOrders.frx":5D8D
         textRM          =   "PhoneOrders.frx":5DA5
         textLB          =   "PhoneOrders.frx":5DBD
         textCB          =   "PhoneOrders.frx":5DD5
         textRB          =   "PhoneOrders.frx":5DED
         colorBack       =   "PhoneOrders.frx":5E05
         colorIntern     =   "PhoneOrders.frx":5E2F
         colorMO         =   "PhoneOrders.frx":5E59
         colorFocus      =   "PhoneOrders.frx":5E83
         colorDisabled   =   "PhoneOrders.frx":5EAD
         colorPressed    =   "PhoneOrders.frx":5ED7
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":5F01
         Cell1           =   "PhoneOrders.frx":6015
         Cell2           =   "PhoneOrders.frx":60FD
         Cell3           =   "PhoneOrders.frx":61E9
         Cell4           =   "PhoneOrders.frx":6DDF
         Cell5           =   "PhoneOrders.frx":79D3
         Cell6           =   "PhoneOrders.frx":7AB3
         Cell7           =   "PhoneOrders.frx":7BA1
         Cell8           =   "PhoneOrders.frx":7C79
         Cell9           =   "PhoneOrders.frx":7D59
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   2
         Left            =   120
         TabIndex        =   83
         Top             =   6545
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":8469
         textLT          =   "PhoneOrders.frx":8481
         textCT          =   "PhoneOrders.frx":8499
         textRT          =   "PhoneOrders.frx":84B1
         textLM          =   "PhoneOrders.frx":84C9
         textRM          =   "PhoneOrders.frx":84E1
         textLB          =   "PhoneOrders.frx":84F9
         textCB          =   "PhoneOrders.frx":8511
         textRB          =   "PhoneOrders.frx":8529
         colorBack       =   "PhoneOrders.frx":8541
         colorIntern     =   "PhoneOrders.frx":856B
         colorMO         =   "PhoneOrders.frx":8595
         colorFocus      =   "PhoneOrders.frx":85BF
         colorDisabled   =   "PhoneOrders.frx":85E9
         colorPressed    =   "PhoneOrders.frx":8613
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":863D
         Cell1           =   "PhoneOrders.frx":8751
         Cell2           =   "PhoneOrders.frx":8839
         Cell3           =   "PhoneOrders.frx":8925
         Cell4           =   "PhoneOrders.frx":951B
         Cell5           =   "PhoneOrders.frx":A10F
         Cell6           =   "PhoneOrders.frx":A1EF
         Cell7           =   "PhoneOrders.frx":A2DD
         Cell8           =   "PhoneOrders.frx":A3B5
         Cell9           =   "PhoneOrders.frx":A495
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   3
         Left            =   3160
         TabIndex        =   84
         Top             =   345
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":ABA5
         textLT          =   "PhoneOrders.frx":ABBD
         textCT          =   "PhoneOrders.frx":ABD5
         textRT          =   "PhoneOrders.frx":ABED
         textLM          =   "PhoneOrders.frx":AC05
         textRM          =   "PhoneOrders.frx":AC1D
         textLB          =   "PhoneOrders.frx":AC35
         textCB          =   "PhoneOrders.frx":AC4D
         textRB          =   "PhoneOrders.frx":AC65
         colorBack       =   "PhoneOrders.frx":AC7D
         colorIntern     =   "PhoneOrders.frx":ACA7
         colorMO         =   "PhoneOrders.frx":ACD1
         colorFocus      =   "PhoneOrders.frx":ACFB
         colorDisabled   =   "PhoneOrders.frx":AD25
         colorPressed    =   "PhoneOrders.frx":AD4F
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":AD79
         Cell1           =   "PhoneOrders.frx":AE8D
         Cell2           =   "PhoneOrders.frx":AF75
         Cell3           =   "PhoneOrders.frx":B061
         Cell4           =   "PhoneOrders.frx":BC57
         Cell5           =   "PhoneOrders.frx":C84B
         Cell6           =   "PhoneOrders.frx":C92B
         Cell7           =   "PhoneOrders.frx":CA19
         Cell8           =   "PhoneOrders.frx":CAF1
         Cell9           =   "PhoneOrders.frx":CBD1
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   4
         Left            =   3160
         TabIndex        =   85
         Top             =   3445
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":D2E1
         textLT          =   "PhoneOrders.frx":D2F9
         textCT          =   "PhoneOrders.frx":D311
         textRT          =   "PhoneOrders.frx":D329
         textLM          =   "PhoneOrders.frx":D341
         textRM          =   "PhoneOrders.frx":D359
         textLB          =   "PhoneOrders.frx":D371
         textCB          =   "PhoneOrders.frx":D389
         textRB          =   "PhoneOrders.frx":D3A1
         colorBack       =   "PhoneOrders.frx":D3B9
         colorIntern     =   "PhoneOrders.frx":D3E3
         colorMO         =   "PhoneOrders.frx":D40D
         colorFocus      =   "PhoneOrders.frx":D437
         colorDisabled   =   "PhoneOrders.frx":D461
         colorPressed    =   "PhoneOrders.frx":D48B
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":D4B5
         Cell1           =   "PhoneOrders.frx":D5C9
         Cell2           =   "PhoneOrders.frx":D6B1
         Cell3           =   "PhoneOrders.frx":D79D
         Cell4           =   "PhoneOrders.frx":E393
         Cell5           =   "PhoneOrders.frx":EF87
         Cell6           =   "PhoneOrders.frx":F067
         Cell7           =   "PhoneOrders.frx":F155
         Cell8           =   "PhoneOrders.frx":F22D
         Cell9           =   "PhoneOrders.frx":F30D
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   5
         Left            =   3160
         TabIndex        =   86
         Top             =   6545
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":FA1D
         textLT          =   "PhoneOrders.frx":FA35
         textCT          =   "PhoneOrders.frx":FA4D
         textRT          =   "PhoneOrders.frx":FA65
         textLM          =   "PhoneOrders.frx":FA7D
         textRM          =   "PhoneOrders.frx":FA95
         textLB          =   "PhoneOrders.frx":FAAD
         textCB          =   "PhoneOrders.frx":FAC5
         textRB          =   "PhoneOrders.frx":FADD
         colorBack       =   "PhoneOrders.frx":FAF5
         colorIntern     =   "PhoneOrders.frx":FB1F
         colorMO         =   "PhoneOrders.frx":FB49
         colorFocus      =   "PhoneOrders.frx":FB73
         colorDisabled   =   "PhoneOrders.frx":FB9D
         colorPressed    =   "PhoneOrders.frx":FBC7
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":FBF1
         Cell1           =   "PhoneOrders.frx":FD05
         Cell2           =   "PhoneOrders.frx":FDED
         Cell3           =   "PhoneOrders.frx":FED9
         Cell4           =   "PhoneOrders.frx":10ACF
         Cell5           =   "PhoneOrders.frx":116C3
         Cell6           =   "PhoneOrders.frx":117A3
         Cell7           =   "PhoneOrders.frx":11891
         Cell8           =   "PhoneOrders.frx":11969
         Cell9           =   "PhoneOrders.frx":11A49
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   6
         Left            =   6200
         TabIndex        =   87
         Top             =   345
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":12159
         textLT          =   "PhoneOrders.frx":12171
         textCT          =   "PhoneOrders.frx":12189
         textRT          =   "PhoneOrders.frx":121A1
         textLM          =   "PhoneOrders.frx":121B9
         textRM          =   "PhoneOrders.frx":121D1
         textLB          =   "PhoneOrders.frx":121E9
         textCB          =   "PhoneOrders.frx":12201
         textRB          =   "PhoneOrders.frx":12219
         colorBack       =   "PhoneOrders.frx":12231
         colorIntern     =   "PhoneOrders.frx":1225B
         colorMO         =   "PhoneOrders.frx":12285
         colorFocus      =   "PhoneOrders.frx":122AF
         colorDisabled   =   "PhoneOrders.frx":122D9
         colorPressed    =   "PhoneOrders.frx":12303
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":1232D
         Cell1           =   "PhoneOrders.frx":12441
         Cell2           =   "PhoneOrders.frx":12529
         Cell3           =   "PhoneOrders.frx":12615
         Cell4           =   "PhoneOrders.frx":1320B
         Cell5           =   "PhoneOrders.frx":13DFF
         Cell6           =   "PhoneOrders.frx":13EDF
         Cell7           =   "PhoneOrders.frx":13FCD
         Cell8           =   "PhoneOrders.frx":140A5
         Cell9           =   "PhoneOrders.frx":14185
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   7
         Left            =   6200
         TabIndex        =   88
         Top             =   3445
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":14895
         textLT          =   "PhoneOrders.frx":148AD
         textCT          =   "PhoneOrders.frx":148C5
         textRT          =   "PhoneOrders.frx":148DD
         textLM          =   "PhoneOrders.frx":148F5
         textRM          =   "PhoneOrders.frx":1490D
         textLB          =   "PhoneOrders.frx":14925
         textCB          =   "PhoneOrders.frx":1493D
         textRB          =   "PhoneOrders.frx":14955
         colorBack       =   "PhoneOrders.frx":1496D
         colorIntern     =   "PhoneOrders.frx":14997
         colorMO         =   "PhoneOrders.frx":149C1
         colorFocus      =   "PhoneOrders.frx":149EB
         colorDisabled   =   "PhoneOrders.frx":14A15
         colorPressed    =   "PhoneOrders.frx":14A3F
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":14A69
         Cell1           =   "PhoneOrders.frx":14B7D
         Cell2           =   "PhoneOrders.frx":14C65
         Cell3           =   "PhoneOrders.frx":14D51
         Cell4           =   "PhoneOrders.frx":15947
         Cell5           =   "PhoneOrders.frx":1653B
         Cell6           =   "PhoneOrders.frx":1661B
         Cell7           =   "PhoneOrders.frx":16709
         Cell8           =   "PhoneOrders.frx":167E1
         Cell9           =   "PhoneOrders.frx":168C1
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   8
         Left            =   6200
         TabIndex        =   89
         Top             =   6545
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":16FD1
         textLT          =   "PhoneOrders.frx":16FE9
         textCT          =   "PhoneOrders.frx":17001
         textRT          =   "PhoneOrders.frx":17019
         textLM          =   "PhoneOrders.frx":17031
         textRM          =   "PhoneOrders.frx":17049
         textLB          =   "PhoneOrders.frx":17061
         textCB          =   "PhoneOrders.frx":17079
         textRB          =   "PhoneOrders.frx":17091
         colorBack       =   "PhoneOrders.frx":170A9
         colorIntern     =   "PhoneOrders.frx":170D3
         colorMO         =   "PhoneOrders.frx":170FD
         colorFocus      =   "PhoneOrders.frx":17127
         colorDisabled   =   "PhoneOrders.frx":17151
         colorPressed    =   "PhoneOrders.frx":1717B
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":171A5
         Cell1           =   "PhoneOrders.frx":172B9
         Cell2           =   "PhoneOrders.frx":173A1
         Cell3           =   "PhoneOrders.frx":1748D
         Cell4           =   "PhoneOrders.frx":18083
         Cell5           =   "PhoneOrders.frx":18C77
         Cell6           =   "PhoneOrders.frx":18D57
         Cell7           =   "PhoneOrders.frx":18E45
         Cell8           =   "PhoneOrders.frx":18F1D
         Cell9           =   "PhoneOrders.frx":18FFD
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   9
         Left            =   9240
         TabIndex        =   90
         Top             =   360
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":1970D
         textLT          =   "PhoneOrders.frx":19725
         textCT          =   "PhoneOrders.frx":1973D
         textRT          =   "PhoneOrders.frx":19755
         textLM          =   "PhoneOrders.frx":1976D
         textRM          =   "PhoneOrders.frx":19785
         textLB          =   "PhoneOrders.frx":1979D
         textCB          =   "PhoneOrders.frx":197B5
         textRB          =   "PhoneOrders.frx":197CD
         colorBack       =   "PhoneOrders.frx":197E5
         colorIntern     =   "PhoneOrders.frx":1980F
         colorMO         =   "PhoneOrders.frx":19839
         colorFocus      =   "PhoneOrders.frx":19863
         colorDisabled   =   "PhoneOrders.frx":1988D
         colorPressed    =   "PhoneOrders.frx":198B7
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":198E1
         Cell1           =   "PhoneOrders.frx":199F5
         Cell2           =   "PhoneOrders.frx":19ADD
         Cell3           =   "PhoneOrders.frx":19BC9
         Cell4           =   "PhoneOrders.frx":1A7BF
         Cell5           =   "PhoneOrders.frx":1B3B3
         Cell6           =   "PhoneOrders.frx":1B493
         Cell7           =   "PhoneOrders.frx":1B581
         Cell8           =   "PhoneOrders.frx":1B659
         Cell9           =   "PhoneOrders.frx":1B739
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   10
         Left            =   9240
         TabIndex        =   91
         Top             =   3445
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":1BE49
         textLT          =   "PhoneOrders.frx":1BE61
         textCT          =   "PhoneOrders.frx":1BE79
         textRT          =   "PhoneOrders.frx":1BE91
         textLM          =   "PhoneOrders.frx":1BEA9
         textRM          =   "PhoneOrders.frx":1BEC1
         textLB          =   "PhoneOrders.frx":1BED9
         textCB          =   "PhoneOrders.frx":1BEF1
         textRB          =   "PhoneOrders.frx":1BF09
         colorBack       =   "PhoneOrders.frx":1BF21
         colorIntern     =   "PhoneOrders.frx":1BF4B
         colorMO         =   "PhoneOrders.frx":1BF75
         colorFocus      =   "PhoneOrders.frx":1BF9F
         colorDisabled   =   "PhoneOrders.frx":1BFC9
         colorPressed    =   "PhoneOrders.frx":1BFF3
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":1C01D
         Cell1           =   "PhoneOrders.frx":1C131
         Cell2           =   "PhoneOrders.frx":1C219
         Cell3           =   "PhoneOrders.frx":1C305
         Cell4           =   "PhoneOrders.frx":1CEFB
         Cell5           =   "PhoneOrders.frx":1DAEF
         Cell6           =   "PhoneOrders.frx":1DBCF
         Cell7           =   "PhoneOrders.frx":1DCBD
         Cell8           =   "PhoneOrders.frx":1DD95
         Cell9           =   "PhoneOrders.frx":1DE75
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   11
         Left            =   9240
         TabIndex        =   92
         Top             =   6600
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":1E585
         textLT          =   "PhoneOrders.frx":1E59D
         textCT          =   "PhoneOrders.frx":1E5B5
         textRT          =   "PhoneOrders.frx":1E5CD
         textLM          =   "PhoneOrders.frx":1E5E5
         textRM          =   "PhoneOrders.frx":1E5FD
         textLB          =   "PhoneOrders.frx":1E615
         textCB          =   "PhoneOrders.frx":1E62D
         textRB          =   "PhoneOrders.frx":1E645
         colorBack       =   "PhoneOrders.frx":1E65D
         colorIntern     =   "PhoneOrders.frx":1E687
         colorMO         =   "PhoneOrders.frx":1E6B1
         colorFocus      =   "PhoneOrders.frx":1E6DB
         colorDisabled   =   "PhoneOrders.frx":1E705
         colorPressed    =   "PhoneOrders.frx":1E72F
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":1E759
         Cell1           =   "PhoneOrders.frx":1E86D
         Cell2           =   "PhoneOrders.frx":1E955
         Cell3           =   "PhoneOrders.frx":1EA41
         Cell4           =   "PhoneOrders.frx":1F637
         Cell5           =   "PhoneOrders.frx":2022B
         Cell6           =   "PhoneOrders.frx":2030B
         Cell7           =   "PhoneOrders.frx":203F9
         Cell8           =   "PhoneOrders.frx":204D1
         Cell9           =   "PhoneOrders.frx":205B1
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   12
         Left            =   12280
         TabIndex        =   93
         Top             =   345
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":20CC1
         textLT          =   "PhoneOrders.frx":20CD9
         textCT          =   "PhoneOrders.frx":20CF1
         textRT          =   "PhoneOrders.frx":20D09
         textLM          =   "PhoneOrders.frx":20D21
         textRM          =   "PhoneOrders.frx":20D39
         textLB          =   "PhoneOrders.frx":20D51
         textCB          =   "PhoneOrders.frx":20D69
         textRB          =   "PhoneOrders.frx":20D81
         colorBack       =   "PhoneOrders.frx":20D99
         colorIntern     =   "PhoneOrders.frx":20DC3
         colorMO         =   "PhoneOrders.frx":20DED
         colorFocus      =   "PhoneOrders.frx":20E17
         colorDisabled   =   "PhoneOrders.frx":20E41
         colorPressed    =   "PhoneOrders.frx":20E6B
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":20E95
         Cell1           =   "PhoneOrders.frx":20FA9
         Cell2           =   "PhoneOrders.frx":21091
         Cell3           =   "PhoneOrders.frx":2117D
         Cell4           =   "PhoneOrders.frx":21D73
         Cell5           =   "PhoneOrders.frx":22967
         Cell6           =   "PhoneOrders.frx":22A47
         Cell7           =   "PhoneOrders.frx":22B35
         Cell8           =   "PhoneOrders.frx":22C0D
         Cell9           =   "PhoneOrders.frx":22CED
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   13
         Left            =   12280
         TabIndex        =   94
         Top             =   3445
         Visible         =   0   'False
         Width           =   3060
         _Version        =   589828
         _ExtentX        =   5397
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":233FD
         textLT          =   "PhoneOrders.frx":23415
         textCT          =   "PhoneOrders.frx":2342D
         textRT          =   "PhoneOrders.frx":23445
         textLM          =   "PhoneOrders.frx":2345D
         textRM          =   "PhoneOrders.frx":23475
         textLB          =   "PhoneOrders.frx":2348D
         textCB          =   "PhoneOrders.frx":234A5
         textRB          =   "PhoneOrders.frx":234BD
         colorBack       =   "PhoneOrders.frx":234D5
         colorIntern     =   "PhoneOrders.frx":234FF
         colorMO         =   "PhoneOrders.frx":23529
         colorFocus      =   "PhoneOrders.frx":23553
         colorDisabled   =   "PhoneOrders.frx":2357D
         colorPressed    =   "PhoneOrders.frx":235A7
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":235D1
         Cell1           =   "PhoneOrders.frx":236E5
         Cell2           =   "PhoneOrders.frx":237CD
         Cell3           =   "PhoneOrders.frx":238B9
         Cell4           =   "PhoneOrders.frx":244AF
         Cell5           =   "PhoneOrders.frx":250A3
         Cell6           =   "PhoneOrders.frx":25183
         Cell7           =   "PhoneOrders.frx":25271
         Cell8           =   "PhoneOrders.frx":25349
         Cell9           =   "PhoneOrders.frx":25429
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   345
         Visible         =   0   'False
         Width           =   2940
         _Version        =   589828
         _ExtentX        =   5186
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":25B39
         textLT          =   "PhoneOrders.frx":25B51
         textCT          =   "PhoneOrders.frx":25B69
         textRT          =   "PhoneOrders.frx":25B81
         textLM          =   "PhoneOrders.frx":25B99
         textRM          =   "PhoneOrders.frx":25BB1
         textLB          =   "PhoneOrders.frx":25BC9
         textCB          =   "PhoneOrders.frx":25BE1
         textRB          =   "PhoneOrders.frx":25BF9
         colorBack       =   "PhoneOrders.frx":25C11
         colorIntern     =   "PhoneOrders.frx":25C3B
         colorMO         =   "PhoneOrders.frx":25C65
         colorFocus      =   "PhoneOrders.frx":25C8F
         colorDisabled   =   "PhoneOrders.frx":25CB9
         colorPressed    =   "PhoneOrders.frx":25CE3
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":25D0D
         Cell1           =   "PhoneOrders.frx":25E21
         Cell2           =   "PhoneOrders.frx":25F09
         Cell3           =   "PhoneOrders.frx":25FF5
         Cell4           =   "PhoneOrders.frx":26BEB
         Cell5           =   "PhoneOrders.frx":277DF
         Cell6           =   "PhoneOrders.frx":278BF
         Cell7           =   "PhoneOrders.frx":279AD
         Cell8           =   "PhoneOrders.frx":27A85
         Cell9           =   "PhoneOrders.frx":27B65
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
      Begin BTNENHLib4.BtnEnh SelectionButton 
         Height          =   3000
         Index           =   14
         Left            =   12240
         TabIndex        =   101
         Top             =   6600
         Visible         =   0   'False
         Width           =   3060
         _Version        =   589828
         _ExtentX        =   5397
         _ExtentY        =   5292
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CornerFactor    =   10
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         SpecialEffect   =   3
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "PhoneOrders.frx":28275
         textLT          =   "PhoneOrders.frx":2828D
         textCT          =   "PhoneOrders.frx":282A5
         textRT          =   "PhoneOrders.frx":282BD
         textLM          =   "PhoneOrders.frx":282D5
         textRM          =   "PhoneOrders.frx":282ED
         textLB          =   "PhoneOrders.frx":28305
         textCB          =   "PhoneOrders.frx":2831D
         textRB          =   "PhoneOrders.frx":28335
         colorBack       =   "PhoneOrders.frx":2834D
         colorIntern     =   "PhoneOrders.frx":28377
         colorMO         =   "PhoneOrders.frx":283A1
         colorFocus      =   "PhoneOrders.frx":283CB
         colorDisabled   =   "PhoneOrders.frx":283F5
         colorPressed    =   "PhoneOrders.frx":2841F
         CellsCount      =   10
         Cell0           =   "PhoneOrders.frx":28449
         Cell1           =   "PhoneOrders.frx":2855D
         Cell2           =   "PhoneOrders.frx":28645
         Cell3           =   "PhoneOrders.frx":28731
         Cell4           =   "PhoneOrders.frx":29327
         Cell5           =   "PhoneOrders.frx":29F1B
         Cell6           =   "PhoneOrders.frx":29FFB
         Cell7           =   "PhoneOrders.frx":2A0E9
         Cell8           =   "PhoneOrders.frx":2A1C1
         Cell9           =   "PhoneOrders.frx":2A2A1
         HollowFrame     =   -1  'True
         LightDirection  =   1
         VistaColorGlossyEffectLower=   13677994
      End
   End
   Begin VB.Label PendingOrders 
      BackStyle       =   0  'Transparent
      Caption         =   "Pending Orders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   96
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No Open Orders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   2880
      TabIndex        =   56
      Top             =   3240
      Width           =   10575
   End
End
Attribute VB_Name = "PhoneOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private DisplayOnly     As Boolean
'Private ScanCheck       As Long

Private DepositCheckNo          As Long
Private SelectedRow             As Integer
Private AllPhoneOrders          As AllCallins
Private TScroll                 As Integer
Private WorkingCheckNo          As Long
Private ChangeOrderDestination  As String
Private Enum DoInstruction
        MarkPaid
        EditCust
End Enum
Private MyInstruction   As DoInstruction
Public Sub ReviewOnly()
'DisplayOnly = True
SSCommand1.Enabled = False
SSCommand2.Enabled = False
SSCommand4.Enabled = False
SSCommand5.Enabled = False
SSCommand3.Enabled = False
SSCommand15.Enabled = False
SSCommand11.Enabled = False
SSCommand10.Enabled = False
BtnEnh13.Enabled = False
BtnEnh10.Enabled = False
BtnEnh4.Enabled = False
BtnEnh6.Enabled = False
BtnEnh1.Enabled = False
SearchHow(6).Enabled = False
SearchHow(7).Enabled = False
SearchHow(5).Enabled = False
End Sub
Public Sub DriverAccess()
BtnEnh13.Enabled = False
BtnEnh10.Enabled = False
BtnEnh4.Enabled = False
BtnEnh6.Enabled = False
BtnEnh1.Enabled = False
SSCommand1.Enabled = False
SSCommand2.Enabled = False
ShowEdit.Enabled = False
SSCommand4.Enabled = True
SSCommand5.Enabled = True
SSCommand3.Enabled = False
SSCommand15.Enabled = False
SSCommand11.Enabled = False
SSCommand10.Enabled = False
SearchHow(6).Enabled = False
SearchHow(7).Enabled = False
SearchHow(5).Enabled = True
SearchHow(5).DoClick '// only shows unassigned orders
End Sub

Private Sub BtnEnh1_Click()
 CurrentPhoneOrders.Caption = "Release Phone Order" & " Open Orders ( " & AllPhoneOrders.count & " )"
 DT_Order_Status = MyOrderStatus(ParkCheck)
 Check_Status = MyOrderStatus(ReleaseCallIn)
 Trans_Log "Release Order"
 BtnEnh9_Click
End Sub
Private Sub BtnEnh10_Click()
CenterObject Me, SearchBox
SearchBox.ZOrder 0
SearchBox.Visible = True
End Sub
Private Sub BtnEnh11_Click()
PrinterPhoneSummary
BtnEnh9_Click
End Sub
Private Sub BtnEnh12_Click()
PrinterPendingOrders
BtnEnh9_Click
End Sub
Private Sub BtnEnh13_Click()
If BtnEnh13.Caption = "Order for Dine In" Then
   CurrentPhoneOrders.Caption = "Transfer To Dining Room" & " Open Orders ( " & AllPhoneOrders.count & " )"
   DT_Order_Status = MyOrderStatus(ParkCheck)
   Check_Status = MyOrderStatus(TranferPhone)
   BtnEnh9_Click
Else
   SearchHow_Click 3
End If
End Sub
Private Sub BtnEnh14_Click(Index As Integer)
    On Error Resume Next
    If Load_CallIn_Order(WorkingCheckNo, CInfo(0).Text) = False Then Exit Sub
    Sales_Totals
    GetDeliveryPayment WorkingCheckNo
    Receipt "Phone In Order", PrintReceipt, "", False
    ReDim_Array
    Clear_Grid Order.Ticket(0)
    SSCommand6_Click
End Sub
Private Sub BtnEnh15_Click()
ShowOrders.Visible = False
ShowOrders.Left = -5000
End Sub
Private Sub BtnEnh16_Click()
MyInstruction = MarkPaid
SearchHow_Click 4
With EditCustomer
   .Clear
   .FixedRows = 1
   .TextMatrix(0, 0) = "Name"
   .TextMatrix(0, 1) = "Phone"
   .TextMatrix(0, 2) = "Check"
   .cols = 3
   .FontSize = 18
   .ColWidth(0) = 3500 * WidScale1024
   .ColWidth(1) = 2500 * WidScale1024
   .ColWidth(2) = 1500 * WidScale1024
   .ColAlignment(0) = flexAlignLeftCenter
   .ColAlignment(1) = flexAlignLeftCenter
   .ColAlignment(2) = flexAlignLeftCenter
   .width = ((.ColWidth(0) + .ColWidth(1)) + .ColWidth(2)) + 300
   .rows = AllPhoneOrders.count + 1
   ShowOrders.width = .width + 500
End With
Dim F As Integer
Dim C As Integer
C = 0
For F = 1 To AllPhoneOrders.count
   If AllPhoneOrders.Item(F).AssignmentTime <> "Open" And AllPhoneOrders.Item(F).OrderType = "Delivery" Then
    C = C + 1
    With EditCustomer
      .TextMatrix(C, 0) = UCase(AllPhoneOrders.Item(F).CallerName)
      .TextMatrix(C, 1) = AllPhoneOrders.Item(F).CallerPhone
      .TextMatrix(C, 2) = AllPhoneOrders.Item(F).PickUpNo
    End With
   End If
Next F
If C = 0 Then
   See_Msg "No Records To Mark", 5, ""
   Exit Sub
End If
CenterObject Me, ShowOrders
ShowOrders.Visible = True
ShowOrders.ZOrder 0
End Sub
Private Sub BtnEnh17_Click()
Dim searchby As String
searchby = ShowNumberPad("Enter Phone", False, 0, 0)
If searchby > "" Then
  If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
     Sql = "Select * from [Open Checks] Where [Customer Phone] Like '%" & searchby & "%' and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by [Customer phone] asc"
  Else
     Sql = "Select * from [Open Checks] Where [Customer Phone] Like '%" & searchby & "%' order by [Customer phone] asc"
  End If
End If
If Sql > "" Then
  Load_CallINs Sql, ""
  TScroll = 1
  ScrollButtons
End If
End Sub
Private Sub BtnEnh18_Click()
    Find_What = "Drivers"
    Load_What = "Set as Drivers"
    GetSelected = ""
    RequestListForm.Show 1
    If GetSelected = "" Then Exit Sub
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Sql = "Select * from [Check Numbers] where [Servers Name] = '" & GetSelected & "'"
    ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    ConnectADORS Sql, Db, RS
      If RS.EOF Then
        See_Msg GetSelected & " has no audit open", 5, "Cannot change status"
        DbClose Db, RS
        Exit Sub
      End If
      If Not RS.EOF Then
      RS![Work Area] = "Driver"
      RS.Update
      End If
      DbClose Db, RS
      See_Msg "Status has been changed to driver", 5, ""
End Sub
Private Sub BtnEnh2_Click()
AssignDrv.Visible = False
AssignDrv.Left = -5000
End Sub
Private Sub BtnEnh3_Click()
If Can_Continue(BtnEnh3.Caption) = False Then Exit Sub
If IsPrinted(DepositCheckNo) Then
  See_Msg "Selecting PrePay ", 5, "Check active in todays sales"
  BtnEnh5_Click
  Exit Sub
End If
Emv_Back.r_AuthAmt = 0
depositpaid = True
DepositAmtPaid = -1
PaidIn.Show 1
depositpaid = False
If DepositAmtPaid > 0 Then
 RecordDeposit DepositCheckNo, CStr(CInfo(0).Text), CStr(CInfo(5).Text), DepositAmtPaid
 CInfo(7).Text = Format(Val(CInfo(7).Text) + DepositAmtPaid, "0.00")
End If
DepositAmtPaid = 0
End Sub
Private Sub BtnEnh4_Click()
If See_Msg("Transfer To Dining Room ?", 4, "") <> 6 Then Exit Sub
Trans_Log "Transfered To Dining Room "
'// move the customer phone order to the dining room.

CurrentPhoneOrders.Caption = "Transfer To Dining Room" & " Open Orders ( " & AllPhoneOrders.count & " )"
DT_Order_Status = MyOrderStatus(ParkCheck)
Check_Status = MyOrderStatus(TranferPhone)
BtnEnh9_Click
End Sub
Private Sub BtnEnh5_Click()
Dim Crd    As Integer
'Dim TemPId As String
Dim OrderTaker As String
OrderTaker = LoggedOnAs(0).EMPName

If Can_Continue(BtnEnh5.Caption) = False Then Exit Sub
If IsPrinted(DepositCheckNo) = False Then
    See_Msg "Use Deposit", 0, "PrePaid Applies To Todays Orders Only"
    Exit Sub
End If

If PaidByCredit(DepositCheckNo).IsPaid Then
   If See_Msg("Paid By Credit Card", 4, "Void Payment??") = 6 Then
     If Delete_Payment_Record(DepositCheckNo, "Phone Order", False) = False Then
      See_Msg "Cannot Void Transaction", 5, "Check Not Voided."
      Exit Sub
     Else
       PaymentAmount.Text = "0.00"
     End If
   Else
     Exit Sub
   End If
End If
  Load_What = "PrePaidDeposit"
  Select_What.Show 1
  If GetSelected = "" Then Exit Sub
For Crd = 0 To UBound(CreditCards, 1) - 1
   If UCase(CreditCards(Crd).CrdName) = UCase(GetSelected) Then
      'TemPId = LoggedOnAs(0).EMPName
      Check_no = DepositCheckNo
      
      If Check_Option(151) = 1 Then
        PreAuthAmount = Val(ShowNumberPad("Enter Amount", False, 0, 0))
        If PreAuthAmount = 0 Then Exit Sub
        If PreAuthAmount > Val(CheckAMT.Text) Then
          PreAuthAmount = Val(CheckAMT.Text)
          See_Msg "Amount greater than check", 5, "new amount is " & Format(PreAuthAmount, "0.00")
        End If
      Else
        PreAuthAmount = Val(CheckAMT.Text)
      End If
      Chk_Sales_data(5) = PreAuthAmount
      If CreditCards(Crd).ReqAuth = TenderTypes.EVM Then
         CurrentStatus = Check_Status
         Check_Status = MyOrderStatus(PreAuthCheck)
         
         If OutOfScope(OrderTaker, "EMV_Sale", Format(Val(CheckAMT.Text), "0.00"), OOCCreditSale) = "Exit" Then
          Exit Sub
         End If
        
         '// stay logged in as phone order for declines
         If IsApproved(Emv_Back.r_Status) Then
          If Emv_Back.r_AuthAmt < ChargeAmount Then
            See_Msg "Partial Approval", 5, "Balance Due " & Format(ChargeAmount - Emv_Back.r_AuthAmt, "0.00")
          Else
            See_Msg "Approved...", 5, ""
          End If
           MarkPhoneOrderPaid Check_no, Emv_Back.r_AuthAmt, Format(Val(CheckAMT.Text), "0.00")
           PaymentAmount.Text = Format(Emv_Back.r_AuthAmt, "0.00")
          If IsDataCapInterface(EMVINTERFACE) Then
            PrintChitSlip FormatEMVReceipt(Emv_Back.r_ReturnResult)
          Else
             Dim CCD As CCData
                With CCD
                 .HeaderData = "No"
                 .HeaderReason = "EMV-"
                 .AuthNo = Emv_Back.r_AuthCode
                 .SaleAmt = Format(Emv_Back.r_AuthAmt, "0.00")
                 .TipAmt = Format(Emv_Back.r_TipAmt, "0.00")
                 .AccNo = Emv_Back.r_Last4
                 .CustName = ""
                 .CType = Emv_Back.r_Cardtype
                 .ServerName = Emv_Back.r_Server
                 .TBLName = Table_Number
                 .ChkNO = CStr(Check_no)
                 .ProcessorResponse = "Approved"
                 .Datetime = Format(Time, "HH:MM ampm")
                 .ckTotal = Format(Chk_Sales_data(5), "0.00")
                 .ExpDate = ""
                End With
                PrintSlip CCD
          End If
          SSCommand6_Click
         ElseIf Emv_Back.r_Status = "DECLINED" Then
            See_Msg "Declined...", 0, Emv_Back.r_Message
            If IsDataCapInterface(EMVINTERFACE) Then
               PrintChitSlip FormatEMVReceipt(Emv_Back.r_ReturnResult)
            End If
            Check_Status = CurrentStatus
            'LoggedOnAs(0).EMPName = TemPId
          Exit Sub
         Else
          See_Msg Emv_Back.r_Message, 0, ""
          Check_Status = CurrentStatus
          'LoggedOnAs(0).EMPName = TemPId
          Exit Sub
         End If
         If Emv_Back.r_AuthAmt < Chk_Sales_data(5) And Emv_Back.r_TipAmt > 0 Then
        
         Else
          ChargeAmount = Format(Val(Emv_Back.r_AuthAmt), "0.00")   '  ChargeAmount + Format(Val(xmls.token("CashBack")), "0.00") ' Emv_Trans.CashBack  'Val(xmls.token("CashBack"))
         End If
         Emv_Back.r_TipAmt = 0
         Check_Status = CurrentStatus
      Else
         
         UpdatePhonePrePayment GetSelected, OrderTaker, DepositCheckNo, Format(Chk_Sales_data(5), "0.00")
         See_Msg "Payment Updated...", 5, ""
         PaymentAmount.Text = Format(Chk_Sales_data(5), "0.00")
     End If
    End If
  Next Crd
  'LoggedOnAs(0).EMPName = TemPId
End Sub
Private Sub BtnEnh6_Click()

If Can_Continue(BtnEnh6.Caption) = False Then Exit Sub
   Load_What = "Drivers"
   Find_What = "Select a Driver"
   RequestListForm.Show 1
   If Trim$(GetSelected) = "" Then Exit Sub
   If Trim$(GetSelected) > "" Then PrintOrderSlips GetSelected
End Sub
Private Sub PrintOrderSlips(DriverName As String)
Dim DTHoldStatus As Integer
Dim CSHoldStatus As Integer
DTHoldStatus = DT_Order_Status
CSHoldStatus = Check_Status
If LoggedOnAs(0).EmpAccessLevel >= Access_Required("Access Phone Orders") Then
  If See_Msg("Print Delivery Check for " & DriverName & " ?", 4, "No to quit") = 6 Then
     Sql = "Select * from [Open Checks] where [Driver Assigned] = '" & DriverName & "'"
  Else
    Exit Sub
    ' Sql = "Select * from [Open Checks] "
    '   Sql = Sql & " where TRY_CAST([Pick Up Date] AS date) = '" & Date & "'"
  End If
Else
  If See_Msg("Print Delivery Check ?", 4, "") = 6 Then
    Sql = "Select * from [Open Checks] where [Driver Assigned] = '" & DriverName & "'"
     Sql = Sql & " and TRY_CAST([Pick Up Date] AS date) = '" & Date & "'"
  Else
    Exit Sub
  End If
End If
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
 If Not Tb Is Nothing Then
    If Not Tb.BOF Then Tb.MoveFirst
 Else
    See_Msg "error connecting", 5, ""
    Exit Sub
 End If
 If Tb.EOF Then
  DbClose Nothing, Tb
  Set GRS = Nothing
  See_Msg "No Order Found", 5, ""
  Exit Sub
 End If
 CheckBelongsTo = "": DT_Order_Status = MyOrderStatus(ParkCheck): Check_Status = MyOrderStatus(CallIncheck)
 Do While Not Tb.EOF
   CheckBelongsTo = Tb![Server Name]
   If Load_CallIn_Order(Tb![Check Number], Tb![Customer Name] & "") = False Then Exit Sub
   Sales_Totals
   GetDeliveryPayment Tb![Check Number]
   Receipt "Delivery ", PrintReceipt, "Reprint", False
   ReDim_Array
   Clear_Grid Order.Ticket(0)
   Table_Number = "": Sold_To = ""
  Tb.MoveNext
 Loop
DbClose Nothing, Tb: Set GRS = Nothing: CheckBelongsTo = ""
DT_Order_Status = DTHoldStatus: Check_Status = CSHoldStatus
End Sub
Private Sub BtnEnh7_Click()
If Can_Continue("Void Check") = False Then Exit Sub
With PhoneList
  Cancel_CallIn_Order CStr(DepositCheckNo), ""
End With
SSCommand6_Click
SSCommand15_Click
End Sub
Private Sub BtnEnh8_Click()
  Dim ScanCheck As Long
  ScanCheck = Val(ShowNumberPad("Enter Check Number", True, 0, 0))
  If ScanCheck > 0 Then
     Sql = "Select * from [Open Checks] where [Check Number] = " & ScanCheck
     Load_CallINs Sql, ""
     TScroll = 1
     ScrollButtons
  End If

End Sub
Private Sub BtnEnh9_Click()
SearchBox.Visible = False
SearchBox.Left = -5000
End Sub
Private Sub CData_Click(Index As Integer)
On Error GoTo ERH
If Index = 1 Then
 If SiteLic.Google = 0 Then
   CAddress.Text = ShowKeyBoard("Enter Address")
 Else
  Dim SA As String
  Dim GetDir As MSaleDirections.GetDirections
  Set GetDir = New MSaleDirections.GetDirections
  SA = GetDir.FindAStreetAddress
  Set GetDir = Nothing
  If InStr(SA, ",") > 0 Then
    Dim StInfo() As String
    StInfo = Split(SA, ",")
    CInfo(1).Text = Replace(StInfo(0), "'", "`")
    CInfo(2).Text = StInfo(1)
    CInfo(3).Text = StInfo(2)
    CInfo(4).Text = StInfo(4)
  End If
 End If
End If
Exit Sub
ERH:
ErrorHandler:
Error_Log Err.Number, Err.Description, "Enter Address", "None"
End Sub

Private Sub CheckAMT_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub CheckAMT_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CheckAMT_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub CInfo_Click(Index As Integer)
Dim VKeyBoard   As String
Dim Directions  As String
Select Case Index
   Case Is = 0, 1, 2, 3, 4, 5
     VKeyBoard = ShowKeyBoard("Enter Info")
End Select
If VKeyBoard > "" Then CInfo(Index).Text = VKeyBoard
If Index < 6 Then Exit Sub
On Error Resume Next
If SiteLic.Google = 0 Then
 See_Msg "No License for Directions", 5, "Please Contact Your Dealer For Help"
 Exit Sub
End If
If Check_Option(59) = 0 Then
  See_Msg "Option Disabled", 0, ""
  'Exit Sub
End If
Dim DriveInfo   As DrvZoneData
DriveInfo = GetDelivery(CAddress.Text, CCity.Text, CState.Text, CZip.Text)
ChDir App.path
End Sub
Private Function IsPrinted(PhoneOrderNo As Long) As Boolean
On Error Resume Next
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
     Sql = "Select * from [Open Checks] where [Presented] = 1 and [Check Number] = " & PhoneOrderNo & " or TRY_CAST([pick up date] AS date) = '" & CDate(Date) & "' and [Check Number] = " & PhoneOrderNo
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
     If Not Tb Is Nothing Then
        If Not Tb.BOF Then Tb.MoveFirst
     Else
        See_Msg "error connecting", 5, ""
        Exit Function
     End If
     If Not Tb.EOF Then
        IsPrinted = True
     Else
        IsPrinted = False
     End If
DbClose Nothing, Tb: Set GRS = Nothing
End Function

Private Sub CInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Or Index = 5 Or Index = 7 Then
  KeyCode = 0
End If

End Sub

Private Sub CInfo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Or Index = 5 Or Index = 7 Then
  See_Msg "Cannot Edit At This Time", 5, ""
  KeyAscii = 0
  Exit Sub
End If
End Sub
Private Sub LoadCustomerToEdit(CNo As String)
MyInstruction = EditCust
With EditCustomer
   .Clear
   .FixedRows = 1
End With
    With EditCustomer
      .TextMatrix(1, 0) = UCase(AllPhoneOrders.Item(CNo).CallerName)
      .TextMatrix(1, 1) = AllPhoneOrders.Item(CNo).CallerPhone
      .TextMatrix(1, 2) = AllPhoneOrders.Item(CNo).PickUpNo
    End With
End Sub

Private Sub DriverName_Click()
Dim AssigmentName As String
AssigmentName = DriverName.List(DriverName.ListIndex)
AssignToADriver AssigmentName
Exit Sub
End Sub
Private Sub AssignToADriver(SelectedDriverName As String)
On Error GoTo ErrorHandler
Dim Db           As ADODB.Connection
Dim Tb           As ADODB.Recordset
Dim Cdb          As ADODB.Connection
Dim CRS          As ADODB.Recordset
Dim LogOn        As String
LogOn = LoggedOnAs(0).EMPName
Dim PhoneOrderNo As Long
OrderIsFor.Caption = ""
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
ChangeOrderDestination = ""
PhoneOrderNo = AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).PickUpNo
If SelectedDriverName = "Change to Pick Up" Then
   Dim UserMsg As String
   If EMVSUPPORTPREAUTH = 0 Or PAXBARTABPREAUTH = 0 Then
     UserMsg = "Payments and delivery charges will be voided"
   Else
     UserMsg = "delivery charges will be voided"
   End If
   If See_Msg("Change To Pickup Order?", 4, UserMsg) <> 6 Then
     DbClose Db, Tb
     Exit Sub
   End If
   Sql = "Select * from [Open Checks] where [Check Number] = " & PhoneOrderNo
   ConnectADORS Sql, Db, Tb
   If Not Tb.EOF Then
      Tb![Time assigned] = ""
      Tb![Driver Assigned] = ""
      Tb![Order Destination] = "Pick Up"
      Tb.Update
   End If
   OrderIsFor.Caption = "Pick Up"
   SelectionButton(SelectedRow).CellsManager.Text(6) = "Pick Up"
   AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType = "Pick Up"
   SetNothing Db
   Call VoidDeliveryCharge(PhoneOrderNo)
   See_Msg "Delivery Charge Voided", 5, ""
   If EMVSUPPORTPREAUTH = 0 Or PAXBARTABPREAUTH = 0 Then
    If Delete_Payment_Record(PhoneOrderNo, "", True) Then
     See_Msg "Payment Voided", 5, ""
      ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
      Sql = "Update [Open Checks] set [payments] = 0 where [Check Number] = " & PhoneOrderNo
      Db.Execute Sql
      SetNothing Db
      AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderPaymentAmt = 0
      AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderPayment = ""
    End If
   End If
   Load_CallINs "", ""
   ScrollButtons
   DbClose Db, Tb
   BtnEnh2_Click
   Exit Sub
ElseIf SelectedDriverName = "Unassign Order" Then
  Sql = "Select * from [Open Checks] where [Check Number] = " & PhoneOrderNo
 ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   Tb![Time assigned] = ""
   Tb![Driver Assigned] = ""
   Tb![Order Destination] = "Delivery"
   AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).PTime = "Open"
   SelectionButton(SelectedRow).CellsManager.Text(7) = "Open"
   AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType = "Delivery"
   SelectionButton(SelectedRow).CellsManager.Text(6) = "Delivery"
  Tb.Update
 End If
 SetNothing Tb
 BtnEnh2_Click
ElseIf SelectedDriverName = "Change To Delivery" Then
   If See_Msg("Change To delivery Order?", 4, "Payments and delivery charges will be added") <> 6 Then
    DbClose Db, Tb
    Exit Sub
   End If
   If Check_Option(179) = 1 Then
      AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress = "Enter Address"
   End If
   Dim DriveInfo   As DrvZoneData
   If AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress <> "Enter Address" Then
    DriveInfo = GetDelivery(AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).City, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).State, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).Zip)
    If DriveInfo.ZoneName = "Out of Area" Then
     See_Msg "Out of Area no delivery available", 5, ""
     OrderIsFor.Caption = "Pick Up"
     DbClose Db, Tb
     Exit Sub
    End If
   End If
   Trans_Log "Changed To Delivery Check Number " & PhoneOrderNo
     Sql = "Select * from [Open Checks] where [Check Number] = " & PhoneOrderNo
     ConnectADORS Sql, Db, Tb
     If Not Tb.EOF Then
       OrderIsFor.Caption = "Delivery"
       Tb![Pick Up Time] = Format(Time, "HH:MM ampm")
       Tb![Time assigned] = "" 'Format(Time, "HH:MM ampm")
       Tb![Order Destination] = "Delivery"
       Tb![Driver Assigned] = "" 'SelectedDriverName & ""
       Tb.Update
     End If
     Tb.Close
     AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType = "Delivery"
     SelectionButton(SelectedRow).CellsManager.Text(6) = "Delivery"
     If AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress = "Enter Address" Then
       See_Msg "Enter Delivery Address", 5, "Address Required"
       LoadCustomerToEdit SelectionButton(SelectedRow).Tag
       DepositCheckNo = PhoneOrderNo
       LoadCustomer AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).CallerName, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).CallerPhone, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).City
       PUTime.Text = AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).PTime
       If AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).StatusColor = &H80000012 Then
         PUDATE.Text = AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).PDate
         PUDATE.Enabled = False
         PickUpInfo.Caption = "Order has already been processed"
       Else
         PUDATE.Text = AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).PDate
         PUDATE.Enabled = True
         PickUpInfo.Caption = "Click On Date or Time To Change"
       End If
       If GetConfigInfo("Phone Orders", "No Restriction For Date Change", "Off", "FUNCTIONS.INI") = "ENABLED" Then
         PUDATE.Enabled = True
         PickUpInfo.Caption = "Click On Date or Time To Change"
       End If
       OrderIsFor.Caption = AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType
       Do While CustomerINfo.Visible
        DoEvents
        Sleep 500
       Loop
       If CInfo(1).Text & "" = "" Then
        AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType = "Pick Up"
        SelectionButton(SelectedRow).CellsManager.Text(6) = "Pick Up"
        DbClose Db, Tb
        Exit Sub
       End If
       Sql = "Select * from [Open Checks] where [Check Number] = " & PhoneOrderNo
       ConnectADORS Sql, Db, Tb
       If Not Tb.EOF Then
        Tb![Pick Up Time] = Format(Time, "HH:MM ampm")
        Tb![Time assigned] = "" 'Format(Time, "HH:MM ampm")
        Tb![Order Destination] = "Delivery"
        Tb![Driver Assigned] = "" 'SelectedDriverName & ""
        Tb![Street] = CInfo(1).Text & ""
        Tb.Update
       End If
      End If
      AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress = CInfo(1).Text & ""
      DriveInfo = GetDelivery(AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).callerAddress, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).City, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).State, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).Zip)
      If DriveInfo.ZoneName = "Out of Area" Then
        Call See_Msg("Out of Area no delivery available", 5, "")
        OrderIsFor.Caption = "Pick Up"
        Tb![Order Destination] = "Pick Up"
        Tb![Driver Assigned] = ""
        Tb![Time assigned] = ""
        Tb.Update
        DbClose Db, Tb
        SSCommand1_Click
        Exit Sub
       'End If
      End If
   DbClose Db, Tb
   ChangeOrderDestination = "AddDeliveryCharge"
   SelectionButton_CellClick SelectedRow, 5
   BtnEnh2_Click
   SSCommand15_Click
   Exit Sub
Else  ' SelectedDriverName <> "Change To Delivery" Then
   If AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType = "Pick Up" Then
      See_Msg "Pick Up order assignment not required", 5, ""
      Exit Sub
   End If
   CheckBelongsTo = SelectedDriverName
   Sql = "Select * from [Check Detail] where [Check Number] = " & PhoneOrderNo
   ConnectADORS Sql, Db, Tb
   Do While Not Tb.EOF
    Tb![Servers Name] = SelectedDriverName & ""
    Tb![Cashier Name] = SelectedDriverName & ""
    If UCase(Tb.fields("Department Report").Value) = "ONLINE TIP" Then  '// OnLine Tip is a non sale revenue defined in the pos software for on line orders
     Tb.fields("Recorded Time").Value = Format(Time, "HH:MM ampm")
     Tb.fields("Deleted").Value = 1
     Tb.fields("Department Report").Value = "Online Tip"
    End If
    Tb.Update
    Tb.MoveNext
   Loop
   '// make sure drive get tip reported -- Pre order in the future will drop the driver tip
   '// tips get picked up based on the audit start.
   SetNothing Tb
   AssignToDriver SelectedDriverName, PhoneOrderNo
   If Check_Option(62) = 1 Then
     Call SendTextMessage("Confirmation: Your Order Is On The Way " & LocationName, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).CallerPhone)
   End If
  Sql = "Select * from [Open Checks] where [Check Number] = " & PhoneOrderNo
  ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     Tb![Server Name] = SelectedDriverName & ""
     Tb![Pick Up Time] = Format(Time, "HH:MM ampm")
     Tb![Time assigned] = Format(Time, "HH:MM ampm")
     Tb![Order Destination] = "Delivery"
     Tb![Driver Assigned] = SelectedDriverName & ""
     AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).PTime = Format(Time, "h:m ampm")
     SelectionButton(SelectedRow).CellsManager.Text(7) = Format(Time, "h:m ampm")
     AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType = "Delivery"
     SelectionButton(SelectedRow).CellsManager.Text(6) = "Delivery"
    Tb.Update
  End If
  DbClose Db, Tb
  If Check_Option(253) = 1 Then
    PrintCreditSlip PhoneOrderNo
  End If
  'Load_CallINs "", ""
  ' ScrollButtons  '// dont use here print wrong receipt to customer
  BtnEnh2_Click
End If
 Dim DTHoldStatus As Integer
 Dim CSHoldStatus As Integer
 DTHoldStatus = DT_Order_Status
 CSHoldStatus = Check_Status
 Check_Status = MyOrderStatus(CallIncheck)
 If PhoneOrders.Load_CallIn_Order(PhoneOrderNo, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).CallerName) = False Then Exit Sub
 'LogOn = LoggedOnAs(0).EMPName
 LoggedOnAs(0).EMPName = SelectedDriverName
 Order.Load_Grid
 Order.Amount_Tend_Update
 Receipt AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).OrderType, PrintReceipt, "", False
 LoggedOnAs(0).EMPName = LogOn
 Start_Check
 Order.Gen_Setup
 DT_Order_Status = DTHoldStatus
 Check_Status = CSHoldStatus
 Load_CallINs "", ""
 ScrollButtons
 UnlockedCheckNumber PhoneOrderNo
 SearchHow(5).DoClick
Exit Sub
ErrorHandler:
Error_Message "Driver Name", Err.Description
'esume
Resume Next
End Sub
Private Sub Dwn_arrow_Click()
If HowmanyButtons <= 15 Then Exit Sub
TScroll = TScroll + 15
ScrollButtons
End Sub
Private Sub EditCustomer_Click()
Dim C As Integer
On Error GoTo ERH
If AllPhoneOrders Is Nothing Then
   Exit Sub
End If

If MyInstruction = MarkPaid Then
  C = EditCustomer.Row
  If Not IsDate(AllPhoneOrders.Item(EditCustomer.TextMatrix(C, 2)).DriveReturned) Then
     MarkClosed AllPhoneOrders.Item(EditCustomer.TextMatrix(C, 2)).PickUpNo
     AllPhoneOrders.Item(EditCustomer.TextMatrix(C, 2)).DriveReturned = Now
     WorkingCheckNo = AllPhoneOrders.Item(EditCustomer.TextMatrix(C, 2)).PickUpNo
     See_Msg "Marked Closed", 5, ""
  End If
Else
    C = EditCustomer.Row
    If C = 0 Then Exit Sub
    SelectedRow = EditCustomer.Row
    WorkingCheckNo = AllPhoneOrders.Item(EditCustomer.TextMatrix(C, 2)).PickUpNo
    AttachPayment AllPhoneOrders.Item(EditCustomer.TextMatrix(SelectedRow, 2)).PickUpNo
End If
BtnEnh15_Click
Exit Sub
ERH:
Err.Clear
End Sub
Private Function AttachPayment(MySelection As Integer) As String
    'SelectedRow = MySelection
    Dim FC As Integer
    For FC = 1 To AllPhoneOrders.count
      If AllPhoneOrders.Item(FC).PickUpNo = MySelection Then
       DepositCheckNo = AllPhoneOrders.Item(FC).PickUpNo
       LoadCustomer AllPhoneOrders.Item(FC).CallerName, AllPhoneOrders.Item(FC).CallerPhone, AllPhoneOrders.Item(FC).callerAddress, AllPhoneOrders.Item(FC).City
       PUTime.Text = AllPhoneOrders.Item(FC).PTime
       If AllPhoneOrders.Item(FC).StatusColor = &H80000012 Then
         PUDATE.Text = AllPhoneOrders.Item(FC).PDate
         PUDATE.Enabled = False
         PickUpInfo.Caption = "Order has already been processed"
       Else
         PUDATE.Text = AllPhoneOrders.Item(FC).PDate
          PickUpInfo.Caption = "Click On Date or Time To Change"
         PUDATE.Enabled = True
       End If
       If GetConfigInfo("Phone Orders", "No Restriction For Date Change", "Off", "FUNCTIONS.INI") = "ENABLED" Then
         PUDATE.Enabled = True
         PickUpInfo.Caption = "Click On Date or Time To Change"
       End If
       OrderIs = AllPhoneOrders.Item(FC).OrderType
       OrderIsFor.Caption = OrderIs
       PaymentAmount.Text = PhoneOrderIsPaid(DepositCheckNo) ' PaidByCredit(DepositCheckNo).Payment
       If AllPhoneOrders.Item(FC).StatusColor = 255 Or AllPhoneOrders.Item(FC).StatusColor = 12870144 Then
        BtnEnh7.Visible = True
       Else
        BtnEnh7.Visible = False
       End If
       WorkingCheck.Caption = MySelection
       SelectedRow = FC
     End If
   Next FC
End Function
Private Sub Exit_Click()
If Check_Status = MyOrderStatus(NewCheck) Or Check_Status = MyOrderStatus(EditCheck) Then
  'do not change resets check_status
Else
 If ButtonFunction = BTNFNCT.LoadCallIn Or ButtonFunction = BTNFNCT.LoadPending Then
  Start_Check
  ResetCaller
  Order.Gen_Setup
  Redim_Grid
  DT_Order_Status = MyOrderStatus(EditCheck)
  Check_Status = MyOrderStatus(HereCheck)
  Order.Order_Destination
End If
End If
AbortSelection = True
If Check_Option(0) = 0 Then
   Check_no = Assign_Check_No(LoggedOnAs(0).EMPName, True)
Else
   Check_no = Advance_Check(False)
End If
Call MemoryCheck(PhoneOrders)
End Sub
Private Sub FilterBy_Click()

If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
  Sql = "Select * from [Open Checks] Where [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by [Order Destination], TRY_CAST([Order Date] AS datetime), TRY_CAST([check Opened] AS datetime) desc"
Else
  Sql = "Select * from [Open Checks] order by [Order Destination] , TRY_CAST([Order Date] AS datetime), TRY_CAST([check Opened] AS datetime) desc"
End If
Load_CallINs Sql, ""
TScroll = 1
ScrollButtons
CurrentPhoneOrders.Caption = "All Orders by Open Time" & " Open Orders ( " & AllPhoneOrders.count & " )"
BtnEnh9_Click
End Sub
Public Sub PrintCreditSlip(Chk As Long)
On Error GoTo ERH
Dim RS As ADODB.Recordset
Dim CCD As CCData
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
  Sql = "Select * from [Debit] Where [Close Out day] = 0 and [check number] = '" & CStr(Chk) & "'"
  Set RS = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & DebitDb), "Debit")
  If Not RS Is Nothing Then
     If Not RS.BOF Then RS.MoveFirst
  Else
     See_Msg "error connecting", 5, "Debit Database"
     Exit Sub
  End If
  Do While Not RS.EOF
     With CCD
     .HeaderData = "No"
     .HeaderReason = "Sale"
     .AuthNo = RS![Account Number] & ""
     .SaleAmt = Format(RS![Amount], "0.00")
     .TipAmt = Format(RS![tip], "0.00")
     .AccNo = RS![Card Number]
     .CustName = ""
     .CType = RS![Account Name]
     .ServerName = RS![Server Name] & ""
     .TBLName = ""
     .ChkNO = CStr(Chk)
     .ProcessorResponse = ""
     .Datetime = Format(Time, "HH:MM ampm")
     .ckTotal = Format(RS![Amount], "0.00")
     .ExpDate = "xxxx"
     End With
     PrintSlip CCD
     RS.MoveNext
  Loop
DbClose Nothing, RS: Set GRS = Nothing
Exit Sub
ERH:
Error_Log Err.Number, "Print Credit Card Slip", "", ""
Err.Clear
End Sub
Private Sub AssignDriver(OrderType As String)
    FilterDriver = False
    Dim GRS As GeneralLib.Utilities
    Set GRS = New GeneralLib.Utilities
    Dim Id As ADODB.Recordset
    DriverName.Clear
    If LoggedOnAs(0).EmpAccessLevel >= Access_Required("Close Delivery Order") Then
      If OrderType = "Delivery" Then
         DriverName.AddItem "Change to Pick Up"
         DriverName.AddItem "Unassign Order"
      End If
    Else
     See_Msg "No access to close delivery orders", 5, ""
     Exit Sub
    End If
    If OrderType = "Pick Up" Then DriverName.AddItem "Change To Delivery"

    If CPath.TerminalFailed Then
      See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
      Exit Sub
    End If
    If System_config = ExpressService Then
      Sql = "Select Distinct [employee Name] from [Employee records] where [Department Name] = 'Delivery' and [time in] > '' and [time out] = '' or [Department Name] = 'Driver' and [time in] > '' and [time out] = ''"
      Set Id = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & TimeRecordsDb), "Employee Records")
      If Not Id Is Nothing Then
        If Not Id.BOF Then Id.MoveFirst
      Else
        See_Msg "error connecting", 5, ""
        Exit Sub
      End If
      If Id.EOF Then
        See_Msg "No Driver Or Delivery Job Codes Found", 5, "Job Code Required."
      Else
        Do While Not Id.EOF
          DriverName.AddItem Id![employee name] & ""
          Id.MoveNext
        Loop
      End If
    Else
      Sql = "select * from [Check Numbers] where [audit status] = 0"
      If GetConfigInfo("Phone Orders", "Drivers Assignment Only Is", "Off", "FUNCTIONS.INI") = "ENABLED" Then
         Sql = Sql & " and [Work Area] = 'Driver'"
      End If
      If LoggedOnAs(0).EmpAccessLevel < Access_Required("Assign Driver Access") Then
       Sql = Sql & " and [Servers Name] = '" & LoggedOnAs(0).EMPName & "'"
      End If
      Set Id = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & "MSaleData"), "Check Numbers")
      If Not Id Is Nothing Then
        If Not Id.BOF Then Id.MoveFirst
      Else
        See_Msg "error connecting", 5, ""
        Exit Sub
      End If
      If Id.EOF Then
         See_Msg "No Driver Audit(s) Found", 5, "Audit Required."
      Else
       Do While Not Id.EOF
        DriverName.AddItem Id![Servers Name]
        Id.MoveNext
       Loop
      End If
    End If
DbClose Nothing, Id: Set GRS = Nothing
If DriverName.ListCount = 0 Then Exit Sub
With AssignDrv
    CenterObject Me, AssignDrv
    .ZOrder 0
    .Visible = True
End With
End Sub
Private Sub MarkClosed(DrvCheck As Long)
On Error Resume Next
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
     Sql = "Update [Open Checks] Set [Driver Return] = '" & Format(Time, "HH:MM ampm") & "' where [Check Number] = " & DrvCheck
     Call GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb))
Set GRS = Nothing
End Sub
Private Sub AssignToServer(SvrName As String, ReservationNo As Long)
On Error Resume Next
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities

   Sql = "Update [Check Detail] set [Servers Name] = '" & SvrName & "' where [Check Number] = " & ReservationNo
   Call GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb))
       
   Sql = "Update [Open Checks] set [Server Name] = '" & SvrName & "',[Pick Up Time] = '" & Format(Time, "HH:MM ampm") & "'"
   Sql = Sql & ", [Time Assigned] =" & Format(Time, "HH:MM ampm") & "', [Driver Assigned] = '' Where [Check Number] = " & ReservationNo
   Call GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb))
Set GRS = Nothing
End Sub
Private Sub AssignToDriver(DriverName As String, PHCheckNo As Long)
On Error GoTo ermsg
Dim Db          As ADODB.Connection
Dim CurrentTime As String
CurrentTime = Time
ConnectADODB SvrPath.TerminalPath & DebitDb, Db
Sql = "Update [Debit] Set [Time] = '" & CurrentTime & "', [Server Name] = '" & DriverName & "' Where [Check Number] = " & Val(PHCheckNo) & " and [Close Out Day] = 0"
Db.Execute Sql
Db.Close
ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
Sql = "Update [DayPayments] Set [Recorded Time] = '" & CurrentTime & "' Where [Check Number] = " & Val(PHCheckNo)
Db.Execute Sql
Db.Close
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
Sql = "Update [Payments] Set [Time] = '" & CurrentTime & "' Where [Check Number] = " & Val(PHCheckNo)
Db.Execute Sql
Db.Close
'// all times have to be the same for finding the original payment
Exit Sub
ermsg:
'// time has to be updated for driver assignment as not to drop tips for driver payout
End Sub
Private Sub SavePhoneUpdate(PHCheckNo As String, OriginalDate As String)
On Error GoTo ermsg
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim DepositAmt  As Currency
Dim CallTime    As SendCallin
Set CallTime = New SendCallin
DepositAmt = 0
If CDate(IsPickUpDate(AllPhoneOrders.Item(PHCheckNo).PDate, PDate)) > Date Then
 ConnectADODB SvrPath.TerminalPath & DebitDb, Db
 Sql = "Select * from [Debit] Where [Check Number] = " & Val(PHCheckNo) & " and [Close Out Day] = 0"
 If ConnectADORS(Sql, Db, Tb) = False Then
   See_Msg "Error connecting to credit card database", 5, ""
 Else
   If Not Tb.EOF Then
     DepositAmt = Tb.fields("Amount").Value + Tb.fields("Tip").Value
     Tb.fields("Check Number").Value = -9999
     Tb.fields("Note").Value = "Moved From " & PHCheckNo
     Tb.Update
   End If
 End If
 DbClose Db, Tb
End If

ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
Sql = "Select * from [Open Checks] Where [Customer Name] = '" & AllPhoneOrders.Item(PHCheckNo).CallerName & "'"
Sql = Sql & " and [Customer Phone] = '" & AllPhoneOrders.Item(PHCheckNo).CallerPhone & "'"
Sql = Sql & " and [Check Number] = " & PHCheckNo
If IsDate(OriginalDate) Then
  Sql = Sql & " and TRY_CAST([Pick Up Date] AS date) = '" & Format(OriginalDate, "MM/DD/YYYY") & "'"
End If
ConnectADORS Sql, Db, Tb
If Not Tb.EOF Then
   Tb![Pick Up Date] = IsPickUpDate(AllPhoneOrders.Item(PHCheckNo).PDate, PDate)
   Tb![Pick Up Time] = IsPickUpDate(AllPhoneOrders.Item(PHCheckNo).PTime, PTime)
   Tb.Update
   If CallTime.CheckTime(Format(AllPhoneOrders.Item(PHCheckNo).PTime, "HH:MM ampm"), AllPhoneOrders.Item(PHCheckNo).PDate, Val(Tb.fields("Production").Value & "")) = False Then
      If Tb![Presented] = 1 Then
        Sql = "Update [Check Detail] Set [Printed] = 0 Where [Customer Name] = '" & AllPhoneOrders.Item(PHCheckNo).CallerName & "'"
        Sql = Sql & " and [Customer Phone] = '" & AllPhoneOrders.Item(PHCheckNo).CallerPhone & "'"
        Sql = Sql & " and [Check Number] = " & PHCheckNo
        Db.Execute Sql
      End If
      Tb![Presented] = 0
   End If
   Sql = "Update [Check Detail] Set [Check Date] = '" & IsPickUpDate(AllPhoneOrders.Item(PHCheckNo).PDate, PDate) & "' Where [Customer Name] = '" & AllPhoneOrders.Item(PHCheckNo).CallerName & "'"
   Sql = Sql & " and [Customer Phone] = '" & AllPhoneOrders.Item(PHCheckNo).CallerPhone & "'"
   Sql = Sql & " and [Check Number] = " & PHCheckNo
   Db.Execute Sql
End If

DbClose Db, Tb
Set CallTime = Nothing
'See_Msg "Updated....", 5, ""
Exit Sub
ermsg:
Error_Message "Save Phone Update", Err.Description
Resume Next
End Sub
Private Sub Form_Load()
ResizeScreen1024 Me
AbortSelection = False
'DisplayOnly = False
If CaptureCreditCardOrder Then
  SSCommand5.Caption = "Close All"
End If
On Error Resume Next
Me.Left = 0
Me.Top = 0
WatchingforNewDeliveries.Enabled = False
If GetConfigInfo("Phone Orders", "Show Dine In", "", "FUNCTIONS.INI") = "ENABLED" Then
  BtnEnh13.Caption = "Order for Dine In"
End If
If Check_Option(0) = 1 Then
  BtnEnh4.Visible = False
End If
'If LoggedOnAs(0).PhoneAccess Then

'End If
'turn off diningroom tranfer for qsr
Select Case ButtonFunction
     Case Is = BTNFNCT.LoadCallIn
       If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
         Sql = "Select * from [Open Checks] Where [presented] = 1 and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by TRY_CAST([Order Date] AS datetime), TRY_CAST([check Opened] AS datetime) desc"
         Load_CallINs Sql, ""
         CurrentPhoneOrders.Caption = "Order for " & LoggedOnAs(0).EMPName & " Only" & " Open Orders ( " & AllPhoneOrders.count & " )"
       Else
         Load_CallINs "", ""
         CurrentPhoneOrders.Caption = "All Phone Orders" & " Open Orders ( " & AllPhoneOrders.count & " )"
       End If
     Case Is = BTNFNCT.LoadPending
       Load_Pending True, False
End Select
TScroll = 1
ScrollButtons
'driver access from id screen 10.31.2014  -- brandon pizza ken
 Dim Tb As ADODB.Recordset
 Dim GRS As GeneralLib.Utilities
 Set GRS = New GeneralLib.Utilities
 Sql = "Select * from [Open Checks] where [Presented] = 0"
 Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
 If Not Tb Is Nothing Then
   PendingOrders.Caption = "Pending Orders " & RecordsCount(Tb)
 Else
   PendingOrders.Caption = "No Order(s) Pending"
 End If
 Tb.Close
 Set GRS = Nothing
 If GetConfigInfo("Phone Orders", "Order Select Default", "Close", "FUNCTIONS.INI") = "Close" Then
   CurrentPhoneOrders.Caption = "Select Order To Close" & " Open Orders ( " & AllPhoneOrders.count & " )"
   DT_Order_Status = MyOrderStatus(CloseCheck)
   Check_Status = MyOrderStatus(CallIncheck)
 End If
 If LoggedOnAs(0).PhoneAccess Then DriverAccess
End Sub
Private Sub ScrollButtons()
On Error GoTo ERH
Dim sb    As Integer
Dim days  As Integer
DepositAmtPaid = 0
BtnEnh2_Click  'must not stay visible if anybutton is pressed
If AllPhoneOrders Is Nothing Then Exit Sub
If AllPhoneOrders.count = 0 Then Exit Sub
For sb = 0 To 14
   SelectionButton(sb).Visible = False
Next sb

If TScroll > AllPhoneOrders.count Then
  TScroll = AllPhoneOrders.count
Else
  'TScroll = 14
End If
If TScroll = 0 Then TScroll = 1
If TScroll < 0 Then TScroll = 1
For sb = 0 To 14
 If TScroll + sb < 0 Then TScroll = 1
 If TScroll + sb <= AllPhoneOrders.count Then
    SelectionButton(sb).CellsManager.Text(1) = AllPhoneOrders.Item(TScroll + sb).CallerName & vbCrLf & AllPhoneOrders.Item(TScroll + sb).CallerPhone & vbCrLf & AllPhoneOrders.Item(TScroll + sb).callerAddress
    SelectionButton(sb).CellsManager.Text(2) = AllPhoneOrders.Item(TScroll + sb).OrderTaker
    'SelectionButton(sb).CellsManager.Text(3) = ""
    'SelectionButton(sb).CellsManager.Text(4) = ""
    'SelectionButton(sb).CellsManager.Text(5) = ""
    SelectionButton(sb).CellsManager.Text(6) = AllPhoneOrders.Item(TScroll + sb).OrderType
    SelectionButton(sb).CellsManager.Text(8) = AllPhoneOrders.Item(TScroll + sb).PickUpNo
    'SelectionButton(sb).CellsManager.Text(9) = ""
    'SelectionButton(sb).CellsManager.Text(10) = ""
    SelectionButton(sb).Tag = 0
    If Check_Option(62) = 0 Then
      SelectionButton(sb).CellsManager.Text(10) = "Attach Payment"
      SelectionButton(sb).CellsManager.PictureNormal(10) = ""
    End If
    'SelectionButton(sb).CellsManager.Text(9) = AllPhoneOrders.Item(TScroll + sb).AssignmentTime
    If AllPhoneOrders.Item(TScroll + sb).OrderType = "Delivery" Then  'Abs(DateDiff("n", CDate(Tb![Time Assigned]), Time))
     If AllPhoneOrders.Item(TScroll + sb).StatusColor = vbRed Then
      days = DateDiff("d", Date, AllPhoneOrders.Item(TScroll + sb).PDate)
      SelectionButton(sb).CellsManager.Text(7) = Format(AllPhoneOrders.Item(TScroll + sb).PDate, "mm/dd") & AllPhoneOrders.Item(TScroll + sb).PTime & " In " & days & " day(s)"
     Else
      SelectionButton(sb).CellsManager.Text(7) = Format(AllPhoneOrders.Item(TScroll + sb).PDate, "mm/dd") & " Min " & AllPhoneOrders.Item(TScroll + sb).OrderTakenTime
     End If
     If AllPhoneOrders.Item(TScroll + sb).AssignmentTime = "Open" Then
      SelectionButton(sb).CellsManager.Text(9) = "Not Assigned"
     Else
      SelectionButton(sb).CellsManager.Text(9) = "Min(s) Out " & AllPhoneOrders.Item(TScroll + sb).AssignmentTime
     End If
    Else
     If AllPhoneOrders.Item(TScroll + sb).StatusColor = &HC46200 Then
       SelectionButton(sb).CellsManager.Text(7) = Format(AllPhoneOrders.Item(TScroll + sb).PDate, "mm/dd") & " " & Format(AllPhoneOrders.Item(TScroll + sb).PTime, "h:mm ampm") & " Today"
     ElseIf AllPhoneOrders.Item(TScroll + sb).StatusColor = vbRed Then
       days = DateDiff("d", Date, AllPhoneOrders.Item(TScroll + sb).PDate)
       SelectionButton(sb).CellsManager.Text(7) = Format(AllPhoneOrders.Item(TScroll + sb).PDate, "mm/dd") & " " & Format(AllPhoneOrders.Item(TScroll + sb).PTime, "h:mm ampm") & " In " & days & " day(s)"
     Else
       SelectionButton(sb).CellsManager.Text(7) = Format(AllPhoneOrders.Item(TScroll + sb).PDate, "mm/dd") & " " & Format(AllPhoneOrders.Item(TScroll + sb).PTime, "h:mm ampm")
     End If
     SelectionButton(sb).CellsManager.Text(9) = Format(AllPhoneOrders.Item(TScroll + sb).OrderTotal, "0.00")
    End If
    SelectionButton(sb).Tag = AllPhoneOrders.Item(TScroll + sb).PickUpNo
    SelectionButton(sb).BackColorInterior = AllPhoneOrders.Item(TScroll + sb).StatusColor
    SelectionButton(sb).Visible = True
 End If
 If sb = AllPhoneOrders.count Then Exit For
Next sb

Exit Sub
ERH:
See_Msg Err.Number, 0, Err.Description
End Sub
Private Sub Load_Pending(Printall As Boolean, ShowAll As Boolean)
On Error GoTo ermsg
Dim TC              As SendCallin
Dim tn              As Integer
Set TC = New SendCallin
Set AllPhoneOrders = Nothing
Set AllPhoneOrders = New AllCallins
Dim DTHoldStatus    As Integer
Dim CSHoldStatus    As Integer
Dim IMLoggedOn      As String
IMLoggedOn = LoggedOnAs(0).EMPName
Dim Tb              As ADODB.Recordset
Dim GRS             As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
DTHoldStatus = DT_Order_Status
CSHoldStatus = Check_Status

'CurrentPhoneOrders.Caption = "Phone Orders Pending" & " Open Orders ( " & AllPhoneOrders.count & " )"
If Printall = False Then
 Sql = "Select * from [Open Checks] where [presented] = 0 and TRY_CAST([Pick Up date] AS date) = '" & Date & "'"
Else
 Sql = "Select * from [Open Checks] where [presented] = 0 " 'and TRY_CAST([Pick Up date] AS date) = '" & Date & "'"
End If
'Sql = Sql & " Order by TRY_CAST(([Pick Up date] + [Pick Up Time]) AS datetime) asc"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
If Not Tb Is Nothing Then
  If Not Tb.BOF Then Tb.MoveFirst
Else
  See_Msg "error connecting", 5, ""
  Exit Sub
End If

Dim Exp_Mins As Boolean
If Tb.EOF Then
   See_Msg "No Pending Orders Found", 5, "For Today"
   TScroll = 1
   Set TC = Nothing
   DbClose Nothing, Tb
   Set GRS = Nothing
   Set AllPhoneOrders = Nothing
   CurrentPhoneOrders.Visible = False
   Exit Sub
End If
      CurrentPhoneOrders.Visible = True
      Do While Not Tb.EOF
         If Printall Then
          Exp_Mins = TC.CheckTime(Tb(25), Tb(24), Val(Tb.fields("Production").Value & ""))
         End If
         If Printall = False Then Exp_Mins = True
         If Exp_Mins = True Then
            If AllPhoneOrders.Item(CStr(Tb![Check Number])) Is Nothing Then
               AllPhoneOrders.add CStr(Tb![Check Number])
            End If
            AllPhoneOrders.Item(CStr(Tb![Check Number])).CallerName = Tb![Customer Name] & ""
            AllPhoneOrders.Item(CStr(Tb![Check Number])).CallerPhone = Tb![Customer Phone] & ""
            AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderType = Tb![Order Destination] & ""
            AllPhoneOrders.Item(CStr(Tb![Check Number])).PDate = Format(IsPickUpDate(Tb![Pick Up Date] & "", PDate), "MM/DD")
            AllPhoneOrders.Item(CStr(Tb![Check Number])).PTime = IsPickUpDate(Tb![Pick Up Time] & "", PTime)
            AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderTaker = Tb![Server Name] & ""
            AllPhoneOrders.Item(CStr(Tb![Check Number])).PickUpNo = Format(Tb![Check Number] & "", "###")
            If CDate(AllPhoneOrders.Item(CStr(Tb![Check Number])).PDate) > Date Then
              AllPhoneOrders.Item(CStr(Tb![Check Number])).StatusColor = vbRed
            Else
              AllPhoneOrders.Item(CStr(Tb![Check Number])).StatusColor = &HC46200     ' &H400000
            End If
            AllPhoneOrders.Item(CStr(Tb![Check Number])).Deposits = Format(Tb.fields("Deposit").Value, "0.00")
            AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderTotal = Format(Tb.fields("Check Total").Value, "0.00")
            AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderPaymentAmt = Format(Val(Tb.fields("Payments").Value & ""), "0.00")
            AllPhoneOrders.Item(CStr(Tb![Check Number])).ProductionTime = Val(Tb![Production] & "")
            AllPhoneOrders.Item(CStr(Tb![Check Number])).callerAddress = Tb.fields("Street").Value & ""
         End If
        Tb.MoveNext
      Loop
DbClose Nothing, Tb: Set GRS = Nothing: Set TC = Nothing
If Printall Then
    'IMLoggedOn = LoggedOnAs(0).EMPName
        For tn = 1 To AllPhoneOrders.count
         If AllPhoneOrders.Item(tn).PickUpNo <> -9999 Then
            LoggedOnAs(0).EMPName = AllPhoneOrders.Item(tn).OrderTaker
            If Load_CallIn_Order(AllPhoneOrders.Item(tn).PickUpNo, AllPhoneOrders.Item(tn).CallerName) = False Then Exit Sub
            Clear_Grid Order.Ticket(0)
            Order.Load_Grid
            DT_Order_Status = MyOrderStatus(ParkCheck)
            If Check_Status = MyOrderStatus(TranferPhone) Then
               'release reservation
            Else
              '// need release here
              Check_Status = MyOrderStatus(ReleaseCallIn)
            End If
            Order.Done.Caption = ReturnLabel("Done")
            AllPhoneOrders.Item(tn).OrderOrgination = "AutoRelease"
            Order.Done.DoClick
            Trans_Log "Sending Pending Orders " & AllPhoneOrders.Item(tn).PickUpNo
         End If
        Next tn
    LoggedOnAs(0).EMPName = IMLoggedOn
End If
DT_Order_Status = DTHoldStatus
Check_Status = CSHoldStatus
CurrentPhoneOrders.Caption = "Phone Orders Pending" & " Open Orders ( " & AllPhoneOrders.count & " )"
Exit Sub
ermsg:
LoggedOnAs(0).EMPName = IMLoggedOn
'// make sure that we dont change cashiers.

Error_Message "Load Pending", Err.Description
Resume Next
End Sub
Public Sub Load_CallINs(UseSql As String, ListBy As String)
On Error GoTo ermsg
Dim Tb As ADODB.Recordset
Dim DriverCheck  As Long
Dim ShowPUT      As Byte
Set AllPhoneOrders = Nothing
Set AllPhoneOrders = New AllCallins
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim DSql  As String
WatchingforNewDeliveries.Enabled = False
 DSql = "Delete from [Open Checks] Where [Pick Up date] is null"
Call GRS.ExecuteSql(DSql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb))
If UseSql = "" Then
    Sql = "Select * from [Open Checks] where [Presented] = 1 or TRY_CAST([Pick Up Date] AS date) = '" & Date & "'"
    If LoggedOnAs(0).PhoneAccess Then
      Sql = Sql & " and [Order Destination] = 'Delivery'"
      Sql = Sql & " and [Driver Assigned] = ''"
    End If
    Sql = Sql & " Order by [check number] asc"
Else
   Sql = UseSql
End If
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
If Not Tb Is Nothing Then
   If Not Tb.BOF Then Tb.MoveFirst
Else
   See_Msg "error connecting", 5, ""
   Exit Sub
End If
Do While Not Tb.EOF
    If AllPhoneOrders.Item(CStr(Tb![Check Number])) Is Nothing Then
       AllPhoneOrders.add CStr(Tb![Check Number])
    End If
    AllPhoneOrders.Item(CStr(Tb![Check Number])).CallerName = Tb.fields("Customer Name").Value & ""
    AllPhoneOrders.Item(CStr(Tb![Check Number])).CallerPhone = Tb.fields("Customer Phone").Value & ""
    AllPhoneOrders.Item(CStr(Tb![Check Number])).callerAddress = Tb.fields("Street").Value & ""
    AllPhoneOrders.Item(CStr(Tb![Check Number])).DriverName = Tb.fields("Driver Assigned").Value & ""
     If Tb![Order Destination] & "" = "Delivery" Then
      If Tb![Presented] = 0 Then
       AllPhoneOrders.Item(CStr(Tb![Check Number])).PTime = IsPickUpDate(Tb![Pick Up Time] & "", PTime)
      Else
       If Tb![Time assigned] & "" = "" Then
        AllPhoneOrders.Item(CStr(Tb![Check Number])).PTime = "Open"
       Else
        AllPhoneOrders.Item(CStr(Tb![Check Number])).PTime = Format(Tb![Time assigned] & "", "H:MM ampm")
       End If
      End If
     Else
      AllPhoneOrders.Item(CStr(Tb![Check Number])).PTime = IsPickUpDate(Tb![Pick Up Time] & "", PTime)
     End If
     AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderType = Tb.fields("Order Destination").Value & ""
     
     If IsPickUpDate(Tb![Pick Up Date], PDate) = Date Then
        AllPhoneOrders.Item(CStr(Tb![Check Number])).PDate = "Today"
     Else
        AllPhoneOrders.Item(CStr(Tb![Check Number])).PDate = Format(Tb![Pick Up Date] & "", "M/D")
     End If
     AllPhoneOrders.Item(CStr(Tb![Check Number])).Deposits = Format(Nz(Tb.fields("Deposit").Value), "0.00")
     AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderTaker = Tb![Server Name] & ""
     AllPhoneOrders.Item(CStr(Tb![Check Number])).PickUpNo = Format(Nz(Tb![Check Number]), "###")
     If IsDate(Tb![Time assigned]) Then
      AllPhoneOrders.Item(CStr(Tb![Check Number])).AssignmentTime = Abs(DateDiff("n", CDate(Tb![Time assigned]), Time))
     Else
      AllPhoneOrders.Item(CStr(Tb![Check Number])).AssignmentTime = "Open"
     End If
     If IsDate(Tb![check Opened]) Then
      If IsDate(Tb.fields("Driver Return").Value) Then
        AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderTakenTime = Abs(DateDiff("n", CDate(Tb![check Opened]), Tb.fields("Driver Return").Value))
      Else
       AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderTakenTime = Abs(DateDiff("n", CDate(Tb![check Opened]), Time))
      End If
     End If
     If IsDate(Tb.fields("Driver Return").Value & "") Then
      AllPhoneOrders.Item(CStr(Tb![Check Number])).DriveReturned = Format(Tb.fields("Driver Return").Value & "", "H:MM ampm")
     Else
      AllPhoneOrders.Item(CStr(Tb![Check Number])).DriveReturned = ""
     End If
     AllPhoneOrders.Item(CStr(Tb![Check Number])).PDate = Format(Tb![Pick Up Date] & "", "M/D")
     AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderTotal = Format(Nz(Tb![Check Total]), "0.00")
     AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderPaymentAmt = Format(Nz(Tb.fields("Payments").Value), "0.00")
     AllPhoneOrders.Item(CStr(Tb![Check Number])).OrderOrgination = Tb.fields("Register Name").Value & ""
     '"Web Server"
   If Tb![Presented] = 1 Then
      AllPhoneOrders.Item(CStr(Tb![Check Number])).StatusColor = &H80000012
   ElseIf CDate(IsPickUpDate(Tb![Pick Up Date] & "", PDate)) <> Format(Date, "MM/DD/YYYY") Then
     AllPhoneOrders.Item(CStr(Tb![Check Number])).StatusColor = vbRed
   Else
     AllPhoneOrders.Item(CStr(Tb![Check Number])).StatusColor = &HC46200    ' &H400000  'Blue Color
   End If
NexR:
Tb.MoveNext
Loop
DbClose Nothing, Tb: Set GRS = Nothing
If AllPhoneOrders.count = 0 Then
  CurrentPhoneOrders.Visible = False
Else
  CurrentPhoneOrders.Visible = True
End If
TScroll = 1
CurrentPhoneOrders.Caption = ReturnLabel("Phone Orders") & " Open Orders ( " & AllPhoneOrders.count & " )"
Me.Refresh
'WatchingforNewDeliveries.Enabled = True
Exit Sub
ermsg:
Error_Message "Load Callins", Err.Description
Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
WatchingforNewDeliveries.Enabled = False
Set AllPhoneOrders = Nothing
End Sub
Public Function GetDeposit(CustomerCheck As Long, CustomerName As String, CustomerPhone As String, CheckTotal As Currency) As Currency
GetDeposit = 0:   CheckTotal = 0
 Sql = "Select * from [Open Checks] Where [Customer Name] = '" & CustomerName & "'"
 Sql = Sql & " and [Customer Phone] = '" & CustomerPhone & "'"
 Sql = Sql & " and [Check Number] = " & CustomerCheck
Dim RS As ADODB.Recordset
Dim GRS         As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Set RS = GRS.ReturnRecordset(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb), "Open Checks")
If Not RS Is Nothing Then
   If Not RS.BOF Then RS.MoveFirst
Else
   See_Msg "error connecting", 5, ""
   Exit Function
End If
If Not RS.EOF Then
 GetDeposit = Val(RS![Deposit] & "")
 CheckTotal = Val(RS![Check Total] & "")
End If
DbClose Nothing, RS
Set GRS = Nothing
End Function
Public Sub RecordDeposit(CustomerCheck As Long, CustomerName As String, CustomerPhone As String, AdditionDeposit As Currency)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
    Sql = "Select * from [Open Checks] Where [Customer Name] = '" & CustomerName & "'"
    Sql = Sql & " and [Customer Phone] = '" & CustomerPhone & "'"
    Sql = Sql & " and [Check Number] = " & CustomerCheck
ConnectADORS Sql, Db, RS
 If Not RS.EOF Then
   RS![Deposit] = Val(RS![Deposit] & "") + AdditionDeposit
   RS.Update
 End If
 DbClose Db, RS
End Sub
Public Sub LoadCustomer(CustomerName As String, PHONENO As String, CustomerAddress As String, CustomerCity As String)
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim CTotal As Currency
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
searchagain:
    If CustomerName = "Call In" Then
     Sql = "Select * from [Customer Data] Where [Customer Phone] = '" & PHONENO & "'"
    Else
     Sql = "Select * from [Customer Data] Where [Customer Name] = '" & CustomerName & "'"
     Sql = Sql & " and [Customer Phone] = '" & PHONENO & "'"
    End If
    If CustomerAddress > "" And CustomerAddress <> "Enter Address" Then
     Sql = Sql & " and [Customer Address] = '" & CustomerAddress & "'"
    End If
ConnectADORS Sql, Db, RS
 If Not RS.EOF Then
    If CustomerAddress > "" Then
      CInfo(1).Text = CustomerAddress
    Else
      CInfo(1).Text = RS![Customer Address] & ""
    End If
    If CustomerCity > "" Then
      CInfo(2).Text = CustomerCity
     Else
      CInfo(2).Text = RS![Customer City] & ""
     End If
    If CustomerName = "Call In" Then
     CInfo(0).Text = RS![Customer Name] & ""
    Else
     CInfo(0).Text = CustomerName
    End If
    CInfo(3).Text = RS![Customer State] & ""
    CInfo(4).Text = RS![Customer Zip] & ""
    CInfo(5).Text = PHONENO
    CInfo(6).Text = RS![Instructions] & ""
    CInfo(7).Text = Format(GetDeposit(DepositCheckNo, CStr(CInfo(0).Text), CStr(CInfo(5).Text), CTotal), "0.00")
    CheckAMT.Text = Format(CTotal, "0.00")
    'PaymentAmount.Text = PhoneOrderIsPaid(DepositCheckNo)
   With CustomerINfo
       CenterObject Me, CustomerINfo
       .ZOrder 0
       .Visible = True
   End With
 Else
   If CustomerAddress > "" Then
    CustomerAddress = ""
    GoTo searchagain
   Else
    See_Msg "Customer Not Found", 5, ""
   End If
 End If
DbClose Db, RS
End Sub
Public Sub UpdateCustomer(CustomerName As String, PHONENO As String, ChkNO As String, SavedDeposit As Currency)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
    Sql = "Select * from [Customer Data] Where [Customer Name] = '" & CustomerName & "'"
    Sql = Sql & " and [Customer Phone] = '" & PHONENO & "'"
ConnectADORS Sql, Db, RS
    If RS.EOF Then RS.AddNew
     RS![Customer Name] = CInfo(0).Text & ""
     RS![Customer Address] = CInfo(1).Text & ""
     RS![Customer City] = CInfo(2).Text & ""
     RS![Customer State] = CInfo(3).Text & ""
     RS![Customer Zip] = CInfo(4).Text & ""
     RS![Customer Phone] = CInfo(5).Text & ""
     RS![Instructions] = CInfo(6).Text & ""
     On Error Resume Next
    RS.Update
   
   SetNothing RS
   Sql = "Select * from [Open Checks] where [Customer Phone] = '" & PHONENO & "'"
   Sql = Sql & " and [Customer Name] = '" & CustomerName & "'"
   Sql = Sql & " and [Check Number] = " & Val(ChkNO)
 ConnectADORS Sql, Db, RS
   Do While Not RS.EOF
     RS![Customer Name] = CInfo(0).Text & ""
     RS![Customer Phone] = CInfo(5).Text & ""
     RS![Street] = CInfo(1).Text & ""
     RS.fields("Deposit").Value = Format(SavedDeposit, "0.00")
     If IsDate(PUTime.Text) Then RS.fields("Pick Up Time").Value = Format(PUTime.Text, "HH:MM ampm")
     If IsDate(PUDATE.Text) Then RS.fields("Pick Up date").Value = Format(PUDATE.Text, "MM/DD/YYYY")
     RS.Update
    RS.MoveNext
   Loop
  Sql = "Select * from [Check Detail] where [Customer Phone] = '" & PHONENO & "'"
  Sql = Sql & " and [Customer Name] = '" & CustomerName & "'"
  Sql = Sql & " and [Check Number] = " & Val(ChkNO)
 ConnectADORS Sql, Db, RS
   Do While Not RS.EOF
     RS![Customer Name] = CInfo(0).Text & ""
     RS![Customer Phone] = CInfo(5).Text & ""
     RS.Update
    RS.MoveNext
   Loop
 DbClose Db, RS
 See_Msg "Updated...", 5, ""
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Update Customer ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function Load_CallIn_Order(Get_Check As Long, OrderName As String) As Boolean
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim GPH         As ADODB.Recordset
Dim RP1         As String
Dim RP2         As String
Dim RP3         As String
Dim RP4         As String
Dim RP5         As String
Dim RP6         As String
Dim IsPresented As Integer
OrderName = Trim$(OrderName)
If NumberInUse(Get_Check) Then
 Load_CallIn_Order = False
 Exit Function
End If
Load_CallIn_Order = True
   
Sql = "Select * from [Open Checks] where [Check Number] = " & Get_Check
If OrderName > "" Then
   Sql = Sql & " and [Customer Name] = '" & OrderName & "'"
End If
If ConnectADODB(SvrPath.TerminalPath & CallInOrderDb, Db) = False Then
   See_Msg "error connecting", 5, ""
   Load_CallIn_Order = False
   Exit Function
End If

If ConnectADORS(Sql, Db, Tb) = False Then
   See_Msg "error connecting", 5, ""
   Load_CallIn_Order = False
   Exit Function
End If
If Tb.EOF Then
   DbClose Db, Tb
   Load_CallIn_Order = False
   See_Msg "Search Failed", 5, ""
   Exit Function
Else
   Table_Number = Tb.fields("Table Number").Value & ""
   Check_no = Get_Check
   Number_In_Party = Tb.fields("Number In Party").Value
   IsPresented = Val(Tb.fields("Presented").Value & "")
   OrderIs = Tb.fields("Order Destination").Value
    
   If Tb![Customer Phone] & "" = "" Then
      Tb![Customer Phone] = "Not Found"
      Tb.Update
   End If
    
   Set AllOrders = Nothing
   Set AllOrders = New AllCallins
   AllOrders.add Tb![Customer Phone]
   With AllOrders.Item(Tb![Customer Phone])
       .ProductionTime = Val(Tb.fields("Production").Value & "")
       '.Deposits = Val(Tb![Deposit] & "")
       .CallerName = Trim$(Tb![Customer Name])  'was set to table number
       OrderName = Trim$(Tb![Customer Name])
       If Check_Status = MyOrderStatus(ReleaseCallIn) Then
         If See_Msg("Release for Now ?", 4, "No keep original order date") = 6 Then
          .PDate = Format(Date, "MM/DD/YYYY")
          .PTime = IsPickUpDate(Time, PTime)
          Tb.fields("Pick Up Date").Value = Date
          Tb.fields("Pick Up Time").Value = Time
          Tb.Update
         Else
          .PDate = IsPickUpDate(Tb.fields("Pick Up Date").Value & "", PDate)
          .PTime = IsPickUpDate(Tb.fields("Pick Up Time").Value & "", PTime)
         End If
       Else
        .PDate = IsPickUpDate(Tb.fields("Pick Up Date").Value & "", PDate)
        .PTime = IsPickUpDate(Tb.fields("Pick Up Time").Value & "", PTime)
        If IsPresented = 0 Then
         If Check_Status = MyOrderStatus(CallIncheck) And CDate(.PDate) > Date Then
          CollectDeposit = True
         End If
        End If
       End If
       .OrderTime = Tb.fields("Check Opened").Value & ""
       .PickUpNo = Check_no
       .CallerPhone = Tb![Customer Phone] & ""
       .callerAddress = Tb![Street] & ""
       .DriverName = Tb![Server Name] & ""
       .VoidOrder = False
        On Error Resume Next
        .DriveReturned = Tb.fields("Driver Return").Value & ""
        If Err.Number <> 0 Then
         .DriveReturned = ""
        End If
        On Error GoTo ErrorHandler
        If Tb![Transfer From] = "House Charge" Then
          .HouseAccountNo = Tb.fields("Member Name").Value & ""
          .OrderPayment = "Web HA"
          .OrderPaymentAmt = 0
        Else
          .HouseAccountNo = ""
          .OrderPayment = DeliveryPrePaidTender(Check_no) ' Tb.Fields("Member Name").Value & ""    '// if not tender name then shows as prepaid
          'If .OrderPayment = "Deposit" Then
           .Deposits = Val(Tb.fields("Deposit").Value & "")
          'Else
           .OrderPaymentAmt = Val(Tb.fields("Payments").Value & "")
           '.Deposits = 0
          'End If
        End If
        .ChangeOrder = ChangeOrderDestination
        Sql = " Select * from [Phone Numbers] Where [Customer Phone] = '" & Tb![Customer Phone] & "" & "'"
        If ConnectADORS(Sql, Db, GPH) = False Then Exit Function
        If Not GPH.EOF Then
         .CustomerId = GPH.fields("CustomerId").Value & ""
        End If
        GPH.Close
        Sql = " Select * from [Customer Data] Where [Customer Phone] = '" & Tb![Customer Phone] & "" & "'"
        If ConnectADORS(Sql, Db, GPH) = False Then Exit Function
        If Not GPH.EOF Then
         .Email = GPH.fields("EMail").Value & ""
         .City = GPH.fields("Customer City").Value & " " & GPH.fields("Customer State").Value & " " & GPH.fields("Customer Zip").Value
        End If
        GPH.Close
        If .CustomerId > "" Then
         Sql = " Select * from [Instructions] Where [CustomerId] = '" & .CustomerId & "" & "'"
         If ConnectADORS(Sql, Db, GPH) = False Then Exit Function
         If Not GPH.EOF Then
          .Staffnotes = GPH.fields("StaffNotes").Value & ""
         End If
         GPH.Close
        End If
        If ChangeOrderDestination = "AddDeliveryCharge" Then
         Dim DriveInfo   As DrvZoneData
         DriveInfo = GetDelivery(AllOrders.Item(Tb![Customer Phone]).callerAddress, AllOrders.Item(Tb![Customer Phone]).City, AllOrders.Item(Tb![Customer Phone]).State, AllOrders.Item(Tb![Customer Phone]).Zip)
         If DriveInfo.ZoneName = "Out of Area" Then
            ChangeOrderDestination = ""
            Exit Function
         End If
         .Zone = DriveInfo.ZoneName
         .Distance = DriveInfo.ZoneDistance
        End If
   End With
End If
SetNothing Tb
Top:
   'Clear_Grid Order.Ticket(0)
   ReDim_Array
   Selection_Number = Abs(GetLastSelectionNumber(Get_Check, Db))
   Sql = "Select * from [Check detail] where [Check Number] = " & Get_Check
   If OrderName > "" Then
     Sql = Sql & " and [customer Name] = '" & OrderName & "'"
   End If
   Sql = Sql & " Order By Abs([Selection Number]),[Pos] asc "
   If ConnectADORS(Sql, Db, Tb) = False Then
       See_Msg "error connecting", 5, ""
       Exit Function
    End If
    ReDim NextRow(Number_In_Party)
    If Tb.EOF Then
      Sql = "Delete from [Open Checks] where [Check Number] = " & Get_Check
      Db.Execute Sql
      Exit Function
    End If
    Do While Not Tb.EOF
      If Val(Tb.fields("Quantity").Value & "") = 0 Then
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 0) = ""
      Else
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 0) = Tb.fields("Quantity").Value
      End If
      If Val(Tb(8)) > 0 Then
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1) = Chr(118) & "oid-> " & Tb.fields("Menu Item").Value
      ElseIf IsPresented = 0 Then
        If Tb![Record Type] = 0 And Check_Option(194) = 1 And Val(Tb(2) & "") > 0 Then
           Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1) = "(" & Tb(2) & ") " & Tb.fields("Menu Item").Value
           Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 0) = ""
        Else
           Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1) = Tb.fields("Menu Item").Value
        End If
      Else
        If Tb![Record Type] = 0 And Check_Option(194) = 1 And Val(Tb.fields("Quantity").Value & "") > 0 Then
           Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1) = "* " & "(" & Tb.fields("Quantity").Value & ") " & Tb.fields("Menu Item").Value
           Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 0) = ""
        Else
           Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1) = "* " & Tb.fields("Menu Item").Value
        End If
      End If
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 28) = Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1)
      If Tb.fields("Price").Value = 0 Then
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 2) = ""
      Else
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 2) = Align(7, Tb.fields("Price").Value)
      End If
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 3) = GetExpandPrint(Tb.fields("Menu Item").Value)
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 10) = Tb.fields("Tax Table").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 11) = Tb.fields("Special Tax").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 12) = Tb.fields("Liquor Tax").Value
      If Tb.fields("Voided").Value = 1 Then
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 13) = -1 * Tb.fields("Voided").Value
      Else
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 13) = Tb.fields("Voided").Value
      End If
      If Val(Abs(Tb.fields("Selection Number").Value)) > Selection_Number Then Selection_Number = (Abs(Tb.fields("Selection Number").Value) + 1)
      If IsPresented = 0 Then
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 14) = Abs(Tb.fields("Selection Number").Value)  'hold selection number for voided items
      Else
         Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 14) = Tb.fields("Selection Number").Value  'hold selection number for voided items
      End If
      
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 3) = Tb![Expand Print]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 4) = Tb![Remote Print]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 5) = Tb![Remote 1]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 6) = Tb![Remote 2]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 7) = Tb![Remote 3]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 8) = Tb![Remote 4]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 9) = Tb![Remote 5]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 30) = GetLink(Tb.fields("Menu Item").Value & "")
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 31) = Tb![Remote 6]
      
      If Check_Option(189) = 1 Then
        Dim GetOriginalItem As String
        GetOriginalItem = Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 30)
        If Not AllMenuItems(CStr(GetOriginalItem)) Is Nothing Then
           RP1 = AllMenuItems(CStr(GetOriginalItem)).Remote1
        Else
           RP1 = "None"
        End If
        If RP1 <> "None" Then Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 5) = GetRemote(RP1, Tb![Department Report] & "")
        If Not AllMenuItems(CStr(GetOriginalItem)) Is Nothing Then
           RP2 = AllMenuItems(CStr(GetOriginalItem)).Remote2
        Else
           RP2 = "None"
        End If
        If RP2 <> "None" Then Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 6) = GetRemote(RP2, Tb![Department Report] & "")
        If Not AllMenuItems(CStr(GetOriginalItem)) Is Nothing Then
           RP3 = AllMenuItems(CStr(GetOriginalItem)).Remote3
        Else
           RP3 = "None"
        End If
        If RP3 <> "None" Then Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 7) = GetRemote(RP3, Tb![Department Report] & "")
        If Not AllMenuItems(CStr(GetOriginalItem)) Is Nothing Then
           RP4 = AllMenuItems(CStr(GetOriginalItem)).Remote4
        Else
           RP4 = "None"
        End If
        If RP4 <> "None" Then Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 8) = GetRemote(RP4, Tb![Department Report] & "")
        If Not AllMenuItems(CStr(GetOriginalItem)) Is Nothing Then
           RP5 = AllMenuItems(CStr(GetOriginalItem)).Remote5
        Else
           RP5 = "None"
        End If
        If RP5 <> "None" Then Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 9) = GetRemote(RP5, Tb![Department Report] & "")
        If Not AllMenuItems(CStr(GetOriginalItem)) Is Nothing Then
           RP6 = AllMenuItems(CStr(GetOriginalItem)).Remote6
        Else
           RP6 = "None"
        End If
        If RP6 <> "None" Then Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 31) = GetRemote(RP6, Tb![Department Report] & "")
      End If
      If Left$(Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 1), 4) = "void" Then
       Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 5) = "None"
       Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 6) = "None"
       Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 7) = "None"
       Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 8) = "None"
       Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 9) = "None"
      End If

      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 15) = Tb.fields("Sort Level").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 16) = Tb.fields("Printed").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 17) = Tb.fields("Disc Type").Value & ""
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 18) = Tb.fields("Disc Amount").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 19) = Val(Tb![Record Type] & "")
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 20) = Tb![Voided By] & ""
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 21) = Tb![Void Reason] & ""
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 22) = Tb![Void Time] & ""
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 23) = Tb.fields("Cateory Name").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 24) = Tb.fields("Seat Number").Value
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 25) = CStr(Tb![Menu Number])
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 26) = Tb![Department Report]
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 32) = Val(Tb![Inclusive Tax] & "")
      Grid_array(Tb.fields("Seat Number").Value, NextRow(Tb.fields("Seat Number").Value), 27) = Val(Tb![Cost Amount] & "")
      NextRow(Tb.fields("Seat Number").Value) = NextRow(Tb.fields("Seat Number").Value) + 1
     Tb.MoveNext
 Loop
 AllOrders.Item(1).OrderGUID = ""
 'If Check_Option(111) = 1 Then
 ' Tb.Close
 ' Sql = "Select * from CustomerNotification where [CheckNumber] = '" & CStr(Check_no) & "'"
 ' If ConnectADORS(Sql, Db, Tb) = False Then
 
 ' Else
 '  If Not Tb.EOF Then
 '   AllOrders.Item(1).OrderGUID = Tb.Fields("GuidId").Value & ""
 '  End If
 ' End If
 'End If
DbClose Db, Tb
Clear_Grid Order.Ticket(0)

Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Call In Order ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub Cancel_CallIn_Order(Get_Check As Long, Which_Check As String)
On Error GoTo ErrorHandler
Dim r_sults  As String
If See_Msg("Are You Sure ?", 4, "Cancel Order") <> 6 Then Exit Sub
If ChangeToRealCheckNumber(Get_Check) = True Then
 Delete_Payment_Record Get_Check, "", False
 Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
    Sql = "Delete from [Open Checks] where [Check Number] = " & Get_Check
   If Which_Check > "" Then Sql = Sql & " and [Customer Name] = '" & Trim$(Which_Check) & "'"
  r_sults = GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb))
   Sql = "Delete from [Check detail] "
   Sql = Sql & " where [Check Number] = " & Get_Check
   If Which_Check > "" Then Sql = Sql & " and [customer Name] = '" & Trim$(Which_Check) & "'"
  r_sults = GRS.ExecuteSql(Sql, ReturnAdoConnection(SvrPath.TerminalPath & CallInOrderDb))
  See_Msg "Order Canceled.", 5, ""
 Set GRS = Nothing
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Load Call In Order ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub PaymentAmount_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub PaymentAmount_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub PaymentAmount_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub PUDATE_Click()
Dim InDate  As String
'Dim OrgDate As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
    InDate = PKT.ReturnDate
    If IsDate(InDate) Then
     'OrgDate = PUDATE.Text
     PUDATE.Text = Format(InDate, "MM/DD/YYYY")
     AllPhoneOrders.Item(WorkingCheck.Caption).PDate = Format(InDate, "MM/DD/YYYY")
    ' SavePhoneUpdate WorkingCheck, OrgDate
    End If
Set PKT = Nothing
End Sub
Private Sub PUTime_Click()
Dim InDate  As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
    InDate = PKT.ReturnTime
    If IsDate(InDate) Then
     PUTime.Text = Format(InDate, "HH:MM ampm")
     AllPhoneOrders.Item(WorkingCheck.Caption).PTime = Format(InDate, "HH:MM ampm")
    Else
      See_Msg "Invalid Time", 5, InDate
    End If
Set PKT = Nothing
End Sub
Private Sub ReOpenCheck_Click()
If Can_Continue(ReOpenCheck.Caption) = True Then
 Load_What = "Reopen Phone Order"
 Find_What = "Reopen Phone Order"
 RequestListForm.SSCommand12.Caption = "Search Check No."
 RequestListForm.Show 1
 If GetSelected = "" Then
  Exit Sub
 End If
 ReopenCheckReview.ReviewClosedCheck Val(GetSelected)
 ReopenCheckReview.SSCommand12.Caption = "Reopen Check"
 ReopenCheckReview.Show 1
 
 If Val(GetSelected) = 0 Then Exit Sub  '// cancel reopen when check number is set to 0
 If IsCallInCheckNumberInUse(Val(GetSelected)) Then
    See_Msg "This Check Number Is In Use", 5, "Check cannot be reopened"
    Exit Sub
 End If
 If ResetPhoneOrder(GetSelected) Then
   UnDoPhoneOrder Val(GetSelected)
   See_Msg "Transaction Reopened", 5, GetSelected
   Trans_Log "Reopened Phone Order " & GetSelected
   SSCommand15_Click
 End If
End If
End Sub
Private Sub SearchHow_Click(Index As Integer)
Dim searchby As String
Sql = ""
BtnEnh9_Click
Select Case Index
   Case Is = 0
      searchby = ShowKeyBoard("Enter Customer Name")
      If searchby > "" Then
        If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
          Sql = "Select * from [Open Checks] Where [Customer Name] Like '%" & searchby & "%' and [Server Name] = '" & LoggedOnAs(0).EMPName & "' Order by [Customer Name] asc"
        Else
          Sql = "Select * from [Open Checks] Where [Customer Name] Like '%" & searchby & "%' Order by [Customer Name] asc"
        End If
      End If
   Case Is = 1
     searchby = ShowKeyBoard("Enter Phone")
     If searchby > "" Then
      If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
        Sql = "Select * from [Open Checks] Where [Customer Phone] Like '%" & searchby & "%' and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by [Customer phone] asc"
      Else
        Sql = "Select * from [Open Checks] Where [Customer Phone] Like '%" & searchby & "%' order by [Customer phone] asc"
      End If
     End If
   Case Is = 2
     searchby = ShowKeyBoard("Enter Street Address")
     
     If searchby > "" Then
      If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
        Sql = "Select * from [Open Checks] Where [Street] Like '%" & searchby & "%' and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by [Street] asc"
      Else
        Sql = "Select * from [Open Checks] Where [Street] Like '%" & searchby & "%' order by [Street] asc"
      End If
     End If
   Case Is = 3
     If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
        See_Msg "No Access", 5, "Access To Low"
        Exit Sub
     Else
        Find_What = "Select Driver"
        Load_What = "Drivers"
        GetSelected = "": searchby = "": Sql = ""
        RequestListForm.Show 1
        searchby = GetSelected
        Sql = ""
        If searchby > "" Then
          If searchby = "All" Then
            Sql = "Select * from [Open Checks] Where [Server Name] ='" & searchby & "'"
          Else
            Sql = "Select * from [Open Checks] Where [Server Name] ='" & searchby & "'"
          End If
          
           Sql = Sql & " Order by TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
          
          SearchHow(3).Caption = "Driver Name"
        End If
     End If
   Case Is = 4
     If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
        Sql = "Select * from [Open Checks] Where [Server Name] = '" & LoggedOnAs(0).EMPName & "'"
     Else
        Sql = "Select * from [Open Checks]"
     End If
     Sql = Sql & " Order by TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
  Case Is = 5
     Sql = "Select * from [Open Checks] Where TRY_CAST([Pick Up date] AS date) <= '" & Date & "' and [Presented] = 1 and [Order Destination] = 'Delivery' and [Time Assigned] is null or TRY_CAST([Pick Up date] AS date) <= '" & Date & "' and [Order Destination] = 'Delivery' and [Time Assigned] = '' order by [Check Number] asc"
     
   Case Is = 6
       Load_Pending False, False
       ScrollButtons
       Me.Refresh
       Exit Sub
   Case Is = 7
        LoadDates "Select Distinct [Date] from [Sales] Where not isnull([date]) order by [Date] asc", SvrPath.TerminalPath & ClosingHistory, "Date"
        If SearchDates Is Nothing Then
         Exit Sub
        End If
         Sql = "Select * from [Open Checks]"
         Sql = Sql & " Where [presented] = 0 and TRY_CAST([Pick Up Date] AS date) Between '" & SearchDates.Item(1).StartDate & "' And '" & SearchDates.Item(1).EndDate & "'"
         Sql = Sql & " order by TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up time] AS datetime) asc"
End Select
If Sql > "" Then
  Load_CallINs Sql, ""
  TScroll = 1
  ScrollButtons
  CurrentPhoneOrders.Caption = " Open Orders ( " & AllPhoneOrders.count & " )"
  If Index = 5 Then
    WatchingforNewDeliveries.Enabled = True '// press unassigned deliveries to turn on this timer.
  End If
End If
End Sub
Private Sub SelectionButton_CellClick(Index As Integer, ByVal nUniqueID As Integer)
On Error GoTo ErrorHandler
Dim DTHoldStatus As Integer
Dim CSHoldStatus As Integer
Dim DriverAssignedToOrder As String
BtnEnh2_Click
SelectedRow = Index
ReleaseRegisterLock
If Check_Status = MyOrderStatus(ReleaseCallIn) Then
  If AllPhoneOrders.Item(SelectionButton(Index).Tag).StatusColor <> &HC46200 And AllPhoneOrders.Item(SelectionButton(Index).Tag).StatusColor <> vbRed Then     '&H00C46200&
      See_Msg "Order Already processed", 5, ""
      Check_Status = MyOrderStatus(CallIncheck)
      DT_Order_Status = MyOrderStatus(CloseCheck)
      Exit Sub
  End If
End If
WorkingCheckNo = SelectionButton(Index).Tag
AssignmentCaption.Caption = "Order Selected " & AllPhoneOrders.Item(SelectionButton(Index).Tag).CallerPhone
OrderAddress.Caption = AllPhoneOrders.Item(SelectionButton(Index).Tag).callerAddress
Trans_Log "Phone Order Accessed " & WorkingCheckNo & " " & AssignmentCaption.Caption
If DT_Order_Status = MyOrderStatus(ParkCheck) Or Check_Status = MyOrderStatus(ReleaseCallIn) Then
    If Load_CallIn_Order(AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo, AllPhoneOrders.Item(SelectionButton(Index).Tag).CallerName) = False Then Exit Sub
    Call MemoryCheck(PhoneOrders)
    Exit Sub
End If
Select Case nUniqueID
  Case Is = 1
  
  Case Is = 2
     'Driver Assigned Caption
  Case Is = 3
    'driver assignment
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).StatusColor = &HC46200 Or AllPhoneOrders.Item(SelectionButton(Index).Tag).StatusColor = vbRed Then    '&H00C46200&
      See_Msg "Must Release Order to Kitchen", 5, "Cannot Assign Driver At This Time"
      Exit Sub
    End If
    
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType = "Reservation" Then Exit Sub
    AssignDriver AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType
    Exit Sub
  Case Is = 4 'Close order
    If LoggedOnAs(0).PhoneAccess Then Exit Sub
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).StatusColor = &HC46200 Or AllPhoneOrders.Item(SelectionButton(Index).Tag).StatusColor = vbRed Then    '&H00C46200&
      See_Msg "Must Release Order to Kitchen", 5, ""
      Exit Sub
    End If
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo = -9999 Then
      See_Msg "Must Be assigned to a Server", 5, "Can not Close"  'Reservation number
      Exit Sub
    End If
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType = "Pick Up" Then
       If Can_Continue("Close Pick Up Order") = False Then Exit Sub
       DriverAssignedToOrder = LoggedOnAs(0).EMPName
    End If
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType = "Catering" Then
       If Can_Continue("Close Catering Orders") = False Then Exit Sub
        DriverAssignedToOrder = LoggedOnAs(0).EMPName
    End If
    
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType = "Delivery" Then
       
       If GetConfigInfo("Phone Orders", "Delivery Closed With No Assignment", "No", "FUNCTIONS.INI") = "ENABLED" Then
         AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderTaker = LoggedOnAs(0).EMPName
         AllPhoneOrders.Item(SelectionButton(Index).Tag).AssignmentTime = Time  'added for D Peel 08/08/2023
         AllPhoneOrders.Item(SelectionButton(Index).Tag).DriverName = LoggedOnAs(0).EMPName
       End If
       
       If AllPhoneOrders.Item(SelectionButton(Index).Tag).AssignmentTime = "Open" Then
          See_Msg "No Driver Assigned", 5, "Assignment Required"
          Exit Sub
       End If
       DriverAssignedToOrder = AllPhoneOrders.Item(SelectionButton(Index).Tag).DriverName
       If DriverAssignedToOrder = "" Then DriverAssignedToOrder = LoggedOnAs(0).EMPName
       If Can_Continue("Close Delivery Order") = False Then Exit Sub
    End If
    
    If DriverName = "Change To Delivery" Then
       DriverAssignedToOrder = LoggedOnAs(0).EMPName
       Exit Sub
    End If
    Type_Message = "": Order_Qty = 0
    If GetConfigInfo("Phone Orders", "Prompt Order Close", "No", "FUNCTIONS.INI") = "Yes" Then
      If See_Msg("Pay Order?", 4, "Close Phone Order") <> 6 Then Exit Sub
    End If
    OrderIs = GetOrderType(GoSale)
    Table_Number = GetOrderType(GoSale)
    Number_In_Party = 1
    If DriverName = "Change to Pick Up" Then
      DriverAssignedToOrder = LoggedOnAs(0).EmpLoggedOnAs
    End If
   If System_config = ExpressService Then
     Check_no = Advance_Check(False)
   Else
     Check_no = Assign_Check_No(DriverAssignedToOrder, True)
   End If
   If Check_no = 0 Then
    See_Msg "No Check Number Assigned", 5, "Cannot Close Order"
    Exit Sub
   End If
   If NumberInUse(AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo) = True Then Exit Sub
    ReDim_Array
    ButtonFunction = BTNFNCT.LoadCallIn
    DT_Order_Status = MyOrderStatus(CloseCheck)
    Check_Status = MyOrderStatus(CallIncheck)
    If AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType = "Delivery" Then
     If Input_Option(35) > "" Then
      If Input_Option(35) <> "Any" Then
       If GetConfigInfo("Phone Orders", "Delivery Close By Terminal", "Off", "FUNCTIONS.INI") = "ENABLED" Then
        If MyRegisterName <> Input_Option(35) Then
          See_Msg "Delivery Order Can Only Close On This Station " & Input_Option(35), 5, ""
          Exit Sub
        End If
       End If
      End If
     End If
    End If
    If Load_CallIn_Order(AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo, AllPhoneOrders.Item(SelectionButton(Index).Tag).CallerName) = False Then Exit Sub
    Call MemoryCheck(PhoneOrders)
  Case Is = 5  'edit order
    If LoggedOnAs(0).PhoneAccess Then Exit Sub
    DT_Order_Status = MyOrderStatus(ParkCheck)
    Check_Status = MyOrderStatus(CallIncheck)
    If Load_CallIn_Order(AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo, AllPhoneOrders.Item(SelectionButton(Index).Tag).CallerName) = False Then Exit Sub
    Call MemoryCheck(PhoneOrders)
    Exit Sub
  Case Is = 6
      If AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType = "Reservation" Then
       DTHoldStatus = DT_Order_Status
       CSHoldStatus = Check_Status
       DT_Order_Status = MyOrderStatus(ParkCheck)
       Check_Status = MyOrderStatus(TranferPhone)
       Load_What = "Drivers"
       Find_What = "Assign Check To ?"
       RequestListForm.Show 1
       If GetSelected = "" Then
         DT_Order_Status = DTHoldStatus
         Check_Status = CSHoldStatus
         Exit Sub
       End If
       If GetSelected > "" Then
         AssignToServer GetSelected, AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo
         UnlockedCheckNumber AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo

       Else
         See_Msg "No Server Selected", 5, "Cannot Assign Order."
         AbortSelection = True
         If Check_Option(0) = 0 Then  '// option 0 is for full service
          Check_no = Assign_Check_No(LoggedOnAs(0).EMPName, True)
         Else
          Check_no = Advance_Check(False)
         End If
       Call MemoryCheck(PhoneOrders)
       Exit Sub
      End If
     End If
    AssignDriver AllPhoneOrders.Item(SelectionButton(Index).Tag).OrderType
    Exit Sub
  Case Is = 7
       Dim InDate As String
       Dim OrgDate As String
       Dim PKT As DateOrTime
       Set PKT = New DateOrTime
         OrgDate = AllPhoneOrders.Item(SelectionButton(Index).Tag).PDate
         InDate = PKT.ReturnDate
         If IsDate(InDate) Then
            AllPhoneOrders.Item(SelectionButton(Index).Tag).PDate = Format(InDate, "MM/DD/YYYY")
            Trans_Log "Pick Up Date Edit New Date " & InDate
         End If
         InDate = PKT.ReturnTime
         If IsDate(InDate) Then
             AllPhoneOrders.Item(SelectionButton(Index).Tag).PTime = Format(InDate, "HH:MM AmPm")
             Trans_Log "Pick Up Time Edit New Time " & InDate
         End If
         Set PKT = Nothing
         If IsDate(InDate) Then
          If IsPastEndOfDay(AllPhoneOrders.Item(SelectionButton(Index).Tag).PDate, AllPhoneOrders.Item(SelectionButton(Index).Tag).PTime) Then
           If PaidByCredit(Val(SelectionButton(Index).Tag)).IsPaid Then
              See_Msg "Moving Charge To Deposit", 0, "Credit Card is Attached to Order"
           End If
          End If
          SavePhoneUpdate SelectionButton(Index).Tag, OrgDate
         End If
         Exit Sub
  Case Is = 8
     If AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo = -9999 Then
       See_Msg "Reservations Must Be Transfer To Dining Room", 5, "Transfer Requied"
       Exit Sub
     End If
     DTHoldStatus = DT_Order_Status
     CSHoldStatus = Check_Status
     Check_Status = MyOrderStatus(CallIncheck)
     If Load_CallIn_Order(AllPhoneOrders.Item(SelectionButton(Index).Tag).PickUpNo, AllPhoneOrders.Item(SelectionButton(Index).Tag).CallerName) = False Then Exit Sub
     Order.Amount_Tend_Update
     Receipt "Phone In Order", RegPrint, "", False
     Start_Check
     Order.Gen_Setup
     DT_Order_Status = DTHoldStatus
     Check_Status = CSHoldStatus
     Exit Sub
   Case Is = 10
     If LoggedOnAs(0).PhoneAccess Then Exit Sub
     If Check_Option(62) = 1 Then
       Call SendTextMessage("Confirmation: Your Orders Ready " & LocationName, AllPhoneOrders.Item(SelectionButton(SelectedRow).Tag).CallerPhone)
     Else
       Call AttachPayment(SelectionButton(Index).Tag)
     End If
   Case Else
    'do nothing
  End Select
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Access Phone Orders ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub ShowEdit_Click()
MyInstruction = EditCust
SSCommand15_Click
With EditCustomer
   .Clear
   .FixedRows = 1
   .TextMatrix(0, 0) = "Name"
   .TextMatrix(0, 1) = "Phone"
   .TextMatrix(0, 2) = "Check"
   .TextMatrix(0, 3) = "Payment Type"
   .cols = 4
   .FontSize = 18
   .ColWidth(0) = 3500 * WidScale1024
   .ColWidth(1) = 2500 * WidScale1024
   .ColWidth(2) = 1500 * WidScale1024
   .ColWidth(3) = 2500 * WidScale1024
   .ColAlignment(0) = flexAlignLeftCenter
   .ColAlignment(1) = flexAlignLeftCenter
   .ColAlignment(2) = flexAlignLeftCenter
   .ColAlignment(3) = flexAlignLeftCenter
   .width = (((.ColWidth(0) + .ColWidth(1)) + .ColWidth(2)) + .ColWidth(3)) + 300
   .rows = AllPhoneOrders.count + 1
   ShowOrders.width = .width + 500
End With
Dim F       As Integer
Dim TenderName As String
For F = 1 To AllPhoneOrders.count
    TenderName = DeliveryPrePaidTender(AllPhoneOrders.Item(F).PickUpNo)
    With EditCustomer
      .Cell(flexcpForeColor, F, 0, F, 3) = AllPhoneOrders.Item(F).StatusColor
      .TextMatrix(F, 0) = UCase(AllPhoneOrders.Item(F).CallerName)
      .TextMatrix(F, 1) = AllPhoneOrders.Item(F).CallerPhone
      .TextMatrix(F, 2) = AllPhoneOrders.Item(F).PickUpNo
      .TextMatrix(F, 3) = TenderName
    End With
Next F
DoEvents
CenterObject Me, ShowOrders
ShowOrders.Visible = True
ShowOrders.ZOrder 0
End Sub

Private Sub SSCommand1_Click()
If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
  Sql = "Select * from [Open Checks] Where [presented] = 1 and [Order Destination] = 'Pick Up' and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by [check number] asc'    'TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
Else
  Sql = "Select * from [Open Checks] Where [presented] = 1 and [Order Destination] = 'Pick Up' order by [check number] asc "   'TRY_CAST([Pick Up Time] AS datetime) asc"
End If
Load_CallINs Sql, "Pick Up Orders Only"
CurrentPhoneOrders.Caption = "Pick Up Orders Only" & " Open Orders ( " & AllPhoneOrders.count & " )"
BtnEnh9_Click
TScroll = 1
ScrollButtons
End Sub
Private Sub SSCommand10_Click()

If Can_Continue(SSCommand10.Caption) = False Then Exit Sub
    CurrentPhoneOrders.Caption = ReturnLabel("Phone Orders") & " Open Orders ( " & AllPhoneOrders.count & " )"
    Load_What = "Drivers"
    Find_What = "Select a Driver"
    RequestListForm.Show 1
    If Trim(GetSelected) > "" Then
      DriverMenu.DVRName.Caption = GetSelected
      DriverMenu.Show 1
      GetSelected = ""
    End If
    BtnEnh9_Click
End Sub
Private Sub SSCommand11_Click()
If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
 See_Msg "Access To Low", 5, "No Access"
 Exit Sub
End If
List_Open_Checks
CurrentPhoneOrders.Caption = ReturnLabel("Phone Orders") & " Open Orders ( " & AllPhoneOrders.count & " )"
End Sub
Private Sub SSCommand12_Click()
End Sub
Private Sub SSCommand13_Click()
End Sub

Private Sub SSCommand15_Click()

If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
  Sql = "Select * from [Open Checks] Where [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by  TRY_CAST([Pick Up Date] AS date),TRY_CAST([Pick Up Time] AS datetime)  asc"
Else
  Sql = "Select * from [Open Checks] order by [Check Number] asc " ' TRY_CAST([Pick Up Date] AS date),TRY_CAST([Pick Up Time] AS datetime) asc"
End If
Load_CallINs Sql, ""
TScroll = 1
ScrollButtons
CurrentPhoneOrders.Caption = "All Orders by Open Time" & " Open Orders ( " & AllPhoneOrders.count & " )"
End Sub
Private Sub SSCommand2_Click()

If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
  Sql = "Select * from [Open Checks] Where [presented] = 1 and [Order Destination] = 'Delivery' and [Server Name] = '" & LoggedOnAs(0).EMPName & "'  order by [check number] asc " ' TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
Else
  Sql = "Select * from [Open Checks] Where [presented] = 1 and [Order Destination] = 'Delivery' order by [check number] asc" ' TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
End If
Load_CallINs Sql, "Delivery Orders Only"
CurrentPhoneOrders.Caption = "Delivery Orders Only" & " Open Orders ( " & AllPhoneOrders.count & " )"
BtnEnh9_Click
TScroll = 1
ScrollButtons
Exit Sub
ermsg:
Error_Message "Phone Orders SS2", Err.Description
Resume Next
End Sub
Private Sub SSCommand3_Click()

If LoggedOnAs(0).EmpAccessLevel < Access_Required("Access Phone Orders") Then
   Sql = "Select * from [Open Checks] Where [Presented] = 1 and [Order Destination] = 'Pick Up' and [Server Name] = '" & LoggedOnAs(0).EMPName & "' order by TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
Else
   Sql = "Select * from [Open Checks] Where [Presented] = 1 and [Order Destination] = 'Pick Up' order by TRY_CAST([Pick Up Date] AS date), TRY_CAST([Pick Up Time] AS datetime) asc"
End If
Load_CallINs Sql, ""
TScroll = 1
ScrollButtons
CurrentPhoneOrders.Caption = "Pick Up Orders By Time Opened" & " Open Orders ( " & AllPhoneOrders.count & " )"
BtnEnh9_Click
End Sub
Private Sub SSCommand4_Click()
If Can_Continue("AssignDriver") = False Then Exit Sub
CurrentPhoneOrders.Caption = ReturnLabel("Phone Orders") & " Open Orders ( " & AllPhoneOrders.count & " )"
AssignDrivers = True
ClosePhone.Show 1
AssignDrivers = False
Form_Load
End Sub
Private Sub SSCommand5_Click()
If Can_Continue(SSCommand5.Caption) = False Then Exit Sub
ClosePhone.Show 1
Call MemoryCheck(ClosePhone)
Form_Load
End Sub
Private Sub SSCommand6_Click()
On Error Resume Next
If OrderIsFor.Caption = "Delivery" Then
  If CInfo(1).Text = "" Then
     If See_Msg("Address required - Quit Change? ", 4, "cannot change order without address") = 6 Then
      CustomerINfo.Visible = False
      Exit Sub
     End If
  End If
End If
Dim OrgDate As String
OrgDate = Format(PUDATE.Text, "MM/DD/YYYY")
If IsDate(OrgDate) Then

Else
 OrgDate = ""
End If
'Dim DriveInfo   As DrvZoneData
'DriveInfo = GetDelivery(CInfo(1).Text, CInfo(2).Text, CInfo(3).Text, CInfo(4).Text)
'AllPhoneOrders.Item(WorkingCheckNo).Zone = DriveInfo.ZoneName
'AllPhoneOrders.Item(WorkingCheckNo).Distance = DriveInfo.ZoneDistance
SavePhoneUpdate CStr(WorkingCheckNo), OrgDate
UpdateCustomer CInfo(0).Text, CInfo(5).Text, CStr(WorkingCheckNo), Format(CInfo(7).Text, "0.00")

CustomerINfo.Visible = False
CustomerINfo.Left = -5000
DepositAmtPaid = 0
End Sub
Private Sub Up_Arrow_Click()
TScroll = TScroll - 15  '  HowmanyButtons
ScrollButtons
End Sub
Private Function HowmanyButtons() As Integer
HowmanyButtons = AllPhoneOrders.count
End Function
Private Sub WatchingforNewDeliveries_Timer()
SearchHow_Click 5
End Sub
