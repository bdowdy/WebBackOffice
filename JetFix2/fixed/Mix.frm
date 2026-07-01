VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Sales_Mix 
   Appearance      =   0  'Flat
   BackColor       =   &H00302821&
   BorderStyle     =   0  'None
   Caption         =   "Sales Mix"
   ClientHeight    =   10545
   ClientLeft      =   180
   ClientTop       =   450
   ClientWidth     =   14835
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
   ScaleHeight     =   10545
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MixByDept 
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
      Height          =   6615
      Left            =   13320
      ScaleHeight     =   6585
      ScaleWidth      =   5865
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ListBox MixList 
         Appearance      =   0  'Flat
         Height          =   5070
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   735
         Left            =   2880
         TabIndex        =   18
         Top             =   5520
         Width           =   2775
         _Version        =   589828
         _ExtentX        =   4895
         _ExtentY        =   1296
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
         BackColorContainer=   4210752
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":0000
         textLT          =   "Mix.frx":0068
         textCT          =   "Mix.frx":0080
         textRT          =   "Mix.frx":0098
         textLM          =   "Mix.frx":00B0
         textRM          =   "Mix.frx":00C8
         textLB          =   "Mix.frx":00E0
         textCB          =   "Mix.frx":00F8
         textRB          =   "Mix.frx":0110
         colorBack       =   "Mix.frx":0128
         colorIntern     =   "Mix.frx":0152
         colorMO         =   "Mix.frx":017C
         colorFocus      =   "Mix.frx":01A6
         colorDisabled   =   "Mix.frx":01D0
         colorPressed    =   "Mix.frx":01FA
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
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   615
         Left            =   3000
         TabIndex        =   20
         Top             =   3960
         Width           =   2415
         _Version        =   589828
         _ExtentX        =   4260
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Select Range"
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
         BackColorContainer=   8454016
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":0224
         textLT          =   "Mix.frx":029C
         textCT          =   "Mix.frx":02B4
         textRT          =   "Mix.frx":02CC
         textLM          =   "Mix.frx":02E4
         textRM          =   "Mix.frx":02FC
         textLB          =   "Mix.frx":0314
         textCB          =   "Mix.frx":032C
         textRB          =   "Mix.frx":0344
         colorBack       =   "Mix.frx":035C
         colorIntern     =   "Mix.frx":0386
         colorMO         =   "Mix.frx":03B0
         colorFocus      =   "Mix.frx":03DA
         colorDisabled   =   "Mix.frx":0404
         colorPressed    =   "Mix.frx":042E
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
         Office2003ColorNorth=   8421504
         Office2003ColorSouth=   8421504
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   8421504
         Office2003ColorSouthMouseOver=   12632256
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   8421504
         Office2003ColorSouthPressed=   8421504
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh10 
         Height          =   735
         Left            =   240
         TabIndex        =   29
         Top             =   5520
         Width           =   2535
         _Version        =   589828
         _ExtentX        =   4471
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Select All"
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
         textCaption     =   "Mix.frx":0458
         textLT          =   "Mix.frx":04CC
         textCT          =   "Mix.frx":04E4
         textRT          =   "Mix.frx":04FC
         textLM          =   "Mix.frx":0514
         textRM          =   "Mix.frx":052C
         textLB          =   "Mix.frx":0544
         textCB          =   "Mix.frx":055C
         textRB          =   "Mix.frx":0574
         colorBack       =   "Mix.frx":058C
         colorIntern     =   "Mix.frx":05B6
         colorMO         =   "Mix.frx":05E0
         colorFocus      =   "Mix.frx":060A
         colorDisabled   =   "Mix.frx":0634
         colorPressed    =   "Mix.frx":065E
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
         Office2003ColorNorthPressed=   3158064
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh11 
         Height          =   615
         Left            =   3000
         TabIndex        =   31
         Top             =   4680
         Width           =   2415
         _Version        =   589828
         _ExtentX        =   4260
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Today"
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
         BackColorContainer=   8454016
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":0688
         textLT          =   "Mix.frx":06F2
         textCT          =   "Mix.frx":070A
         textRT          =   "Mix.frx":0722
         textLM          =   "Mix.frx":073A
         textRM          =   "Mix.frx":0752
         textLB          =   "Mix.frx":076A
         textCB          =   "Mix.frx":0782
         textRB          =   "Mix.frx":079A
         colorBack       =   "Mix.frx":07B2
         colorIntern     =   "Mix.frx":07DC
         colorMO         =   "Mix.frx":0806
         colorFocus      =   "Mix.frx":0830
         colorDisabled   =   "Mix.frx":085A
         colorPressed    =   "Mix.frx":0884
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
         Office2003ColorNorth=   8421504
         Office2003ColorSouth=   8421504
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   12632256
         Office2003ColorSouthMouseOver=   3767430
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   8421504
         Office2003ColorSouthPressed=   4484987
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh20 
         Height          =   615
         Left            =   3000
         TabIndex        =   48
         Top             =   720
         Width           =   2415
         _Version        =   589828
         _ExtentX        =   4260
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Mix By Department"
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
         BackColorContainer=   8454016
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":08AE
         textLT          =   "Mix.frx":0930
         textCT          =   "Mix.frx":0948
         textRT          =   "Mix.frx":0960
         textLM          =   "Mix.frx":0978
         textRM          =   "Mix.frx":0990
         textLB          =   "Mix.frx":09A8
         textCB          =   "Mix.frx":09C0
         textRB          =   "Mix.frx":09D8
         colorBack       =   "Mix.frx":09F0
         colorIntern     =   "Mix.frx":0A1A
         colorMO         =   "Mix.frx":0A44
         colorFocus      =   "Mix.frx":0A6E
         colorDisabled   =   "Mix.frx":0A98
         colorPressed    =   "Mix.frx":0AC2
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
      Begin BTNENHLib4.BtnEnh BtnEnh21 
         Height          =   615
         Left            =   3000
         TabIndex        =   49
         Top             =   1440
         Width           =   2415
         _Version        =   589828
         _ExtentX        =   4260
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Mix By Category"
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
         BackColorContainer=   8454016
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":0AEC
         textLT          =   "Mix.frx":0B6A
         textCT          =   "Mix.frx":0B82
         textRT          =   "Mix.frx":0B9A
         textLM          =   "Mix.frx":0BB2
         textRM          =   "Mix.frx":0BCA
         textLB          =   "Mix.frx":0BE2
         textCB          =   "Mix.frx":0BFA
         textRB          =   "Mix.frx":0C12
         colorBack       =   "Mix.frx":0C2A
         colorIntern     =   "Mix.frx":0C54
         colorMO         =   "Mix.frx":0C7E
         colorFocus      =   "Mix.frx":0CA8
         colorDisabled   =   "Mix.frx":0CD2
         colorPressed    =   "Mix.frx":0CFC
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
      Begin VB.Label RType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Departments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date Range"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   21
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   2880
         Top             =   240
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   6375
         Left            =   120
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.PictureBox bergbox 
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
      Height          =   4815
      Left            =   5520
      ScaleHeight     =   4785
      ScaleWidth      =   7185
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox ICost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   46
         Text            =   "0.00"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox MenuItem 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   480
         TabIndex        =   41
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox ItemList 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   480
         TabIndex        =   34
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton SSCommand18 
         Caption         =   "Delete Item"
         Height          =   350
         Left            =   480
         TabIndex        =   33
         Top             =   3120
         Width           =   1215
      End
      Begin VSFlex8DAOCtl.VSFlexGrid Berg 
         Height          =   2895
         Left            =   3480
         TabIndex        =   35
         Top             =   480
         Width           =   3375
         _cx             =   2002655041
         _cy             =   2002654194
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16745216
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
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
         Cols            =   2
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
         Editable        =   1
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
      Begin VB.CommandButton SSCommand19 
         Caption         =   "Save Item"
         Height          =   350
         Left            =   1920
         TabIndex        =   36
         Top             =   3120
         Width           =   1215
      End
      Begin BTNENHLib4.BtnEnh BtnEnh15 
         Height          =   735
         Left            =   2160
         TabIndex        =   42
         Top             =   3840
         Width           =   1815
         _Version        =   589828
         _ExtentX        =   3201
         _ExtentY        =   1296
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
         Surface         =   15
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":0D26
         textLT          =   "Mix.frx":0D8E
         textCT          =   "Mix.frx":0DA6
         textRT          =   "Mix.frx":0DBE
         textLM          =   "Mix.frx":0DD6
         textRM          =   "Mix.frx":0DEE
         textLB          =   "Mix.frx":0E06
         textCB          =   "Mix.frx":0E1E
         textRB          =   "Mix.frx":0E36
         colorBack       =   "Mix.frx":0E4E
         colorIntern     =   "Mix.frx":0E78
         colorMO         =   "Mix.frx":0EA2
         colorFocus      =   "Mix.frx":0ECC
         colorDisabled   =   "Mix.frx":0EF6
         colorPressed    =   "Mix.frx":0F20
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
      Begin BTNENHLib4.BtnEnh BtnEnh16 
         Height          =   735
         Left            =   240
         TabIndex        =   43
         Top             =   3840
         Width           =   1815
         _Version        =   589828
         _ExtentX        =   3201
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Delete"
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
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":0F4A
         textLT          =   "Mix.frx":0FB6
         textCT          =   "Mix.frx":0FCE
         textRT          =   "Mix.frx":0FE6
         textLM          =   "Mix.frx":0FFE
         textRM          =   "Mix.frx":1016
         textLB          =   "Mix.frx":102E
         textCB          =   "Mix.frx":1046
         textRB          =   "Mix.frx":105E
         colorBack       =   "Mix.frx":1076
         colorIntern     =   "Mix.frx":10A0
         colorMO         =   "Mix.frx":10CA
         colorFocus      =   "Mix.frx":10F4
         colorDisabled   =   "Mix.frx":111E
         colorPressed    =   "Mix.frx":1148
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
         Office2003ColorNorth=   179
         Office2003ColorSouth=   117
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   179
         Office2003ColorSouthMouseOver=   128
         Office2003ColorBorderMouseOver=   1812708
         Office2003ColorNorthPressed=   179
         Office2003ColorSouthPressed=   179
         Office2003ColorBorderPressed=   1812708
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh17 
         Height          =   735
         Left            =   4680
         TabIndex        =   44
         Top             =   3840
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
         Surface         =   15
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":1172
         textLT          =   "Mix.frx":11DA
         textCT          =   "Mix.frx":11F2
         textRT          =   "Mix.frx":120A
         textLM          =   "Mix.frx":1222
         textRM          =   "Mix.frx":123A
         textLB          =   "Mix.frx":1252
         textCB          =   "Mix.frx":126A
         textRB          =   "Mix.frx":1282
         colorBack       =   "Mix.frx":129A
         colorIntern     =   "Mix.frx":12C4
         colorMO         =   "Mix.frx":12EE
         colorFocus      =   "Mix.frx":1318
         colorDisabled   =   "Mix.frx":1342
         colorPressed    =   "Mix.frx":136C
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
      Begin VB.CommandButton SSCommand1 
         Caption         =   "Load Menu From Mix"
         Height          =   350
         Left            =   480
         TabIndex        =   47
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton SSCommand3 
         Caption         =   "Load Menu From Menu File"
         Height          =   350
         Left            =   480
         TabIndex        =   40
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Cost"
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Item Ingredient List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Create Inventory Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Menu Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3255
         Index           =   1
         Left            =   3360
         Top             =   240
         Width           =   3615
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   360
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   360
         Top             =   240
         Width           =   2895
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   3495
         Left            =   120
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.PictureBox MenuSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Height          =   8895
      Left            =   -5640
      ScaleHeight     =   8865
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   -4320
      Visible         =   0   'False
      Width           =   10095
      Begin VB.ComboBox Cats 
         Appearance      =   0  'Flat
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "None"
         Top             =   8160
         Width           =   2655
      End
      Begin VB.CheckBox IndSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Group Menu Items Sold by Employee"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6120
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   495
         Left            =   6240
         TabIndex        =   9
         Top             =   4680
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   873
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
         Surface         =   15
         BackColorContainer=   16744576
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":1396
         textLT          =   "Mix.frx":13FE
         textCT          =   "Mix.frx":1416
         textRT          =   "Mix.frx":142E
         textLM          =   "Mix.frx":1446
         textRM          =   "Mix.frx":145E
         textLB          =   "Mix.frx":1476
         textCB          =   "Mix.frx":148E
         textRB          =   "Mix.frx":14A6
         colorBack       =   "Mix.frx":14BE
         colorIntern     =   "Mix.frx":14E8
         colorMO         =   "Mix.frx":1512
         colorFocus      =   "Mix.frx":153C
         colorDisabled   =   "Mix.frx":1566
         colorPressed    =   "Mix.frx":1590
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
      Begin VB.ComboBox ContestName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "Mix.frx":15BA
         Left            =   6240
         List            =   "Mix.frx":15BC
         TabIndex        =   8
         Text            =   "None"
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox ETime 
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
         Height          =   285
         Left            =   6120
         TabIndex        =   5
         Text            =   "None"
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox STime 
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
         Height          =   285
         Left            =   6120
         TabIndex        =   4
         Text            =   "None"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.ListBox Sel_Server 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5730
         Left            =   1080
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   1200
         Width           =   4575
      End
      Begin BTNENHLib4.BtnEnh SSCommand9 
         Height          =   855
         Left            =   6000
         TabIndex        =   6
         Top             =   7680
         Width           =   3855
         _Version        =   589828
         _ExtentX        =   6800
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
         BackColorContainer=   8421504
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":15BE
         textLT          =   "Mix.frx":1626
         textCT          =   "Mix.frx":163E
         textRT          =   "Mix.frx":1656
         textLM          =   "Mix.frx":166E
         textRM          =   "Mix.frx":1686
         textLB          =   "Mix.frx":169E
         textCB          =   "Mix.frx":16B6
         textRB          =   "Mix.frx":16CE
         colorBack       =   "Mix.frx":16E6
         colorIntern     =   "Mix.frx":1710
         colorMO         =   "Mix.frx":173A
         colorFocus      =   "Mix.frx":1764
         colorDisabled   =   "Mix.frx":178E
         colorPressed    =   "Mix.frx":17B8
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
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   735
         Left            =   3400
         TabIndex        =   10
         Top             =   240
         Width           =   2250
         _Version        =   589828
         _ExtentX        =   3969
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Load Menu Items from Menu Database"
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
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":17E2
         textLT          =   "Mix.frx":1886
         textCT          =   "Mix.frx":189E
         textRT          =   "Mix.frx":18B6
         textLM          =   "Mix.frx":18CE
         textRM          =   "Mix.frx":18E6
         textLB          =   "Mix.frx":18FE
         textCB          =   "Mix.frx":1916
         textRB          =   "Mix.frx":192E
         colorBack       =   "Mix.frx":1946
         colorIntern     =   "Mix.frx":1970
         colorMO         =   "Mix.frx":199A
         colorFocus      =   "Mix.frx":19C4
         colorDisabled   =   "Mix.frx":19EE
         colorPressed    =   "Mix.frx":1A18
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
      Begin BTNENHLib4.BtnEnh Get_Day 
         Height          =   1455
         Left            =   8040
         TabIndex        =   12
         Top             =   5880
         Width           =   1815
         _Version        =   589828
         _ExtentX        =   3201
         _ExtentY        =   2566
         _StockProps     =   66
         Caption         =   "Select Date Range"
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
         textCaption     =   "Mix.frx":1A42
         textLT          =   "Mix.frx":1AC4
         textCT          =   "Mix.frx":1ADC
         textRT          =   "Mix.frx":1AF4
         textLM          =   "Mix.frx":1B0C
         textRM          =   "Mix.frx":1B24
         textLB          =   "Mix.frx":1B3C
         textCB          =   "Mix.frx":1B54
         textRB          =   "Mix.frx":1B6C
         colorBack       =   "Mix.frx":1B84
         colorIntern     =   "Mix.frx":1BAE
         colorMO         =   "Mix.frx":1BD8
         colorFocus      =   "Mix.frx":1C02
         colorDisabled   =   "Mix.frx":1C2C
         colorPressed    =   "Mix.frx":1C56
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
      Begin BTNENHLib4.BtnEnh Arrow_Up 
         Height          =   855
         Left            =   175
         TabIndex        =   13
         Top             =   4650
         Width           =   855
         _Version        =   589828
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "PG UP"
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
         Picture         =   "C:\Icons\ARROWS\POINT05.ICO"
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "Mix.frx":1C80
         textCaption     =   "Mix.frx":1C98
         textLT          =   "Mix.frx":1D02
         textCT          =   "Mix.frx":1D1A
         textRT          =   "Mix.frx":1D32
         textLM          =   "Mix.frx":1D4A
         textRM          =   "Mix.frx":1D62
         textLB          =   "Mix.frx":1D7A
         textCB          =   "Mix.frx":1D92
         textRB          =   "Mix.frx":1DAA
         colorBack       =   "Mix.frx":1DC2
         colorIntern     =   "Mix.frx":1DEC
         colorMO         =   "Mix.frx":1E16
         colorFocus      =   "Mix.frx":1E40
         colorDisabled   =   "Mix.frx":1E6A
         colorPressed    =   "Mix.frx":1E94
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
      Begin BTNENHLib4.BtnEnh Arrow_Dwn 
         Height          =   855
         Left            =   175
         TabIndex        =   14
         Top             =   5520
         Width           =   855
         _Version        =   589828
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "PG DOWN"
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
         Picture         =   "C:\Icons\ARROWS\POINT03.ICO"
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "Mix.frx":1EBE
         textCaption     =   "Mix.frx":1ED6
         textLT          =   "Mix.frx":1F44
         textCT          =   "Mix.frx":1F5C
         textRT          =   "Mix.frx":1F74
         textLM          =   "Mix.frx":1F8C
         textRM          =   "Mix.frx":1FA4
         textLB          =   "Mix.frx":1FBC
         textCB          =   "Mix.frx":1FD4
         textRB          =   "Mix.frx":1FEC
         colorBack       =   "Mix.frx":2004
         colorIntern     =   "Mix.frx":202E
         colorMO         =   "Mix.frx":2058
         colorFocus      =   "Mix.frx":2082
         colorDisabled   =   "Mix.frx":20AC
         colorPressed    =   "Mix.frx":20D6
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
      Begin BTNENHLib4.BtnEnh SSCommand6 
         Height          =   615
         Left            =   3400
         TabIndex        =   15
         Top             =   6960
         Width           =   2250
         _Version        =   589828
         _ExtentX        =   3969
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Clear All"
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
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":2100
         textLT          =   "Mix.frx":2172
         textCT          =   "Mix.frx":218A
         textRT          =   "Mix.frx":21A2
         textLM          =   "Mix.frx":21BA
         textRM          =   "Mix.frx":21D2
         textLB          =   "Mix.frx":21EA
         textCB          =   "Mix.frx":2202
         textRB          =   "Mix.frx":221A
         colorBack       =   "Mix.frx":2232
         colorIntern     =   "Mix.frx":225C
         colorMO         =   "Mix.frx":2286
         colorFocus      =   "Mix.frx":22B0
         colorDisabled   =   "Mix.frx":22DA
         colorPressed    =   "Mix.frx":2304
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
      Begin BTNENHLib4.BtnEnh SSCommand4 
         Height          =   615
         Left            =   1080
         TabIndex        =   16
         Top             =   6960
         Width           =   2250
         _Version        =   589828
         _ExtentX        =   3969
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Select All"
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
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":232E
         textLT          =   "Mix.frx":23A2
         textCT          =   "Mix.frx":23BA
         textRT          =   "Mix.frx":23D2
         textLM          =   "Mix.frx":23EA
         textRM          =   "Mix.frx":2402
         textLB          =   "Mix.frx":241A
         textCB          =   "Mix.frx":2432
         textRB          =   "Mix.frx":244A
         colorBack       =   "Mix.frx":2462
         colorIntern     =   "Mix.frx":248C
         colorMO         =   "Mix.frx":24B6
         colorFocus      =   "Mix.frx":24E0
         colorDisabled   =   "Mix.frx":250A
         colorPressed    =   "Mix.frx":2534
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
         Office2003ColorNorthPressed=   3158064
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh6 
         Height          =   375
         Left            =   3840
         TabIndex        =   25
         Top             =   8160
         Width           =   1815
         _Version        =   589828
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "Select Category"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   15
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":255E
         textLT          =   "Mix.frx":25DC
         textCT          =   "Mix.frx":25F4
         textRT          =   "Mix.frx":260C
         textLM          =   "Mix.frx":2624
         textRM          =   "Mix.frx":263C
         textLB          =   "Mix.frx":2654
         textCB          =   "Mix.frx":266C
         textRB          =   "Mix.frx":2684
         colorBack       =   "Mix.frx":269C
         colorIntern     =   "Mix.frx":26C6
         colorMO         =   "Mix.frx":26F0
         colorFocus      =   "Mix.frx":271A
         colorDisabled   =   "Mix.frx":2744
         colorPressed    =   "Mix.frx":276E
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
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   495
         Left            =   7680
         TabIndex        =   26
         Top             =   4680
         Width           =   1455
         _Version        =   589828
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   66
         Caption         =   "Delete"
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
         textCaption     =   "Mix.frx":2798
         textLT          =   "Mix.frx":2804
         textCT          =   "Mix.frx":281C
         textRT          =   "Mix.frx":2834
         textLM          =   "Mix.frx":284C
         textRM          =   "Mix.frx":2864
         textLB          =   "Mix.frx":287C
         textCB          =   "Mix.frx":2894
         textRB          =   "Mix.frx":28AC
         colorBack       =   "Mix.frx":28C4
         colorIntern     =   "Mix.frx":28EE
         colorMO         =   "Mix.frx":2918
         colorFocus      =   "Mix.frx":2942
         colorDisabled   =   "Mix.frx":296C
         colorPressed    =   "Mix.frx":2996
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
         Office2003ColorNorth=   179
         Office2003ColorSouth=   117
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   179
         Office2003ColorSouthMouseOver=   128
         Office2003ColorBorderMouseOver=   1812708
         Office2003ColorNorthPressed=   179
         Office2003ColorSouthPressed=   179
         Office2003ColorBorderPressed=   1812708
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh8 
         Height          =   735
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   2250
         _Version        =   589828
         _ExtentX        =   3969
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   "Load Sold Menu Items from Sales History"
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
         BackColorContainer=   8438015
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":29C0
         textLT          =   "Mix.frx":2A6E
         textCT          =   "Mix.frx":2A86
         textRT          =   "Mix.frx":2A9E
         textLM          =   "Mix.frx":2AB6
         textRM          =   "Mix.frx":2ACE
         textLB          =   "Mix.frx":2AE6
         textCB          =   "Mix.frx":2AFE
         textRB          =   "Mix.frx":2B16
         colorBack       =   "Mix.frx":2B2E
         colorIntern     =   "Mix.frx":2B58
         colorMO         =   "Mix.frx":2B82
         colorFocus      =   "Mix.frx":2BAC
         colorDisabled   =   "Mix.frx":2BD6
         colorPressed    =   "Mix.frx":2C00
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
         Office2003ColorNorth=   10485760
         Office2003ColorSouth=   4194304
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   10485760
         Office2003ColorSouthMouseOver=   10485760
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   10485760
         Office2003ColorSouthPressed=   4194304
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand8 
         Height          =   1455
         Left            =   6000
         TabIndex        =   28
         Top             =   5880
         Width           =   1815
         _Version        =   589828
         _ExtentX        =   3201
         _ExtentY        =   2566
         _StockProps     =   66
         Caption         =   "Today"
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
         textCaption     =   "Mix.frx":2C2A
         textLT          =   "Mix.frx":2C94
         textCT          =   "Mix.frx":2CAC
         textRT          =   "Mix.frx":2CC4
         textLM          =   "Mix.frx":2CDC
         textRM          =   "Mix.frx":2CF4
         textLB          =   "Mix.frx":2D0C
         textCB          =   "Mix.frx":2D24
         textRB          =   "Mix.frx":2D3C
         colorBack       =   "Mix.frx":2D54
         colorIntern     =   "Mix.frx":2D7E
         colorMO         =   "Mix.frx":2DA8
         colorFocus      =   "Mix.frx":2DD2
         colorDisabled   =   "Mix.frx":2DFC
         colorPressed    =   "Mix.frx":2E26
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
      Begin VB.CheckBox GroupBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Group Menu Items and Modifiers with Same Name Together"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6120
         TabIndex        =   30
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Filter Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   55
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Select By Individual Category"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   23
         Top             =   7800
         Width           =   2895
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Mix Report Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   6220
         TabIndex        =   7
         Top             =   3720
         Width           =   4095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End Time (Click Label Sets To None)"
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
         Left            =   6120
         TabIndex        =   3
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Start Time (Click Label Sets To None)"
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
         Left            =   6120
         TabIndex        =   2
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   8655
         Left            =   120
         Top             =   120
         Width           =   5655
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   3255
         Left            =   5880
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2055
         Index           =   0
         Left            =   5880
         Top             =   3480
         Width           =   4095
      End
   End
   Begin VB.Frame Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   3840
      TabIndex        =   50
      Top             =   2400
      Width           =   6975
      Begin BTNENHLib4.BtnEnh SSCommand5 
         Height          =   840
         Left            =   360
         TabIndex        =   51
         Top             =   2280
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Menu Item(s) by Individual"
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
         textCaption     =   "Mix.frx":2E50
         textLT          =   "Mix.frx":2EE4
         textCT          =   "Mix.frx":2EFC
         textRT          =   "Mix.frx":2F14
         textLM          =   "Mix.frx":2F2C
         textRM          =   "Mix.frx":2F44
         textLB          =   "Mix.frx":2F5C
         textCB          =   "Mix.frx":2F74
         textRB          =   "Mix.frx":2F8C
         colorBack       =   "Mix.frx":2FA4
         colorIntern     =   "Mix.frx":2FCE
         colorMO         =   "Mix.frx":2FF8
         colorFocus      =   "Mix.frx":3022
         colorDisabled   =   "Mix.frx":304C
         colorPressed    =   "Mix.frx":3076
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
      Begin BTNENHLib4.BtnEnh BtnEnh13 
         Height          =   840
         Left            =   300
         TabIndex        =   52
         Top             =   360
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Daily Menu Inventory"
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
         textCaption     =   "Mix.frx":30A0
         textLT          =   "Mix.frx":3128
         textCT          =   "Mix.frx":3140
         textRT          =   "Mix.frx":3158
         textLM          =   "Mix.frx":3170
         textRM          =   "Mix.frx":3188
         textLB          =   "Mix.frx":31A0
         textCB          =   "Mix.frx":31B8
         textRB          =   "Mix.frx":31D0
         colorBack       =   "Mix.frx":31E8
         colorIntern     =   "Mix.frx":3212
         colorMO         =   "Mix.frx":323C
         colorFocus      =   "Mix.frx":3266
         colorDisabled   =   "Mix.frx":3290
         colorPressed    =   "Mix.frx":32BA
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
      Begin BTNENHLib4.BtnEnh BtnEnh14 
         Height          =   840
         Left            =   360
         TabIndex        =   53
         Top             =   3240
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Inventory Usage"
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
         textCaption     =   "Mix.frx":32E4
         textLT          =   "Mix.frx":3362
         textCT          =   "Mix.frx":337A
         textRT          =   "Mix.frx":3392
         textLM          =   "Mix.frx":33AA
         textRM          =   "Mix.frx":33C2
         textLB          =   "Mix.frx":33DA
         textCB          =   "Mix.frx":33F2
         textRB          =   "Mix.frx":340A
         colorBack       =   "Mix.frx":3422
         colorIntern     =   "Mix.frx":344C
         colorMO         =   "Mix.frx":3476
         colorFocus      =   "Mix.frx":34A0
         colorDisabled   =   "Mix.frx":34CA
         colorPressed    =   "Mix.frx":34F4
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
         Height          =   840
         Left            =   360
         TabIndex        =   54
         Top             =   4200
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Inventory Usage Setup"
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
         textCaption     =   "Mix.frx":351E
         textLT          =   "Mix.frx":35A8
         textCT          =   "Mix.frx":35C0
         textRT          =   "Mix.frx":35D8
         textLM          =   "Mix.frx":35F0
         textRM          =   "Mix.frx":3608
         textLB          =   "Mix.frx":3620
         textCB          =   "Mix.frx":3638
         textRB          =   "Mix.frx":3650
         colorBack       =   "Mix.frx":3668
         colorIntern     =   "Mix.frx":3692
         colorMO         =   "Mix.frx":36BC
         colorFocus      =   "Mix.frx":36E6
         colorDisabled   =   "Mix.frx":3710
         colorPressed    =   "Mix.frx":373A
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
         Left            =   360
         TabIndex        =   56
         Top             =   5160
         Width           =   6255
         _Version        =   589828
         _ExtentX        =   11033
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
         BackColorContainer=   8421504
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Mix.frx":3764
         textLT          =   "Mix.frx":37CC
         textCT          =   "Mix.frx":37E4
         textRT          =   "Mix.frx":37FC
         textLM          =   "Mix.frx":3814
         textRM          =   "Mix.frx":382C
         textLB          =   "Mix.frx":3844
         textCB          =   "Mix.frx":385C
         textRB          =   "Mix.frx":3874
         colorBack       =   "Mix.frx":388C
         colorIntern     =   "Mix.frx":38B6
         colorMO         =   "Mix.frx":38E0
         colorFocus      =   "Mix.frx":390A
         colorDisabled   =   "Mix.frx":3934
         colorPressed    =   "Mix.frx":395E
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
      Begin BTNENHLib4.BtnEnh BtnEnh22 
         Height          =   840
         Left            =   360
         TabIndex        =   57
         Top             =   1320
         Width           =   3015
         _Version        =   589828
         _ExtentX        =   5318
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Export Sales Mix"
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
         textCaption     =   "Mix.frx":3988
         textLT          =   "Mix.frx":3A08
         textCT          =   "Mix.frx":3A20
         textRT          =   "Mix.frx":3A38
         textLM          =   "Mix.frx":3A50
         textRM          =   "Mix.frx":3A68
         textLB          =   "Mix.frx":3A80
         textCB          =   "Mix.frx":3A98
         textRB          =   "Mix.frx":3AB0
         colorBack       =   "Mix.frx":3AC8
         colorIntern     =   "Mix.frx":3AF2
         colorMO         =   "Mix.frx":3B1C
         colorFocus      =   "Mix.frx":3B46
         colorDisabled   =   "Mix.frx":3B70
         colorPressed    =   "Mix.frx":3B9A
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
      Begin BTNENHLib4.BtnEnh PrintStoreMenu 
         Height          =   855
         Left            =   3480
         TabIndex        =   58
         Top             =   360
         Width           =   3135
         _Version        =   589828
         _ExtentX        =   5530
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Store Menu"
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
         textCaption     =   "Mix.frx":3BC4
         textLT          =   "Mix.frx":3C44
         textCT          =   "Mix.frx":3C5C
         textRT          =   "Mix.frx":3C74
         textLM          =   "Mix.frx":3C8C
         textRM          =   "Mix.frx":3CA4
         textLB          =   "Mix.frx":3CBC
         textCB          =   "Mix.frx":3CD4
         textRB          =   "Mix.frx":3CEC
         colorBack       =   "Mix.frx":3D04
         colorIntern     =   "Mix.frx":3D2E
         colorMO         =   "Mix.frx":3D58
         colorFocus      =   "Mix.frx":3D82
         colorDisabled   =   "Mix.frx":3DAC
         colorPressed    =   "Mix.frx":3DD6
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
      Begin BTNENHLib4.BtnEnh PrintCountDownItems 
         Height          =   855
         Left            =   3480
         TabIndex        =   59
         Top             =   1320
         Width           =   3135
         _Version        =   589828
         _ExtentX        =   5530
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Count Down Items"
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
         textCaption     =   "Mix.frx":3E00
         textLT          =   "Mix.frx":3E8C
         textCT          =   "Mix.frx":3EA4
         textRT          =   "Mix.frx":3EBC
         textLM          =   "Mix.frx":3ED4
         textRM          =   "Mix.frx":3EEC
         textLB          =   "Mix.frx":3F04
         textCB          =   "Mix.frx":3F1C
         textRB          =   "Mix.frx":3F34
         colorBack       =   "Mix.frx":3F4C
         colorIntern     =   "Mix.frx":3F76
         colorMO         =   "Mix.frx":3FA0
         colorFocus      =   "Mix.frx":3FCA
         colorDisabled   =   "Mix.frx":3FF4
         colorPressed    =   "Mix.frx":401E
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
      Begin BTNENHLib4.BtnEnh PrintModifiers 
         Height          =   825
         Left            =   3480
         TabIndex        =   60
         Top             =   3120
         Width           =   3135
         _Version        =   589828
         _ExtentX        =   5530
         _ExtentY        =   1455
         _StockProps     =   66
         Caption         =   "Print Modifiers"
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
         textCaption     =   "Mix.frx":4048
         textLT          =   "Mix.frx":40C6
         textCT          =   "Mix.frx":40DE
         textRT          =   "Mix.frx":40F6
         textLM          =   "Mix.frx":410E
         textRM          =   "Mix.frx":4126
         textLB          =   "Mix.frx":413E
         textCB          =   "Mix.frx":4156
         textRB          =   "Mix.frx":416E
         colorBack       =   "Mix.frx":4186
         colorIntern     =   "Mix.frx":41B0
         colorMO         =   "Mix.frx":41DA
         colorFocus      =   "Mix.frx":4204
         colorDisabled   =   "Mix.frx":422E
         colorPressed    =   "Mix.frx":4258
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
      Begin BTNENHLib4.BtnEnh ShowModifierList 
         Height          =   825
         Left            =   3480
         TabIndex        =   61
         Top             =   2280
         Width           =   3135
         _Version        =   589828
         _ExtentX        =   5530
         _ExtentY        =   1455
         _StockProps     =   66
         Caption         =   "View Modifier List"
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
         textCaption     =   "Mix.frx":4282
         textLT          =   "Mix.frx":4306
         textCT          =   "Mix.frx":431E
         textRT          =   "Mix.frx":4336
         textLM          =   "Mix.frx":434E
         textRM          =   "Mix.frx":4366
         textLB          =   "Mix.frx":437E
         textCB          =   "Mix.frx":4396
         textRB          =   "Mix.frx":43AE
         colorBack       =   "Mix.frx":43C6
         colorIntern     =   "Mix.frx":43F0
         colorMO         =   "Mix.frx":441A
         colorFocus      =   "Mix.frx":4444
         colorDisabled   =   "Mix.frx":446E
         colorPressed    =   "Mix.frx":4498
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
      Begin BTNENHLib4.BtnEnh LiquorPluList 
         Height          =   825
         Left            =   3480
         TabIndex        =   62
         Top             =   4200
         Width           =   3135
         _Version        =   589828
         _ExtentX        =   5530
         _ExtentY        =   1455
         _StockProps     =   66
         Caption         =   "Liquor PLU Dump"
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
         textCaption     =   "Mix.frx":44C2
         textLT          =   "Mix.frx":4540
         textCT          =   "Mix.frx":4558
         textRT          =   "Mix.frx":4570
         textLM          =   "Mix.frx":4588
         textRM          =   "Mix.frx":45A0
         textLB          =   "Mix.frx":45B8
         textCB          =   "Mix.frx":45D0
         textRB          =   "Mix.frx":45E8
         colorBack       =   "Mix.frx":4600
         colorIntern     =   "Mix.frx":462A
         colorMO         =   "Mix.frx":4654
         colorFocus      =   "Mix.frx":467E
         colorDisabled   =   "Mix.frx":46A8
         colorPressed    =   "Mix.frx":46D2
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
End
Attribute VB_Name = "Sales_Mix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Selection    As SaleMixInstruction
Private Display     As Integer
Public Load_View    As Integer
Public HMix         As Boolean
Private Sub Arrow_dwn_Click()
Display = Display + 1
If Display >= Sel_Server.ListCount - 1 Then Display = Sel_Server.ListCount - 1
Sel_Server.ListIndex = Display
End Sub
Private Sub arrow_up_Click()
Display = Display - 1
If Display < 0 Then Display = 0
Sel_Server.ListIndex = Display
End Sub

Private Sub BtnEnh1_Click()
Dim LT As Long
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
 If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Sub
    Sql = "Delete from [Contests] Where [Contest Name] = '" & ContestName.Text & "'"
  Db.Execute Sql
  OpenTb "Select * from [Contests] Where [Contest Name] = '" & ContestName.Text & "'", Tb, Db
   For LT = 0 To Sel_Server.ListCount - 1
      If Sel_Server.Selected(LT) = True Then
        Tb.AddNew
        Tb![Menu Item] = Sel_Server.List(LT) & ""
        Tb![Contest Name] = ContestName.Text
        Tb.fields("StartTime").Value = STime.Text & ""
        Tb.fields("EndTime").Value = ETime.Text & ""
        Tb.Update
      End If
    Next LT
  DbClose Db, Tb
ContestName.AddItem ContestName.Text
See_Msg "Saved", 5, ""
End Sub
Private Sub BtnEnh10_Click()
Dim SM As Integer
If MixList.ListCount = 0 Then Exit Sub
If MixList.Selected(0) = False Then
  For SM = 0 To MixList.ListCount - 1
    MixList.Selected(SM) = True
  Next SM
  BtnEnh10.Caption = "Uncheck All"
Else
  For SM = 0 To MixList.ListCount - 1
    MixList.Selected(SM) = False
  Next SM
  BtnEnh10.Caption = "Select All"
End If
End Sub

Private Sub BtnEnh11_Click()
Dim F As Integer
Dim Itm As Integer
F = -1
ReDim Item_Search(MixList.SelCount)
  For Itm = 0 To MixList.ListCount - 1
    If MixList.Selected(Itm) = True Then
      F = F + 1
      Item_Search(F) = MixList.List(Itm)
    End If
  Next Itm
If F = -1 Then
  See_Msg "Must Select a Department", 5, ""
  Exit Sub
End If
CurrentDepartmentMix Item_Search(), F
PrintDeptMix Format(Date, "MM/DD/YYYY"), Format(Date, "MM/DD/YYYY"), F
End Sub
Public Sub CurrentCategoryMix(SearchItems() As String, SNumber As Integer)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim Ndb As ADODB.Connection
Dim Nrs As ADODB.Recordset
Set SMix = New SalesMix
'Added to accept date entry as well as numeric
Dim qty As Single
Dim Reg As Integer
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Ndb
  Sql = "Select * from [Check Detail]"
  Sql = Sql & " Where [Close Out Day] = 0"
ConnectADORS Sql, Ndb, Nrs
If Nrs.EOF Then
   See_Msg "No Sales Found", 5, ""
   DbClose Ndb, Nrs
   Exit Sub
End If
SetNothing Nrs
        Sql = "Select Distinct [Menu Item],[Cateory Name],[Department Report], Sum([Quantity]) as QtySold, Sum([Price]) as TotalSold ,Sum([Cost Amount]) as TCost From [Check Detail] "
        Sql = Sql & " Where [Close Out Day] = 0"
        Sql = Sql & " and [Cateory Name] = '" & SearchItems(0) & "'"
        For Reg = 1 To SNumber
           Sql = Sql & " Or [Close Out Day] = 0"
           Sql = Sql & " and [Cateory Name] = '" & SearchItems(Reg) & "'"
        Next Reg
        Sql = Sql & " Group By [Menu Item],[Cateory Name],[Department Report]"
        Sql = Sql & " Order by [Cateory Name],[Menu Item] asc"
        ConnectADORS Sql, Ndb, Tb
        Update.Show
        Update.UPdateWhere.Caption = "Reading Sales "
        Update.Refresh
           Do While Not Tb.EOF
                      With SMix
                            If .Item(Tb![Menu Item] & Tb![Cateory Name]) Is Nothing Then
                               .add Tb![Menu Item] & Tb![Cateory Name]
                            End If
                            qty = Tb!QtySold
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).SalesBy = ""
                             If qty = 0 Then qty = 1
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).MenuItemName = Tb![Menu Item]
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).QtySold = .Item(Tb![Menu Item] & Tb![Cateory Name]).QtySold + qty
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).DollarsSold = (.Item(Tb![Menu Item] & Tb![Cateory Name]).DollarsSold + Format(Val(Tb!Totalsold & ""), "0.00"))
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).SoldCost = (.Item(Tb![Menu Item] & Tb![Cateory Name]).SoldCost + Format(Val(Tb!TCost & ""), "0.00"))
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).MenuDeptName = Tb![Department Report]
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).MenuCatName = Tb![Cateory Name]
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).IsDiscount = False
                        End With
                        If Tb.AbsolutePosition > 0 Then
                         Update.Progress.ProgressBarValue = (Tb.AbsolutePosition / RecordsCount(Tb)) * 100
                        End If
                      Update.Refresh
              Tb.MoveNext
            Loop
DbClose Db, Tb
Call MemoryCheck(Update)
Leave:
Exit Sub
ErrorHandler:
If Err.Number = 6 Then Resume Next
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Department Mix ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub CurrentDepartmentMix(SearchItems() As String, SNumber As Integer)
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim Ndb As ADODB.Connection
Dim Nrs As ADODB.Recordset

Set SMix = New SalesMix
'Added to accept date entry as well as numeric
Dim qty As Single
Dim Reg As Integer
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Ndb
  Sql = "Select * from [Check Detail]"
  Sql = Sql & " Where [Close Out Day] = 0"
ConnectADORS Sql, Ndb, Nrs
If Nrs.EOF Then
   See_Msg "No Sales Found", 5, ""
   DbClose Ndb, Nrs
   Exit Sub
End If
SetNothing Nrs
        Sql = "Select Distinct [Menu Item],[Cateory Name],[Department Report], Sum([Quantity]) as QtySold, Sum([Price]) as TotalSold ,Sum([Cost Amount]) as TCost, Count([Menu Item]) as TMods From [Check Detail] "
        Sql = Sql & " Where [Close Out Day] = 0 and [Voided] =0 "
        If RType.Caption = "Category Mix Report" Then
          Sql = Sql & " and [Cateory Name] = '" & SearchItems(0) & "'"
          For Reg = 1 To SNumber
           Sql = Sql & " Or [Close Out Day] = 0"
           Sql = Sql & " and [Cateory Name] = '" & SearchItems(Reg) & "'"
          Next Reg
        Else
          Sql = Sql & " and [Department Report] = '" & SearchItems(0) & "'"
          For Reg = 1 To SNumber
           Sql = Sql & " Or [Close Out Day] = 0"
           Sql = Sql & " and [Department Report] = '" & SearchItems(Reg) & "'"
          Next Reg
        End If
        Sql = Sql & " Group By [Menu Item],[Cateory Name],[Department Report]"
        Sql = Sql & " Order by [Department Report],[Menu Item] asc"
        ConnectADORS Sql, Ndb, Tb
        Update.Show
        Update.UPdateWhere.Caption = "Reading Sales "
        Update.Refresh
           Do While Not Tb.EOF
                      With SMix
                            If .Item(Tb![Menu Item] & Tb![Cateory Name]) Is Nothing Then
                               .add Tb![Menu Item] & Tb![Cateory Name]
                            End If
                            qty = Tb!QtySold
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).SalesBy = ""
                             If qty = 0 Then qty = Tb.fields("TMods").Value
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).MenuItemName = Tb![Menu Item]
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).QtySold = .Item(Tb![Menu Item] & Tb![Cateory Name]).QtySold + qty
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).DollarsSold = (.Item(Tb![Menu Item] & Tb![Cateory Name]).DollarsSold + Format(Val(Tb!Totalsold & ""), "0.00"))
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).SoldCost = (.Item(Tb![Menu Item] & Tb![Cateory Name]).SoldCost + Format(Val(Tb!TCost & ""), "0.00"))
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).MenuDeptName = Tb![Department Report]
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).MenuCatName = Tb![Cateory Name]
                             .Item(Tb![Menu Item] & Tb![Cateory Name]).IsDiscount = False
                        End With
                        If Tb.AbsolutePosition > 0 Then
                         Update.Progress.ProgressBarValue = (Tb.AbsolutePosition / RecordsCount(Tb)) * 100
                        End If
                      Update.Refresh
              Tb.MoveNext
            Loop
DbClose Db, Tb
Call MemoryCheck(Update)
Leave:
Exit Sub
ErrorHandler:
If Err.Number = 6 Then Resume Next
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Department Mix ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub BtnEnh12_Click()
End Sub
Private Sub BtnEnh13_Click()
PrintDailyInventory False
End Sub

Private Sub BtnEnh14_Click()
Dim SvrName() As String
ReDim SvrName(0)
BtnEnh2_Click
SSCommand4_Click
Dim F As Integer, Itm As Integer
Dim NS As Currency
Sql = "Select Distinct [Date] from [Sales]"
LoadDates Sql, SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then Exit Sub
   Sleep 0
   If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If CDate(SearchDates.Item(1).StartDate) = CDate(Date) Then
      ReDim Item_Search(1)
      NS = ServerSalesContest(True)
      InventoryUsage "Inventory Usage For " & vbCrLf & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate, NS
    Else
      HistoryItemCount SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
      InventoryUsage "Inventory Usage For " & vbCrLf & SearchDates.Item(1).StartDate & " to " & SearchDates.Item(1).EndDate, GetNetSales(SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate)
    End If
   Else
     See_Msg "Invalid Date", 5, ""
     Exit Sub
   End If

End Sub
Private Sub BtnEnh15_Click()
On Error GoTo ErrorHandler
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim AL      As Integer
If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Sub
  Sql = "Delete from [Berg Data] Where [Menu Item] = '" & MenuItem.Text & "'"
  Db.Execute Sql
  Sql = "Select * from [Berg Data]"
   ConnectADORS Sql, Db, RS
     With Berg
       For AL = 1 To .rows - 1
        If .TextMatrix(AL, 0) = "" Then Exit For
          RS.AddNew
          RS![Menu Item] = Trim$(MenuItem.Text)
          RS![Plu No] = .TextMatrix(AL, 0) & ""
          RS![Quantity] = Val(.TextMatrix(AL, 1) & "")
          RS![Printer] = "Inventory"
          RS.Update
       Next AL
     End With
DbClose Db, RS
See_Msg "Saved....", 5, ""
Berg.Clear
MenuItem.Text = ""
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Save Usage")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub BtnEnh16_Click()
On Error Resume Next
Dim Db      As ADODB.Connection
Dim RS As ADODB.Recordset
'OpenDb SvrPath.TerminalPath & "CrMenu.mdb", Db
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
  Sql = "Delete from [Berg Data] Where [Menu Item] = '" & MenuItem.Text & "' and [Printer] = 'Inventory'"
Db.Execute Sql
MenuItem.Text = ""
DbClose Db, RS
See_Msg "Removed....", 5, ""
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "delete usage ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh17_Click()
bergbox.Visible = False
End Sub

Private Sub BtnEnh18_Click()
LoadIngredientList
CenterObject Me, bergbox
bergbox.Visible = True
bergbox.ZOrder 0
End Sub

Private Sub BtnEnh19_Click()
End Sub

Private Sub BtnEnh2_Click()
Load_View = 0
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Records As ADODB.Recordset
Dim LT As Long
Dim AllSort As SortAll
Set AllSort = New SortAll
MenuItem.Clear
OpenDb "MSaleData", Db
OpenTb "Select * from [Shift Keys] Order by [Category Name] asc", Tb, Db
 Do While Not Tb.EOF
    If AllSort.Item(Tb![Category Name] & "" & Tb![Key Name] & "") Is Nothing Then
      AllSort.add Tb![Category Name] & "" & Tb![Key Name] & ""
      AllSort.Item(Tb![Category Name] & "" & Tb![Key Name] & "").MenuItem = Tb![Category Name] & ""
      AllSort.Item(Tb![Category Name] & "" & Tb![Key Name] & "").ModName = Tb![Key Name] & ""
    End If
   Tb.MoveNext
 Loop
SetNothing Tb
OpenTb "Select * from [MenuItemCountDown] Where [Quick Pick] = 0 Order By [Menu Item Name] asc", Tb, Db
Sel_Server.Clear
Do While Not Tb.EOF
    If IsThere(Sel_Server, Tb(0)) = False Then
     Sel_Server.AddItem Tb(0) & ""
     MenuItem.AddItem Tb(0) & ""
     List1.AddItem Tb![Department Name] & "", Sel_Server.NewIndex
      For LT = 1 To AllSort.count
         If AllSort.Item(LT).MenuItem = Tb![Department Name] & "" Then
           Sel_Server.AddItem AllSort.Item(LT).ModName & " " & Tb(0) & ""
           MenuItem.AddItem AllSort.Item(LT).ModName & " " & Tb(0) & ""
         End If
      Next LT
    End If
    Tb.MoveNext
Loop
SetNothing Tb
OpenTb "Select * from [Category Information]", Tb, Db
Cats.Clear
Cats.Text = "None"
Do While Not Tb.EOF
    Cats.AddItem Tb![Category Name] & ""
    Tb.MoveNext
Loop
SetNothing Tb
OpenTb "Select * from [Sub Categories]", Tb, Db
Do While Not Tb.EOF
    Cats.AddItem Tb![Sub Name] & ""
    Tb.MoveNext
Loop
DbClose Db, Tb
Sql = "Select Distinct [Modifier Category], [Modifier Cost]"
Sql = Sql & "From [Modifiers] "
Sql = Sql & "Group By [Modifier category], [Modifier Cost]"
Sql = Sql & "Order By [Modifier Category] asc"
'OpenDb SvrPath.LocalTerminalPath & "MenuMods.mdb", DB
ConnectADODB "MSaleData", Db
ConnectADORS Sql, Db, Records
  Do While Not Records.EOF
   If InStr(Records![Modifier Category], "!") > 0 Then
     If IsThere(Sel_Server, Mid(Records![Modifier Category], 2, Len(Records![Modifier Category]))) = False Then
      Sel_Server.AddItem Mid(Records![Modifier Category], 2, Len(Records![Modifier Category]))
      MenuItem.AddItem Mid(Records![Modifier Category], 2, Len(Records![Modifier Category]))
    End If
   Else
     If IsThere(Sel_Server, Records![Modifier Category] & "") = False Then
      Sel_Server.AddItem Records![Modifier Category] & ""
      MenuItem.AddItem Records![Modifier Category] & ""
     End If
   End If
   Records.MoveNext
  Loop
DbClose Db, Records

If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Sub
 OpenTb "Select Distinct [Contest Name] from [Contests]", Tb, Db
 ContestName.Clear
 ContestName.Text = "None"
 ContestName.AddItem "Server Sales Mix"
 ContestName.AddItem "Server Contest Mix"
 ContestName.AddItem "Daily Mix"
 Do While Not Tb.EOF
    ContestName.AddItem Tb![Contest Name] & ""
    Tb.MoveNext
 Loop
 DbClose Db, Tb

Display = 8
Set AllSort = Nothing
End Sub

Private Sub BtnEnh20_Click()
On Error Resume Next
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
OpenDb "MSaleData", Db
OpenTb "Select * from [Report Categories] Order by [Report Name] asc", Tb, Db
MixList.Clear
Do While Not Tb.EOF
    MixList.AddItem Check_Null(Tb(0))
    Tb.MoveNext
Loop
DbClose Db, Tb
With MixByDept
CenterObject Me, MixByDept
.Visible = True
.ZOrder 0
End With
RType.Caption = "Department Mix Report"
BtnEnh10.Visible = True
DoEvents
LoadItemCost
End Sub

Private Sub BtnEnh21_Click()
LoadBreadCount
End Sub

Private Sub BtnEnh22_Click()
LoadDates ("Select Distinct TRY_CAST([Date] AS datetime) from [Sales]"), SvrPath.TerminalPath & "Closing History.MDB", "Date"
If SearchDates Is Nothing Then

Else
  Call ExportProductMix(SvrPath.TerminalPath, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate)
  Set SearchDates = Nothing
End If


End Sub

Private Sub BtnEnh3_Click()
MixByDept.Visible = False
MixByDept.Left = -5000
End Sub
Private Sub BtnEnh4_Click()
Dim F As Integer
Dim Itm As Integer
F = -1
ReDim Item_Search(MixList.SelCount)
  For Itm = 0 To MixList.ListCount - 1
    If MixList.Selected(Itm) = True Then
      F = F + 1
      Item_Search(F) = MixList.List(Itm)
    End If
  Next Itm
If F = -1 Then
  See_Msg "Must Select a Department", 5, ""
  Exit Sub
End If
Sql = "Select Distinct [Date] from [Sales]"
LoadDates Sql, SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then Exit Sub
   If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
     If RType.Caption = "Department Mix Report" Then
     DepartmentMix SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, Item_Search(), F, Department
     Else
     DepartmentMix SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, Item_Search(), F, Category
     End If
     PrintDeptMix SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, F
   Else
     See_Msg "Invalid Date", 5, ""
     Exit Sub
   End If
End Sub
Private Sub BtnEnh5_Click()
End Sub
Private Sub LoadBreadCount()
On Error Resume Next
MixList.Clear
Dim L  As Integer
Dim TXTFileInfo() As Variant
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "DAT Files\" & "BreadCount.Dat"), TXTFileInfo()) Then
    For L = 0 To UBound(TXTFileInfo)
       MixList.AddItem CStr(TXTFileInfo(L))
       MixList.Selected(L) = True
    Next L
Else
  See_Msg "See Category Setup", 5, "Nothing Selected"
  Exit Sub
End If
With MixByDept
CenterObject Me, MixByDept
.Visible = True
.ZOrder 0
End With
BtnEnh10.Visible = False
RType.Caption = "Category Mix Report"
DoEvents
LoadItemCost
End Sub

Private Sub BtnEnh7_Click()
If ContestName.Text = "" Or ContestName.Text = "None" Then Exit Sub
Dim LT As Long
Dim Db As ADODB.Connection
'OpenDb SvrPath.TerminalPath & "CrMenu.mdb", Db
If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Sub
 Sql = "Delete from [Contests] Where [Contest Name] = '" & ContestName.Text & "'"
Db.Execute Sql
SetNothing Db
RemoveItem ContestName, ContestName.Text
See_Msg "Removed...", 5, ""
End Sub
Private Sub BtnEnh8_Click()
On Error GoTo ErrorHandler
'If See_Msg("Allow Duplicate Items", 4, "Duplicates By Category") = 6 Then
  GroupBy.Value = 0
'Else
' GroupBy.Value = 1
'End If
Load_View = 0
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim LT As Long
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
If GroupBy.Value = -1 Then
 ConnectADORS "Select Distinct [Menu Item] from [Detail] Group by [Menu Item] Order by [Menu Item] asc ", Db, Tb
Else
 ConnectADORS "Select Distinct [Menu Item],[Category Name] from [Detail] Group by [Menu Item],[Category Name] Order by [Menu Item] asc ", Db, Tb
End If
Sel_Server.Clear
Do While Not Tb.EOF
    Sel_Server.AddItem Tb(0) & ""
    MenuItem.AddItem Tb(0) & ""
  'If GroupBy.Value = 0 Then
  '  List1.AddItem Tb![Category Name] & "", Sel_Server.NewIndex
  'End If
  Tb.MoveNext
Loop
Cats.Clear
ConnectADORS "Select Distinct [Category Name] from [Detail] Group by [Category Name] order by [Category Name] asc ", Db, Tb
Do While Not Tb.EOF
    If Tb(0) & "" > "" Then
    Cats.AddItem Tb(0) & ""
    End If
    Tb.MoveNext
Loop
DbClose Db, Tb
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Group Revenue ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub BtnEnh9_Click()
End Sub

Private Sub Command1_Click()

End Sub

Private Sub ContestName_Click()
'BtnEnh2_Click
Dim LT As Long
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
'OpenDb SvrPath.TerminalPath & "CrMenu.mdb", Db
If ConnectADODB(SvrPath.LocalTerminalPath & "MSaleData", Db) = False Then Exit Sub
STime.Text = "None"
ETime.Text = "None"
OpenTb "Select * from [Contests] Where [Contest Name] = '" & ContestName.Text & "' Order by [Menu Item] asc", Tb, Db
If Not Tb.EOF Then
  STime.Text = Tb.fields("Starttime").Value & ""
  ETime.Text = Tb.fields("Endtime").Value & ""
 Do While Not Tb.EOF
    For LT = 0 To Sel_Server.ListCount - 1
     If Check_Null(Tb(0)) = Sel_Server.List(LT) And Sel_Server.Selected(LT) = False Then
      Sel_Server.Selected(LT) = True
      Exit For
     End If
    Next LT
    Tb.MoveNext
 Loop
End If
DbClose Db, Tb
End Sub
Private Sub ETime_Click()
Dim rTime As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
rTime = PKT.ReturnTime
If rTime > "" Then
  ETime.Text = rTime
Else
  ETime.Text = "None"
End If
Set PKT = Nothing
End Sub
Private Sub ETime_DblClick()
ETime.Text = "None"
End Sub
Private Sub Exit_Click()
Set GetMenuCost = Nothing
Call MemoryCheck(Sales_Mix)
End Sub
Private Sub RemoveFromList(ItemName As String, LBox As ComboBox)
Dim RM As Integer
For RM = 0 To LBox.ListCount - 1
If Trim(ItemName) = Trim(LBox.List(RM)) Then
LBox.RemoveItem (RM)
Exit For
End If
Next RM
End Sub
Private Sub Form_Load()
ResizeScreen1024 Me
On Error Resume Next
CenterObject Me, Box
With Berg
    .Clear
    .cols = 3
    .rows = 20
    .ColWidth(0) = 2000
    .ColWidth(1) = 1000
    .ColWidth(2) = 0
    .TextMatrix(0, 0) = "Name"
    .TextMatrix(0, 1) = "Quantity"
    .TextMatrix(0, 2) = "PLU"
    .FixedCols = 0
    .FixedRows = 1
End With
End Sub
Private Sub LoadIngredientList()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
'OpenDb SvrPath.TerminalPath & "Crmenu.mdb", Db
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
    Sql = "Select * from [Ingredient List] Order by [Ingredient Item] asc"
    ConnectADORS Sql, Db, RS
    ItemList.Clear
    Berg.ColComboList(0) = ""
    Do While Not RS.EOF
     Berg.ColComboList(0) = Berg.ColComboList(0) & RS![Ingredient Item] & "" & "|"
     ItemList.AddItem RS![Ingredient Item] & ""
     RS.MoveNext
    Loop
    DbClose Db, RS
End Sub
Private Sub MixByMealPeriod(SDate As String, EDate As String)
On Error GoTo ErrorHandler
LoadMealPeriods
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim MP As Integer
Dim DP As Integer
Dim CP As Integer
Dim PD As Integer
Dim MaxDept As Integer
Dim MaxCat  As Integer
Dim MaxDepartments As Integer
Dim DeptSales() As BusinessTimes
Dim CategorySales() As BusinessTimes
Dim DepartmentSales() As BusinessTimes
Dim POnly   As Boolean
Dim MHour As Long
Me.Refresh
BtnEnh9.Caption = "Loading Periods"
Me.Refresh
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
  Sql = "Select [Department Name] From [Detail] where [Department name] > '' and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(EDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Group by [Department Name] Order by [Department Name] asc"
  ConnectADORS ReturnDateString(Sql), Db, Tb
    If Not Tb.EOF Then
      MaxDepartments = RecordsCount(Tb)
    Else
      MaxDepartments = 0
    End If
    ReDim DepartmentSales(MaxDepartments)
    CP = -1
    Do While Not Tb.EOF
       CP = CP + 1
        DepartmentSales(CP).PeriodDepart = Tb![Department Name]
        DepartmentSales(CP).PeriodSales = 0
        DepartmentSales(CP).Breakfast = 0
        DepartmentSales(CP).Lunch = 0
        DepartmentSales(CP).Thru = 0
        DepartmentSales(CP).Dinner = 0
        DepartmentSales(CP).Late = 0
       Tb.MoveNext
    Loop
  SetNothing Tb
  Sql = "Select  [Department Name], [Category Name] From [Detail] where [Category name] > '' and TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(EDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
  Sql = Sql & " Group by [Department Name],[Category Name] Order by [Department Name], [Category Name] asc"
  ConnectADORS ReturnDateString(Sql), Db, Tb
    If Not Tb.EOF Then
      MaxCat = RecordsCount(Tb)
    Else
      MaxCat = 0
    End If
    ReDim CategorySales(MaxCat)
    CP = -1
    Do While Not Tb.EOF
       CP = CP + 1
        CategorySales(CP).PeriodCategory = Tb![Category Name]
        CategorySales(CP).PeriodDepart = Tb![Department Name]
        CategorySales(CP).PeriodSales = 0
        CategorySales(CP).Breakfast = 0
        CategorySales(CP).Lunch = 0
        CategorySales(CP).Thru = 0
        CategorySales(CP).Dinner = 0
        CategorySales(CP).Late = 0
       Tb.MoveNext
    Loop
    SetNothing Tb
POnly = False
Sql = "Select Distinct [Department Name], [Category Name], [Menu Item] From [Detail] where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(EDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
If See_Msg("Filter Mix", 4, "Items with Prices Only ?") = 6 Then
  Sql = Sql & " and [Item Price] > 0"
  POnly = True
End If
Sql = Sql & " Group by [Department Name],[Category Name],[Menu Item] Order by [Department Name], [Category Name],[Menu Item] asc"
ConnectADORS ReturnDateString(Sql), Db, Tb
  If Not Tb.EOF Then
    MaxDept = RecordsCount(Tb)
  Else
    MaxDept = 0
  End If
ReDim DeptSales(MaxDept)
DP = -1
Do While Not Tb.EOF
   DP = DP + 1
   DeptSales(DP).PeriodCategory = Tb![Category Name] & ""
   DeptSales(DP).PeriodDepart = Tb![Department Name] & ""
   DeptSales(DP).PeriodName = Tb![Menu Item] & ""
   DeptSales(DP).PeriodSales = 0
   DeptSales(DP).Breakfast = 0
   DeptSales(DP).Lunch = 0
   DeptSales(DP).Thru = 0
   DeptSales(DP).Dinner = 0
   DeptSales(DP).Late = 0
   Tb.MoveNext
Loop
SetNothing Tb
Sql = "Select * From [Detail] where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(SDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(EDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
If POnly Then
  Sql = Sql & " and [Item Price] > 0"
End If
ConnectADORS ReturnDateString(Sql), Db, Tb
For MP = 0 To 4
If Not Tb.BOF Then Tb.MoveFirst
If BusinessPeriods(MP).PeriodName <> "None" Then
 If IsDate(BusinessPeriods(MP).PeriodStart) And IsDate(BusinessPeriods(MP).PeriodEnd) Then
   Do While Not Tb.EOF
    'BtnEnh9.Caption = "Searching " & Tb![Time] & " " & Tb![Menu Item] & ""
    'Sleep 0
    If IsDate(Tb![Time]) Then
         Select Case Hour(Format(Tb![Time], "HH:MM ampm"))
             Case Is = 0, 1, 2, 3
                 MHour = Hour(Format(Tb![Time], "HH:MM ampm")) + 24
             Case Else
                 MHour = Hour(Format(Tb![Time], "HH:MM ampm"))
         End Select
         If MHour >= Hour(BusinessPeriods(MP).PeriodStart) And MHour <= BusinessPeriods(MP).PeriodEndHour Then
          For DP = 0 To MaxDept
           If DeptSales(DP).PeriodName = Tb![Menu Item] & "" And DeptSales(DP).PeriodDepart = Tb![Department Name] & "" Then
             DeptSales(DP).PeriodSales = DeptSales(DP).PeriodSales + Format(Val(Tb![Item Price] & ""), "0.00")
             DeptSales(DP).PeriodCount = DeptSales(DP).PeriodCount + Format(Val(Tb![Quantity] & ""), "0.00")
             Select Case MP
                 Case Is = 0
                   DeptSales(DP).Breakfast = DeptSales(DP).Breakfast + Format(Val(Tb![Quantity] & ""), "0.00")
                 Case Is = 1
                   DeptSales(DP).Lunch = DeptSales(DP).Lunch + Format(Val(Tb![Quantity] & ""), "0.00")
                 Case Is = 2
                   DeptSales(DP).Thru = DeptSales(DP).Thru + Format(Val(Tb![Quantity] & ""), "0.00")
                 Case Is = 3
                   DeptSales(DP).Dinner = DeptSales(DP).Dinner + Format(Val(Tb![Quantity] & ""), "0.00")
                 Case Is = 4
                   DeptSales(DP).Late = DeptSales(DP).Late + Format(Val(Tb![Quantity] & ""), "0.00")
             End Select
           End If
          Next DP
         End If
    End If
    Tb.MoveNext
   Loop
  End If
 End If
Next MP
DbClose Db, Tb
SetPage ReportView.View, "Meal Period Mix Report"
PrintLine "Date Start : " & SDate, -1
PrintLine "Date End   : " & EDate, -1
DrawLine
PrintLine "Item", 3000: PrintLine BusinessPeriods(0).PeriodName, 4000: PrintLine BusinessPeriods(1).PeriodName, 5000: PrintLine BusinessPeriods(2).PeriodName, 6000: PrintLine BusinessPeriods(3).PeriodName, 7000: PrintLine BusinessPeriods(4).PeriodName, 8000: PrintLine "Total", 9000: PrintLine "Sales", -1
DrawLine
For PD = 0 To MaxDepartments
  For DP = 0 To MaxDept
   If DepartmentSales(PD).PeriodDepart = DeptSales(DP).PeriodDepart Then
     DepartmentSales(PD).Breakfast = DepartmentSales(PD).Breakfast + DeptSales(DP).Breakfast
     DepartmentSales(PD).Lunch = DepartmentSales(PD).Lunch + DeptSales(DP).Lunch
     DepartmentSales(PD).Thru = DepartmentSales(PD).Thru + DeptSales(DP).Thru
     DepartmentSales(PD).Dinner = DepartmentSales(PD).Dinner + DeptSales(DP).Dinner
     DepartmentSales(PD).Late = DepartmentSales(PD).Late + DeptSales(DP).Late
     DepartmentSales(PD).PeriodCount = DepartmentSales(PD).PeriodCount + DeptSales(DP).PeriodCount
     DepartmentSales(PD).PeriodSales = DepartmentSales(PD).PeriodSales + DeptSales(DP).PeriodSales
   End If
  Next DP
Next PD
 For CP = 0 To MaxCat
  For DP = 0 To MaxDept
   If CategorySales(CP).PeriodCategory = DeptSales(DP).PeriodCategory And CategorySales(CP).PeriodDepart = DeptSales(DP).PeriodDepart Then
     CategorySales(CP).Breakfast = CategorySales(CP).Breakfast + DeptSales(DP).Breakfast
     CategorySales(CP).Lunch = CategorySales(CP).Lunch + DeptSales(DP).Lunch
     CategorySales(CP).Thru = CategorySales(CP).Thru + DeptSales(DP).Thru
     CategorySales(CP).Dinner = CategorySales(CP).Dinner + DeptSales(DP).Dinner
     CategorySales(CP).Late = CategorySales(CP).Late + DeptSales(DP).Late
     CategorySales(CP).PeriodCount = CategorySales(CP).PeriodCount + DeptSales(DP).PeriodCount
     CategorySales(CP).PeriodSales = CategorySales(CP).PeriodSales + DeptSales(DP).PeriodSales
   End If
  Next DP
Next CP
For PD = 0 To MaxDepartments
  If DepartmentSales(PD).PeriodCategory <> "*" And DepartmentSales(PD).PeriodCount > 0 Then
    DrawLine
    PrintLine "--> " & DepartmentSales(PD).PeriodDepart, 3000
    PrintLine Format(DepartmentSales(PD).Breakfast, "0.00"), 4000
    PrintLine Format(DepartmentSales(PD).Lunch, "0.00"), 5000
    PrintLine Format(DepartmentSales(PD).Thru, "0.00"), 6000
    PrintLine Format(DepartmentSales(PD).Dinner, "0.00"), 7000
    PrintLine Format(DepartmentSales(PD).Late, "0.00"), 8000
    PrintLine Format(DepartmentSales(PD).PeriodCount, "0.00"), 9000
    PrintLine Format(DepartmentSales(PD).PeriodSales, "$0.00"), -1
    DrawLine
  For CP = 0 To MaxCat
   If DepartmentSales(PD).PeriodDepart = CategorySales(CP).PeriodDepart And CategorySales(CP).PeriodCount > 0 Then
    'If CategorySales(CP).PeriodCount > 0 Then
     DrawLine
     PrintLine "      *** " & CategorySales(CP).PeriodCategory, 3000
     PrintLine Format(CategorySales(CP).Breakfast, "0.00"), 4000
     PrintLine Format(CategorySales(CP).Lunch, "0.00"), 5000
     PrintLine Format(CategorySales(CP).Thru, "0.00"), 6000
     PrintLine Format(CategorySales(CP).Dinner, "0.00"), 7000
     PrintLine Format(CategorySales(CP).Late, "0.00"), 8000
     PrintLine Format(CategorySales(CP).PeriodCount, "0.00"), 9000
     PrintLine Format(CategorySales(CP).PeriodSales, "$0.00"), -1
     DrawLine
     For DP = 0 To MaxDept
       If CategorySales(CP).PeriodDepart = DeptSales(DP).PeriodDepart And CategorySales(CP).PeriodCategory = DeptSales(DP).PeriodCategory And DeptSales(DP).PeriodCount > 0 Then
           PrintLine DeptSales(DP).PeriodName, 3000
           PrintLine Format(DeptSales(DP).Breakfast, "0.00"), 4000
           PrintLine Format(DeptSales(DP).Lunch, "0.00"), 5000
           PrintLine Format(DeptSales(DP).Thru, "0.00"), 6000
           PrintLine Format(DeptSales(DP).Dinner, "0.00"), 7000
           PrintLine Format(DeptSales(DP).Late, "0.00"), 8000
           PrintLine Format(DeptSales(DP).PeriodCount, "0.00"), 9000
           PrintLine Format(DeptSales(DP).PeriodSales, "$0.00"), -1
      End If
    Next DP
   End If
  Next CP
 End If
Next PD
DrawLine
PrintLine "", -1
FinishJob 1
BtnEnh9.Caption = "Sales Mix By Meal Period"
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Mix Bu ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub PrintDeptMix(MixPeriodS As String, MixPeriodE As String, Ind_Selection As Integer)
Dim CTR As Integer
Dim Avt As Integer
CTR = 0
Dim W As Integer
Dim DeptName As String
Dim TotalUnits As Single
Dim DptSale    As Single
Dim DptCost    As Single
Dim Header     As String
    Header = RType.Caption & vbCrLf
    Header = Header & MixPeriodS & " To " & MixPeriodE
    SetPage ReportView.View, Header
    SetBold
    Dim F$
    F = "+<2500|+1000|+>1000|+>1000|+>1000;"
    Dim H$
    'H = ReportType & "|Quantity|Sales|Cost|%;"
    SetRegular 0
    Dim Data As String
    TotalUnits = 0: CatperTotal = 0: TotalDoll = 0
    TCost = 0: CategoryFoodCost = 0
    DeptName = "": Data = ""
    For W = 0 To Ind_Selection
      DeptName = Item_Search(W)
      DptCost = 0: DptSale = 0
      H = DeptName & "|Quantity|Sales|Cost|%;"
     For i = 1 To SMix.count
      If Trim$(UCase(Item_Search(W))) = Trim$(UCase(SMix.Item(i).MenuDeptName)) Or Trim$(UCase(Item_Search(W))) = Trim$(UCase(SMix.Item(i).MenuCatName)) Then
       
       PerofSale = 0
        If Check_Option(125) = 1 And SMix.Item(i).DollarsSold = 0 Then
          'added only prints items with a sale total
        Else
            Data = Data & SMix.Item(i).MenuItemName & "|"
            Data = Data & Format(SMix.Item(i).QtySold, "##") & "|"
            Data = Data & Format(SMix.Item(i).DollarsSold, "0.00") & "|"
            Data = Data & Format(SMix.Item(i).SoldCost, "0.00") & "|"
            If SMix.Item(i).DollarsSold > 0 And SMix.Item(i).SoldCost > 0 Then
             Data = Data & Format((SMix.Item(i).SoldCost / SMix.Item(i).DollarsSold) * 100, "0.00") & ";"
            Else
             Data = Data & "0" & ";"
            End If
            TTrans = TTrans + Val(SMix.Item(i).QtySold)
            CatperTotal = CatperTotal + PerofSale
            TotalDoll = TotalDoll + SMix.Item(i).DollarsSold
            TotalUnits = TotalUnits + SMix.Item(i).QtySold
            TCost = TCost + SMix.Item(i).SoldCost
            DptCost = DptCost + SMix.Item(i).SoldCost
            DptSale = DptSale + SMix.Item(i).DollarsSold
            FoodCost = FoodCost + TCost
            CategoryFoodCost = CategoryFoodCost + SMix.Item(i).SoldCost
        End If
      End If
     Next i
        Data = Data & "Total" & "|"
        Data = Data & Format(TTrans, "0") & "|"
        Data = Data & Format(DptSale, "0.00") & "|"
        Data = Data & Format(DptCost, "0.00") & "|"
        If DptCost > 0 And DptSale > 0 Then
           Data = Data & Format((DptCost / DptSale) * 100, "0.00") & ";"
        Else
           Data = Data & "0" & ";"
        End If
        ReportView.View.TableBorder = tbTopBottom
        ReportView.View.addTable F, H, Data, vbBlack, vbWhite
        Data = ""
    Next W
    If TotalUnits > 0 Then
        Data = Data & "Total" & "|"
        Data = Data & Format(TotalUnits, "0") & "|"
        Data = Data & Format(TotalDoll, "0.00") & "|"
        Data = Data & Format(CategoryFoodCost, "0.00") & "|"
        If DptCost > 0 And DptSale > 0 Then
          Data = Data & Format((CategoryFoodCost / TotalDoll) * 100, "0.00") & ";"
        Else
          Data = Data & "0" & ";"
        End If
    End If
    ReportView.View.TableBorder = tbTopBottom
    ReportView.View.addTable F, "", Data, vbBlack, vbWhite
    FinishJob 1
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Print Dept Mix ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub MSFlexGrid1_Click()
End Sub
Private Sub Get_Day_Click()
Load_View = 1
Dim F As Integer, Itm As Integer
F = 0
ReDim Item_Search(Sel_Server.SelCount)
  For Itm = 0 To Sel_Server.ListCount - 1
    If Sel_Server.Selected(Itm) = True Then
      Item_Search(F) = Sel_Server.List(Itm)
      F = F + 1
    End If
  Next Itm
Ind_Selection = Sel_Server.SelCount
Sql = "Select Distinct [Date] from [Sales]"
LoadDates Sql, SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then Exit Sub
Dim SvrName() As String
ReDim SvrName(0)
   If IsDate(SearchDates.Item(1).StartDate) And IsDate(SearchDates.Item(1).EndDate) Then
    If UCase(ContestName.Text) = "GAMES" Then
      GameCountFromHistory SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, STime.Text, ETime.Text, IndSales.Value, SvrName(), GroupBy.Value
      PrintGameMix SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, STime.Text, ETime.Text, "Game Report", SvrName(), False
    Else
      IndMixFromHistory SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, STime.Text, ETime.Text, IndSales.Value, SvrName(), GroupBy.Value
      If GroupBy.Value = 0 Then
       PrintCustomMix SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, STime.Text, ETime.Text, "Date Range Item Mix By Category", SvrName(), False
      Else
       PrintCustomMix SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, STime.Text, ETime.Text, "Date Range Item Mix", SvrName(), False
      End If
    End If
   Else
     See_Msg "Invalid Date", 5, ""
     Exit Sub
   End If

End Sub

Private Sub GrpName_Change()

End Sub

Private Sub ItemList_Click()
On Error Resume Next
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
'OpenDb SvrPath.TerminalPath & "Crmenu.mdb", Db
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
 Sql = "Select * from [Ingredient List] Where [Ingredient Item] = '" & ItemList.Text & "'"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
 ICost.Text = Format(Val(RS![Ingredient Cost] & ""), "0.00")
End If
DbClose Db, RS
End Sub
Private Sub Label3_Click()
STime.Text = "None"
End Sub

Private Sub Label6_Click()
ETime.Text = "None"
End Sub

Private Sub LiquorPluList_Click()
On Error GoTo ErrorHandler
Dim PrinterPort As String
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim Heading As String
If Check_Option(96) = 0 Then
Printer_Drivers "Receipt"
CenterText = EC(8)
LeftText = EC(9)
Dim PF As Integer
PF = FreeFile
PrinterPort = "Report"
OneMoreTime:
Open PrinterPort For Output Shared As #PF
End If
Heading = "Liquor PLU Dump"
'Plist = PList &
If Check_Option(96) = 0 Then
Print #PF, EC(1); EC(4); CenterText; Heading; EC(2); EC(7)
Print #PF, EC(4); LeftText; EC(7)
Print #PF, ; EC(7)
Print #PF, Format(Time, "HH:MM AmPm");
Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
Print #PF, String(40, 45); EC(7)
Else
SetPage ReportView.View, Heading
End If
'OpenDb SvrPath.LocalTerminalPath & "Crmenu.mdb", Db
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
 Sql = "Select * from [Berg Data] order by [Menu Item] asc"
ConnectADORS Sql, Db, Tb
'Home Phone
If Check_Option(96) = 0 Then
Print #PF, "Menu Item            PLU       Qty"; EC(7)
Print #PF, String(40, 45); EC(7)
Else

PrintLine "Menu Name", 4000
PrintLine "PLU", 6000
PrintLine "Quantity", -1
DrawLine
End If
Do While Not Tb.EOF
If Check_Option(96) = 0 Then
    Print #PF, Trim(Tb(1) & ""); Tab(21); Tb(2) & ""; Tab(29); Tb(3) & ""; EC(7)
Else
    PrintLine Tb![Menu Item], 4000
    PrintLine Tb![Plu No] & "", 6000
    PrintLine Tb![Quantity] & "", -1
End If
Tb.MoveNext
Loop
DbClose Db, Tb
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(5)
    Close #PF
SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS17")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub List2_Click()

End Sub

Private Sub List1_Click()

End Sub

Private Sub MenuItem_Click()
With Berg
    .Clear
    .cols = 3
    .rows = 20
    .ColWidth(0) = 2000
    .ColWidth(1) = 1000
    .ColWidth(2) = 0
    .TextMatrix(0, 0) = "Name"
    .TextMatrix(0, 1) = "Quantity"
    .TextMatrix(0, 2) = "PLU"
    .FixedCols = 0
    .FixedRows = 1
End With
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim lf As Integer
'OpenDb SvrPath.TerminalPath & "Crmenu.mdb", Db
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
  Sql = "Select * from [Berg Data] Where [Menu Item] = '" & MenuItem.Text & "'"
ConnectADORS Sql, Db, RS
lf = 0
With Berg
   Do While Not RS.EOF
       lf = lf + 1
       .TextMatrix(lf, 0) = RS![Plu No] & ""
       .TextMatrix(lf, 1) = Val(RS![Quantity] & "")
       .TextMatrix(lf, 2) = "" 'RS![Export PLU] & ""
     RS.MoveNext
   Loop
End With
DbClose Db, RS
End Sub

Private Sub OrderDestinationReport_Click()
End Sub

Private Sub PrintCountDownItems_Click()
On Error GoTo ErrorHandler
Dim PrinterPort     As String
Dim Heading         As String
Dim Db              As ADODB.Connection
Dim Tb              As ADODB.Recordset
Dim InventoryValue  As Currency
If Check_Option(96) = 0 Then
Printer_Drivers "Receipt"
CenterText = EC(8)
LeftText = EC(9)

Dim PF As Integer
PF = FreeFile
PrinterPort = "Report"
OneMoreTime:
  Open PrinterPort For Output Shared As #PF
End If
ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
  Sql = "Select * from [MenuItemCountDown] where [Quick Pick] = 0 and [Inventory Amount] <> " & 9999 & " order by [Menu Item Name] asc"
ConnectADORS Sql, Db, Tb
If Tb.EOF Then
  See_Msg "No Inventory Set", 5, ""
  Close
  DbClose Db, Tb
  Exit Sub
End If
Heading = "Menu Count Down"
'Plist = PList &
If Check_Option(96) = 0 Then
Print #PF, EC(1); EC(4); CenterText; Heading; EC(3); EC(7)
Print #PF, EC(4); LeftText; EC(7)
Print #PF, ; EC(7)
Print #PF, Format(Time, "HH:MM AmPm");
Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
Print #PF, String(40, 45); EC(7)
Else
SetPage ReportView.View, Heading
End If
'Home Phone
If Check_Option(96) = 0 Then
Print #PF, "Item Name                Amount On Hand"; EC(7)
Print #PF, String(40, 45); EC(7)
Else
PrintLine "Item Name", 3500
PrintLine "On Hand", 5500
PrintLine "Count", 7500
PrintLine "Cost", 9500
PrintLine "Value", -1
DrawLine
End If
InventoryValue = 0
Do While Not Tb.EOF
If Check_Option(96) = 0 Then
If Tb![Inventory Amount] <> 9999 Then Print #PF, Trim(Tb![Menu Item Name]); Tab(30); Format(Tb![Inventory Amount], "####0")
Else
 If Tb![Inventory Amount] <> 9999 Then
  PrintLine Trim(Tb![Menu Item Name]), 3700
  PrintLine Format(Tb![Inventory Amount], "####0"), 5300
  PrintLine "(________)", 7500
  PrintLine Format(Tb![Product Cost], "$0.00"), 9500
  PrintLine Format((Tb![Inventory Amount] * Tb![Product Cost]), "$0.00"), -1
  InventoryValue = InventoryValue + (Tb![Inventory Amount] * Tb![Product Cost])
 End If
End If
Tb.MoveNext
Loop
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    PrintLine "Total Inventory", 9500
    PrintLine Format(InventoryValue, "$0.00"), -1
    ReportView.View.EndDoc
    ReportView.Show 1
End If
DbClose Db, Tb
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS3")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub ShowModifierList_Click()
On Error GoTo ErrorHandler
Dim PrinterPort As String
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Heading As String
If Check_Option(96) = 0 Then
Printer_Drivers "Receipt"
CenterText = EC(8)
LeftText = EC(9)

Dim PF As Integer
PF = FreeFile
PrinterPort = "Report"
OneMoreTime:
Open PrinterPort For Output Shared As #PF
End If
Heading = "Modifier List"
'Plist = PList &
If Check_Option(96) = 0 Then
Print #PF, EC(1); EC(4); CenterText; Heading; EC(2); EC(7)
Print #PF, EC(4); LeftText; EC(7)
Print #PF, ; EC(7)
Print #PF, Format(Time, "HH:MM AmPm");
Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
Print #PF, String(40, 45); EC(7)
Else
SetPage ReportView.View, Heading
End If
OpenDb "MSaleData", Db
 Sql = "Select * from [Modifiers] order by [Modifier Name],[Modifier Category] asc"
ConnectADORS Sql, Db, Tb

'Home Phone
If Check_Option(96) = 0 Then
Print #PF, "Modifier Name        $         Mod By"; EC(7)
Print #PF, String(40, 45); EC(7)
Else

PrintLine "Modifier Name", 4000
PrintLine "Cost", 6000
PrintLine "Category", 9000
PrintLine "Modified By", -1
DrawLine
End If
Do While Not Tb.EOF
If Check_Option(96) = 0 Then
Print #PF, Trim(Tb(1)); Tab(21); Format(Tb(2), "###.#0"); Tab(29); Left(Tb(3), 11); EC(7)
Else
PrintLine Tb![Modifier Name], 4000
PrintLine Format(Tb(2), "0.00"), 6000
PrintLine Tb![Modifier Category], 9000
PrintLine Tb![Modified By], -1
End If
Tb.MoveNext
Loop
If Check_Option(96) = 0 Then
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
DbClose Db, Tb
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS4")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub


Private Sub PrintModifiers_Click()
On Error GoTo ErrorHandler
Dim PrinterPort As String, Heading As String
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
If Check_Option(96) = 0 Then
Printer_Drivers "Receipt"
CenterText = EC(8)
LeftText = EC(9)
Dim PF As Integer
PF = FreeFile
PrinterPort = "Report"
OneMoreTime:
Open PrinterPort For Output Shared As #PF
Else
SetPrinter 0
End If
Heading = "Modifier Options"
'Plist = PList &
If Check_Option(96) = 0 Then
Print #PF, EC(1); EC(4); CenterText; Heading; EC(2); EC(7)
Print #PF, EC(4); LeftText; EC(7)
Print #PF, ; EC(7)
Print #PF, Format(Time, "HH:MM AmPm");
Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
Print #PF, String(40, 45); EC(5)
Else
SetPage ReportView.View, Heading
End If
OpenDb "MSaleData", Db
OpenTb "Select * from [Modifier Names]", Tb, Db
If Check_Option(96) = 0 Then
    Print #PF, "Modifier Name         Mand   Multi   Opt"; EC(7)
    Print #PF, String(40, 45); EC(5)
Else
    PrintLine "Modifier Name", 4000
    PrintLine "required", 6000
    PrintLine "Multi Select", 7500
    PrintLine "Optional", 8800
    PrintLine "Expand Print", -1
    DrawLine
End If
Dim Clro As String, Xppt As String
Do While Not Tb.EOF
If Check_Option(96) = 0 Then
If Tb(1) = 1 Then Mand = "Y" Else Mand = "N"
If Tb(2) = 1 Then multi = "Y" Else multi = "N"
If Tb(3) = 1 Then opt = "Y" Else opt = "N"
Else
If Tb(1) = 1 Then Mand = "Yes" Else Mand = "No"
If Tb(2) = 1 Then multi = "Yes" Else multi = "No"
If Tb(3) = 1 Then opt = "Yes" Else opt = "No"
If Tb(4) = 1 Then Xppt = "Yes" Else Xppt = "No"

End If
If Check_Option(96) = 0 Then
    Print #PF, Trim(Tb(0)); Tab(25); Mand; Tab(31); multi; Tab(38); opt; EC(7)
Else
    PrintLine Tb(0), 4200
    PrintLine Mand, 6400
    PrintLine multi, 7700
    PrintLine opt, 9000
    PrintLine Xppt, -1
End If
Tb.MoveNext
Loop
If Check_Option(96) = 0 Then
   Print #PF, String(40, 45); EC(7)
   Print #PF, EC(6); EC(5)
   Close #PF
   SendJobToPrinter PrinterPort, "Receipt", False
Else
    DrawLine
    ReportView.View.EndDoc
    ReportView.Show 1
End If
DbClose Db, Tb
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS2")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub PrintStoreMenu_Click()
'''''''''''''
On Error GoTo ErrorHandler
Dim F           As String
Dim H           As String
Dim Data        As String
Dim TSales      As Currency
Dim PrinterPort As String
Dim Heading     As String
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim RCat        As String

If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    CenterText = EC(8)
    LeftText = EC(9)
    Dim PF As Integer
    PF = FreeFile
    PrinterPort = "Reports Menu"
OneMoreTime:
    Open PrinterPort For Output Shared As #PF
End If
Heading = "Menu List"
If Check_Option(96) = 0 Then
    Print #PF, EC(1); EC(4); CenterText; Heading; EC(2); EC(7)
    Print #PF, EC(4); LeftText; EC(7)
    Print #PF, ; EC(7)
    Print #PF, Format(Time, "HH:MM AmPm");
    Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
    Print #PF, String(40, 45); EC(7)
End If
OpenDb "MSaleData", Db
 Sql = "Select * from [MenuItemCountDown] where [Quick Pick] = 0 order by [Report Category],[Menu Item Name] asc"
ConnectADORS Sql, Db, Tb
If Check_Option(96) = 0 Then
    Print #PF, "Item Name                $  Tax  Remote"; EC(7)
    Print #PF, String(40, 45); EC(7)
Else
    F = "+<2200|+<1200|+>1000|+500|+<1200|+<1200|+<1200|+<1200|+<1100|+<1100|+<1100|+<1100;"
    H = "Item Name|PLU|Price|Tax|Mod1|Mod2|Mod3|Mod4|Mod5|Mod6|Rmt 1|Rmt 2;"
End If
Data = ""
Do While Not Tb.EOF
 If Tb![Taxable] = 1 Then
 Taxb = "Y"
Else
 Taxb = "N"
End If
If Tb![remote Printing] = 1 Then
 RemP = "Y"
Else
 RemP = "N"
End If
If Check_Option(96) = 0 Then
 Print #PF, Trim(Tb![Menu Item Name]); Tab(24); Format(SetDualPrice(Tb![Regular price]), "###.#0"); Tab(33); Taxb; Tab(38); RemP; EC(7)
Else
    Data = Data & Trim$(Tb![Menu Item Name]) & "|"
    Data = Data & Tb![Bar Code] & "|"
    Data = Data & Format(SetDualPrice(Tb![Regular price]), "0.00") & "|"
    Data = Data & Taxb & "|"
    Data = Data & Trim$(Tb![Modifier 1] & "") & "|"
    Data = Data & Trim$(Tb![Modifier 2] & "") & "|"
    Data = Data & Trim$(Tb![Modifier 3] & "") & "|"
    Data = Data & Trim$(Tb![Modifier 4] & "") & "|"
    Data = Data & Trim$(Tb![Modifier 5] & "") & "|"
    Data = Data & Trim$(Tb![Remote 1] & "") & "|"
    Data = Data & Trim$(Tb![Remote 2] & "") & ";"
End If
Tb.MoveNext
Loop
DbClose Db, Tb
If Check_Option(96) = 0 Then
  Print #PF, String(40, 45); EC(7)
  Print #PF, EC(6); EC(5)
  Close #PF
  SendJobToPrinter PrinterPort, "Receipt", False
Else
 ReportView.View.Orientation = orLandscape
 
 SetPage ReportView.View, "Menu Items " & Date
 ReportView.View.StartTable
 ReportView.View.TableBorder = tbBoxColumns
 ReportView.View.addTable F, H, Data, vbGreen
 ReportView.View.EndTable
 ReportView.View.Enabled = True
 ReportView.View.EndDoc
 ReportView.Show 1
End If
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS1")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub ShowDates_Click()
End Sub
Private Sub SSCommand1_Click()
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim LT As Long
ConnectADODB SvrPath.TerminalPath & ProductHistory, Db
ConnectADORS "Select Distinct [Menu Item] from [Detail] Group by [Menu Item] Order by [Menu Item] asc ", Db, Tb
MenuItem.Clear
Do While Not Tb.EOF
  MenuItem.AddItem Tb(0) & ""
  Tb.MoveNext
Loop
DbClose Db, Tb
End Sub
Private Sub SSCommand10_Click()
End Sub
Private Sub SSCommand11_Click()
End Sub
Private Sub SSCommand12_Click()
End Sub
Private Sub SSCommand13_Click()
End Sub

Private Sub SSCommand18_Click()
On Error GoTo erbox
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim Up_date As Integer
If ItemList.Text = "" Then Exit Sub
     'OpenDb SvrPath.TerminalPath & "Crmenu.mdb", Db
     If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
     Sql = "Delete from [Ingredient List] Where [Ingredient Item] = '" & ItemList.Text & "'"
     Db.Execute Sql
     DbClose Db, RS
 See_Msg "Removed.....", 5, "Update Complete."
 RemoveFromList ItemList.Text, ItemList
 ItemList.Text = ""
Exit Sub
erbox:
See_Msg Err.Description, 0, "Error"
End Sub
Private Sub SSCommand19_Click()
On Error GoTo erbox
Dim Db      As ADODB.Connection
Dim RS      As ADODB.Recordset
Dim Up_date As Integer
If ItemList.Text = "" Then Exit Sub
'OpenDb SvrPath.TerminalPath & "Crmenu.mdb", Db
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
   Sql = "Select * from [Ingredient List] Where [Ingredient Item] = '" & ItemList.Text & "'"
ConnectADORS Sql, Db, RS
  If RS.EOF Then RS.AddNew
  RS![Ingredient Item] = ItemList.Text
  RS![Ingredient Cost] = Val(ICost.Text & "")
  RS.Update
 DbClose Db, RS
 See_Msg "Saved.....", 5, "Update Complete."
 ItemList.Text = ""
 ICost.Text = 0#
 LoadIngredientList
Exit Sub
erbox:
See_Msg Err.Description, 0, "Error"
End Sub
Private Sub SSCommand2_Click()
End Sub
Private Sub SSCommand27_Click(Index As Integer)
End Sub

Private Sub SSCommand3_Click()
BtnEnh2_Click
End Sub

Private Sub SSCommand4_Click()
Dim Itm As Integer
  For Itm = 0 To Sel_Server.ListCount - 1
     Sel_Server.Selected(Itm) = True
  Next Itm
End Sub
Private Sub SSCommand5_Click()
BtnEnh2_Click
With MenuSelect
CenterObject Me, MenuSelect
.Visible = True
.ZOrder 0
End With
DoEvents

End Sub
Private Sub SSCommand6_Click()
Dim Itm As Integer
  For Itm = 0 To Sel_Server.ListCount - 1
     Sel_Server.Selected(Itm) = False
  Next Itm
End Sub
Private Sub SSCommand7_Click()
End Sub
Private Sub SSCommand8_Click()
Load_View = 1
DoEvents
Dim F As Integer, Itm As Integer
F = 0
ReDim Item_Search(Sel_Server.SelCount)
  For Itm = 0 To Sel_Server.ListCount - 1
    If Sel_Server.Selected(Itm) = True Then
      Item_Search(F) = Sel_Server.List(Itm)
      F = F + 1
    End If
  Next Itm
Ind_Selection = Sel_Server.SelCount
Dim SvrName() As String
ReDim SvrName(0)
If UCase(ContestName.Text) = "GAMES" Then
 TodaysGame 0, STime.Text, ETime.Text, ContestName.Text
 GameReport Format(Date, "MM/DD"), Format(Date, "MM/DD"), STime.Text, ETime.Text, False
Else
 IndSalesMix 0, STime.Text, ETime.Text, IndSales.Value, SvrName(), GroupBy.Value
 PrintCustomMix Format(Date, "MM/DD"), Format(Date, "MM/DD"), STime.Text, ETime.Text, "Week To Date Item Mix", SvrName(), False
End If
End Sub
Private Sub SSCommand9_Click()
MenuSelect.Visible = False
MenuSelect.Left = -4000
Sleep 0
End Sub
Private Sub STime_Click()
Dim rTime As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
rTime = PKT.ReturnTime
If rTime > "" Then
  STime.Text = rTime
Else
  STime.Text = "None"
End If
Set PKT = Nothing
End Sub
Private Sub STime_DblClick()
STime.Text = "None"
End Sub
                                                                                                                          