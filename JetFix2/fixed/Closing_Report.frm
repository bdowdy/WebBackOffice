VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Closing_Report 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Nightly Sales Report"
   ClientHeight    =   10815
   ClientLeft      =   480
   ClientTop       =   480
   ClientWidth     =   14655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10815
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CloseDateEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   -1800
      ScaleHeight     =   6705
      ScaleWidth      =   3165
      TabIndex        =   59
      Top             =   2520
      Visible         =   0   'False
      Width           =   3200
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   855
         Left            =   240
         TabIndex        =   62
         Top             =   600
         Width           =   2655
         _Version        =   589828
         _ExtentX        =   4683
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Continue and Append Sales File"
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
         textCaption     =   "Closing_Report.frx":0000
         textLT          =   "Closing_Report.frx":009C
         textCT          =   "Closing_Report.frx":00B4
         textRT          =   "Closing_Report.frx":00CC
         textLM          =   "Closing_Report.frx":00E4
         textRM          =   "Closing_Report.frx":00FC
         textLB          =   "Closing_Report.frx":0114
         textCB          =   "Closing_Report.frx":012C
         textRB          =   "Closing_Report.frx":0144
         colorBack       =   "Closing_Report.frx":015C
         colorIntern     =   "Closing_Report.frx":0186
         colorMO         =   "Closing_Report.frx":01B0
         colorFocus      =   "Closing_Report.frx":01DA
         colorDisabled   =   "Closing_Report.frx":0204
         colorPressed    =   "Closing_Report.frx":022E
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
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   855
         Left            =   240
         TabIndex        =   63
         Top             =   1440
         Width           =   2655
         _Version        =   589828
         _ExtentX        =   4683
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Move Existing Sales To A New Date"
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
         textCaption     =   "Closing_Report.frx":0258
         textLT          =   "Closing_Report.frx":02FA
         textCT          =   "Closing_Report.frx":0312
         textRT          =   "Closing_Report.frx":032A
         textLM          =   "Closing_Report.frx":0342
         textRM          =   "Closing_Report.frx":035A
         textLB          =   "Closing_Report.frx":0372
         textCB          =   "Closing_Report.frx":038A
         textRB          =   "Closing_Report.frx":03A2
         colorBack       =   "Closing_Report.frx":03BA
         colorIntern     =   "Closing_Report.frx":03E4
         colorMO         =   "Closing_Report.frx":040E
         colorFocus      =   "Closing_Report.frx":0438
         colorDisabled   =   "Closing_Report.frx":0462
         colorPressed    =   "Closing_Report.frx":048C
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
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   855
         Left            =   240
         TabIndex        =   64
         Top             =   2280
         Width           =   2655
         _Version        =   589828
         _ExtentX        =   4683
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Change To A New Close Out Date"
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
         textCaption     =   "Closing_Report.frx":04B6
         textLT          =   "Closing_Report.frx":0552
         textCT          =   "Closing_Report.frx":056A
         textRT          =   "Closing_Report.frx":0582
         textLM          =   "Closing_Report.frx":059A
         textRM          =   "Closing_Report.frx":05B2
         textLB          =   "Closing_Report.frx":05CA
         textCB          =   "Closing_Report.frx":05E2
         textRB          =   "Closing_Report.frx":05FA
         colorBack       =   "Closing_Report.frx":0612
         colorIntern     =   "Closing_Report.frx":063C
         colorMO         =   "Closing_Report.frx":0666
         colorFocus      =   "Closing_Report.frx":0690
         colorDisabled   =   "Closing_Report.frx":06BA
         colorPressed    =   "Closing_Report.frx":06E4
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
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   855
         Left            =   240
         TabIndex        =   65
         Top             =   3120
         Width           =   2655
         _Version        =   589828
         _ExtentX        =   4683
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Over Write Existing Sales"
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
         textCaption     =   "Closing_Report.frx":070E
         textLT          =   "Closing_Report.frx":07A0
         textCT          =   "Closing_Report.frx":07B8
         textRT          =   "Closing_Report.frx":07D0
         textLM          =   "Closing_Report.frx":07E8
         textRM          =   "Closing_Report.frx":0800
         textLB          =   "Closing_Report.frx":0818
         textCB          =   "Closing_Report.frx":0830
         textRB          =   "Closing_Report.frx":0848
         colorBack       =   "Closing_Report.frx":0860
         colorIntern     =   "Closing_Report.frx":088A
         colorMO         =   "Closing_Report.frx":08B4
         colorFocus      =   "Closing_Report.frx":08DE
         colorDisabled   =   "Closing_Report.frx":0908
         colorPressed    =   "Closing_Report.frx":0932
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
         Office2003ColorNorth=   206
         Office2003ColorSouth=   128
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   206
         Office2003ColorSouthMouseOver=   206
         Office2003ColorBorderMouseOver=   1812708
         Office2003ColorNorthPressed=   206
         Office2003ColorSouthPressed=   128
         Office2003ColorBorderPressed=   1812708
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh BtnEnh5 
         Height          =   855
         Left            =   240
         TabIndex        =   73
         Top             =   3960
         Width           =   2655
         _Version        =   589828
         _ExtentX        =   4683
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Cancel Close Out"
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
         textCaption     =   "Closing_Report.frx":095C
         textLT          =   "Closing_Report.frx":09DC
         textCT          =   "Closing_Report.frx":09F4
         textRT          =   "Closing_Report.frx":0A0C
         textLM          =   "Closing_Report.frx":0A24
         textRM          =   "Closing_Report.frx":0A3C
         textLB          =   "Closing_Report.frx":0A54
         textCB          =   "Closing_Report.frx":0A6C
         textRB          =   "Closing_Report.frx":0A84
         colorBack       =   "Closing_Report.frx":0A9C
         colorIntern     =   "Closing_Report.frx":0AC6
         colorMO         =   "Closing_Report.frx":0AF0
         colorFocus      =   "Closing_Report.frx":0B1A
         colorDisabled   =   "Closing_Report.frx":0B44
         colorPressed    =   "Closing_Report.frx":0B6E
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Closing_Report.frx":0B98
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
         Height          =   1695
         Left            =   120
         TabIndex        =   61
         Top             =   4980
         Width           =   2895
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H000040C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   4455
         Left            =   120
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label DateEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "History Date Conflict"
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
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   2895
      End
   End
   Begin BTNENHLib4.BtnEnh SSCommand1 
      Height          =   1335
      Left            =   11640
      TabIndex        =   58
      Top             =   7680
      Width           =   1695
      _Version        =   589828
      _ExtentX        =   2990
      _ExtentY        =   2355
      _StockProps     =   66
      Caption         =   $"Closing_Report.frx":0C74
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
      textCaption     =   "Closing_Report.frx":0C8A
      textLT          =   "Closing_Report.frx":0D0E
      textCT          =   "Closing_Report.frx":0D26
      textRT          =   "Closing_Report.frx":0D3E
      textLM          =   "Closing_Report.frx":0D56
      textRM          =   "Closing_Report.frx":0D6E
      textLB          =   "Closing_Report.frx":0D86
      textCB          =   "Closing_Report.frx":0D9E
      textRB          =   "Closing_Report.frx":0DB6
      colorBack       =   "Closing_Report.frx":0DCE
      colorIntern     =   "Closing_Report.frx":0DF8
      colorMO         =   "Closing_Report.frx":0E22
      colorFocus      =   "Closing_Report.frx":0E4C
      colorDisabled   =   "Closing_Report.frx":0E76
      colorPressed    =   "Closing_Report.frx":0EA0
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
   Begin VB.PictureBox Progression 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   1800
      ScaleHeight     =   1260
      ScaleWidth      =   4125
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   4150
      Begin BTNENHLib4.BtnEnh Progress 
         Height          =   390
         Left            =   240
         TabIndex        =   74
         Top             =   600
         Width           =   3615
         _Version        =   589828
         _ExtentX        =   6376
         _ExtentY        =   688
         _StockProps     =   66
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
         Surface         =   5
         BackColorContainer=   12648384
         LogPixels       =   96
         SpecialEffectFactor=   2
         UserData        =   0.1
         textCaption     =   "Closing_Report.frx":0ECA
         textLT          =   "Closing_Report.frx":0EE2
         textCT          =   "Closing_Report.frx":0EFA
         textRT          =   "Closing_Report.frx":0F12
         textLM          =   "Closing_Report.frx":0F2A
         textRM          =   "Closing_Report.frx":0F42
         textLB          =   "Closing_Report.frx":0F5A
         textCB          =   "Closing_Report.frx":0F72
         textRB          =   "Closing_Report.frx":0F8A
         colorBack       =   "Closing_Report.frx":0FA2
         colorIntern     =   "Closing_Report.frx":0FCC
         colorMO         =   "Closing_Report.frx":0FF6
         colorFocus      =   "Closing_Report.frx":1020
         colorDisabled   =   "Closing_Report.frx":104A
         colorPressed    =   "Closing_Report.frx":1074
         Style           =   8
         ProgressBarValue=   1
         ProgressBarColor=   8421504
      End
      Begin VB.Label UPdateWhere 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Progress"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   3615
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   120
         Top             =   120
         Width           =   3855
      End
   End
   Begin VSFlex8DAOCtl.VSFlexGrid CCards 
      Height          =   3495
      Left            =   6600
      TabIndex        =   51
      Top             =   2160
      Width           =   3255
      _cx             =   1969428077
      _cy             =   1969428501
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
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   5
      GridLinesFixed  =   4
      GridLineWidth   =   1
      Rows            =   100
      Cols            =   3
      FixedRows       =   0
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
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Index           =   7
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   9360
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Index           =   6
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   5
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   4
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   3
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   390
      Index           =   2
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Index           =   1
      Left            =   11640
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox S_Data 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Index           =   0
      Left            =   11640
      MaxLength       =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VSFlex8DAOCtl.VSFlexGrid POuts 
      Height          =   3495
      Left            =   10200
      TabIndex        =   52
      Top             =   2160
      Width           =   3135
      _cx             =   1969427866
      _cy             =   1969428501
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
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   5
      GridLinesFixed  =   5
      GridLineWidth   =   1
      Rows            =   100
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
   Begin BTNENHLib4.BtnEnh SSCommand3 
      Height          =   855
      Left            =   11640
      TabIndex        =   55
      Top             =   9120
      Width           =   1695
      _Version        =   589828
      _ExtentX        =   2990
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
      BackColorContainer=   12632256
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Closing_Report.frx":109E
      textLT          =   "Closing_Report.frx":1106
      textCT          =   "Closing_Report.frx":111E
      textRT          =   "Closing_Report.frx":1136
      textLM          =   "Closing_Report.frx":114E
      textRM          =   "Closing_Report.frx":1166
      textLB          =   "Closing_Report.frx":117E
      textCB          =   "Closing_Report.frx":1196
      textRB          =   "Closing_Report.frx":11AE
      colorBack       =   "Closing_Report.frx":11C6
      colorIntern     =   "Closing_Report.frx":11F0
      colorMO         =   "Closing_Report.frx":121A
      colorFocus      =   "Closing_Report.frx":1244
      colorDisabled   =   "Closing_Report.frx":126E
      colorPressed    =   "Closing_Report.frx":1298
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
   Begin BTNENHLib4.BtnEnh SSCommand5 
      Height          =   735
      Left            =   11640
      TabIndex        =   56
      Top             =   6840
      Width           =   1695
      _Version        =   589828
      _ExtentX        =   2990
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Erase"
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
      textCaption     =   "Closing_Report.frx":12C2
      textLT          =   "Closing_Report.frx":132C
      textCT          =   "Closing_Report.frx":1344
      textRT          =   "Closing_Report.frx":135C
      textLM          =   "Closing_Report.frx":1374
      textRM          =   "Closing_Report.frx":138C
      textLB          =   "Closing_Report.frx":13A4
      textCB          =   "Closing_Report.frx":13BC
      textRB          =   "Closing_Report.frx":13D4
      colorBack       =   "Closing_Report.frx":13EC
      colorIntern     =   "Closing_Report.frx":1416
      colorMO         =   "Closing_Report.frx":1440
      colorFocus      =   "Closing_Report.frx":146A
      colorDisabled   =   "Closing_Report.frx":1494
      colorPressed    =   "Closing_Report.frx":14BE
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
   Begin BTNENHLib4.BtnEnh SSCommand4 
      Height          =   735
      Left            =   11640
      TabIndex        =   57
      Top             =   6000
      Width           =   1695
      _Version        =   589828
      _ExtentX        =   2990
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Calculator"
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
      textCaption     =   "Closing_Report.frx":14E8
      textLT          =   "Closing_Report.frx":155C
      textCT          =   "Closing_Report.frx":1574
      textRT          =   "Closing_Report.frx":158C
      textLM          =   "Closing_Report.frx":15A4
      textRM          =   "Closing_Report.frx":15BC
      textLB          =   "Closing_Report.frx":15D4
      textCB          =   "Closing_Report.frx":15EC
      textRB          =   "Closing_Report.frx":1604
      colorBack       =   "Closing_Report.frx":161C
      colorIntern     =   "Closing_Report.frx":1646
      colorMO         =   "Closing_Report.frx":1670
      colorFocus      =   "Closing_Report.frx":169A
      colorDisabled   =   "Closing_Report.frx":16C4
      colorPressed    =   "Closing_Report.frx":16EE
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
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   4320
      TabIndex        =   76
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label TipWith 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tips Withheld"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   75
      Top             =   5880
      Width           =   2200
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   9360
      TabIndex        =   72
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposits Paid In"
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
      Height          =   255
      Index           =   29
      Left            =   6840
      TabIndex        =   71
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Refunds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   70
      Top             =   5520
      Width           =   2200
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   69
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "DAILY TOTALS"
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
      Height          =   255
      Left            =   6480
      TabIndex        =   68
      Top             =   5880
      Width           =   4800
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Non Sales Revenue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   27
      Left            =   2040
      TabIndex        =   67
      Top             =   5160
      Width           =   2200
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   4320
      TabIndex        =   66
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   54
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   26
      Left            =   2040
      TabIndex        =   53
      Top             =   2280
      Width           =   2200
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   4215
      Left            =   11520
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   50
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Non Cash Deposit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   25
      Left            =   2040
      TabIndex        =   49
      Top             =   4800
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash from Preorders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   24
      Left            =   2040
      TabIndex        =   48
      Top             =   4440
      Width           =   2200
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   47
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Deposit 3"
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
      Index           =   23
      Left            =   2040
      TabIndex        =   44
      Top             =   8040
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Paid Outs"
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
      Height          =   255
      Index           =   22
      Left            =   6840
      TabIndex        =   43
      Top             =   8160
      Width           =   2400
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   9360
      TabIndex        =   42
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   9360
      TabIndex        =   41
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   9360
      TabIndex        =   40
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   9360
      TabIndex        =   39
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   38
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   9360
      TabIndex        =   37
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Over / Short"
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
      Index           =   21
      Left            =   6840
      TabIndex        =   36
      Top             =   9285
      Width           =   2400
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bank Deposits"
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
      Height          =   255
      Index           =   20
      Left            =   6840
      TabIndex        =   35
      Top             =   7680
      Width           =   2400
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Safe Amount"
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
      Index           =   15
      Left            =   2040
      TabIndex        =   34
      Top             =   9480
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Deposit"
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
      Index           =   14
      Left            =   2040
      TabIndex        =   33
      Top             =   9000
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Deposit 4"
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
      Index           =   13
      Left            =   2040
      TabIndex        =   32
      Top             =   8520
      Width           =   2200
   End
   Begin VB.Label Titles 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Deposit 3"
      ForeColor       =   &H00000000&
      Height          =   15
      Index           =   12
      Left            =   2040
      TabIndex        =   31
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Deposit 2"
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
      Index           =   11
      Left            =   2040
      TabIndex        =   30
      Top             =   7560
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash from Audits"
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
      Index           =   10
      Left            =   2040
      TabIndex        =   29
      Top             =   7080
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discounts"
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
      Height          =   255
      Index           =   19
      Left            =   6840
      TabIndex        =   28
      Top             =   6720
      Width           =   2400
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Voids"
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
      Height          =   255
      Index           =   18
      Left            =   6840
      TabIndex        =   27
      Top             =   6240
      Width           =   2400
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Non Cash Tenders"
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
      Height          =   255
      Index           =   17
      Left            =   6840
      TabIndex        =   26
      Top             =   7200
      Width           =   2475
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Voids"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   9600
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Titles 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "NON CASH TENDERS"
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
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   23
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Titles 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "PAID OUTS"
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
      Height          =   255
      Index           =   7
      Left            =   10200
      TabIndex        =   22
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Index           =   6
      Left            =   4320
      TabIndex        =   21
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   20
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   19
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   16
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Sales_data 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net Liquor Tax"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   14
      Top             =   3720
      Width           =   2200
   End
   Begin VB.Label Titles 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash to Account For"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   13
      Top             =   6240
      Width           =   2200
      WordWrap        =   -1  'True
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Voids"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      Top             =   2640
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net Discounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   11
      Top             =   4080
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net Special Tax"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net Sales Tax"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   2200
   End
   Begin VB.Label Titles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
      Width           =   2200
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Sales Close Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   6480
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   6600
      Top             =   9135
      Width           =   4545
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   4215
      Left            =   6480
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   1920
      Top             =   6720
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   8280
      Left            =   1800
      Top             =   1800
      Width           =   4575
   End
End
Attribute VB_Name = "Closing_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CancelClose          As Boolean
Public Clear_Text           As Integer
Public CashtoAcctFor        As Currency
Public EraseOk              As Boolean
Public TotalPO              As Currency
Public TotalTip             As Currency
Public Misc_Price           As Single
Public ChargeBack           As Currency
Public NotTaxed             As Currency
Public AcctFor              As Currency
Public Gross_Total          As Currency
Public CashDep              As Currency
Public TotalCharges         As Currency
Public TDepositCollected    As Currency
Public GratuityCollected    As Currency
Public GiftCertSold         As Currency
Public TotalRefunds         As Currency
Public CashTenders          As Currency
Public CashTenderQty        As Long
Public CloseDate            As String
Public OverWriteSales       As Boolean
Public CreditCardDep        As Currency
Public CheckDeposit         As Currency
Dim Void                    As FinancialItems
Dim JobTipOut               As FinancialItems
Dim ListRefunds             As FinancialItems
'Public WithEvents HSFTP     As SFTPLib.SFTP
Public SvrTipsCollect       As Currency
Public Function HotScheduleExport() As String
If Val(GetSetting("Micro$ale", "Register Setup", "HotSchedules")) = 0 Then Exit Function
Dim Db       As ADODB.Connection
Dim RS       As ADODB.Recordset
Dim FileName As String
Dim CPANYID  As String
Dim CCEPTID  As String
Dim LOCID    As String
Dim HireDate As String
Dim DOB      As String
Dim ff       As Integer
Dim EMPID    As String
Dim EMPNme() As String
Dim TXTFileInfo() As Variant
Dim SalesDate     As String
Dim NickName      As String
Dim Term          As String
Dim TValue        As Integer
On Error Resume Next
Kill SvrPath.TerminalPath & "Hot Schedules\*.*"
Kill SvrPath.TerminalPath & "Hot Schedules DownLoad\*.*"
On Error GoTo ERH
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "Hot Schedules.INI"), TXTFileInfo()) Then
 CPANYID = TXTFileInfo(4)
 CCEPTID = TXTFileInfo(5)
 LOCID = TXTFileInfo(6)
End If
FileName = Format(Date, "_YYYYMMDD") & Format(Time, "_HHMMSS")
If CPANYID = "" Or CCEPTID = "" Or LOCID = "" Then
   HotScheduleExport = "No Setup Information"
   Exit Function
End If
If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Db) = False Then
  HotScheduleExport = "Error Connecting to Check Stats"
  Exit Function
End If
 
 Sql = "Select * from [Sales] order by TRY_CAST([Date] AS datetime) desc"
 If ConnectADORS(Sql, Db, RS) = False Then
   HotScheduleExport = "Error Connecting to Closing History"
   Exit Function
 End If
 If Not RS.EOF Then
    SalesDate = Format(RS.fields("Date").Value, "MM/DD/YYYY")
 Else
    SalesDate = Format(Date, "MM/DD/YYYY")
 End If
 DbClose Db, RS
 
 If ConnectADODB(SvrPath.TerminalPath & ProductHistory, Db) = False Then
   HotScheduleExport = "Error Connecting to Product History"
   Exit Function
 End If
 
 Sql = ReturnDateString("Select * from [Check Header] where TRY_CAST([date] AS datetime) between '" & Format(CDate(CDate(SalesDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(SalesDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'")
 If ConnectADORS(Sql, Db, RS) = False Then
    HotScheduleExport = "Error Connecting to Employee Records"
   Exit Function
 End If
 ff = FreeFile
 SalesDate = Format(SalesDate, "YYYY-MM-DD hh:mm:ss.000")
 Open SvrPath.TerminalPath & "Hot Schedules\" & "Sales" & FileName & ".TXT" For Output As ff
 Do While Not RS.EOF
   Print #ff, CPANYID & "|" & CCEPTID & "|" & LOCID & "|" & SalesDate & "|" & Format(RS.fields("Date").Value, "YYYY-MM-DD hh:mm:ss.000") & "|" & Format(((RS.fields("Check Total").Value - RS.fields("Tax Total").Value) + Abs(RS.fields("Discount Total").Value)) - Abs(RS.fields("Non Sales Total").Value), "0.00") & "|" & "1" & "|" & GetOrderDestNo(RS.fields("Order Destination").Value & "") & "|" & "1"
   RS.MoveNext
 Loop
 Close ff
 DbClose Db, RS
 If ConnectADODB("MSaleData", Db) = False Then
   HotScheduleExport = "Error Connecting to Employee Records"
   Exit Function
 End If
 Sql = "Select * from [General Data] order by [Last Name] asc"
 If ConnectADORS(Sql, Db, RS) = False Then
   HotScheduleExport = "Error Connecting to Employee Records"
   Exit Function
 End If
'// start employee export
If Not RS.EOF Then
    ff = FreeFile
    Open SvrPath.TerminalPath & "Hot Schedules\" & "empl_master" & FileName & ".TXT" For Output As ff
    Do While Not RS.EOF
      NickName = RS.fields("Nick Name").Value & ""
      If NickName = "None" Then NickName = ""
      HireDate = ""
      HireDate = RS.fields("Hire Date").Value & ""
      If IsDate(HireDate) Then HireDate = Format(RS.fields("Hire Date").Value, "YYYY-MM-DD hh:mm:ss.000")
      DOB = RS.fields("DOB").Value & ""
      If IsDate(DOB) Then DOB = Format(RS.fields("DOB").Value, "YYYY-MM-DD hh:mm:ss.000")
      Term = RS.fields("Terminated").Value & ""
      If IsDate(Term) Then
        Term = Format(RS.fields("Terminated").Value, "YYYY-MM-DD hh:mm:ss.000")
        TValue = -1
      Else
        Term = ""
        TValue = 1
      End If
      Print #ff, CPANYID & "|" & CCEPTID & "|" & LOCID & "|" & RS.fields("Company Id").Value & "" & "|" & RS.fields("First Name").Value & "" & "|" & RS.fields("Last Name").Value & "" & "|" & RS.fields("Home Phone").Value & "" & "|" & _
                  "" & "|" & RS.fields("Phone # 2").Value & "" & "|" & RS.fields("Address").Value & "" & "|" & RS.fields("City").Value & "" & "|" & RS.fields("State").Value & "" & "|" & RS.fields("Zip").Value & "" & "|" & _
                  Term & "|" & NickName & "|" & HireDate & "|" & RS.fields("Company Id").Value & "" & "|" & DOB & "|" & TValue & "|"
      RS.MoveNext
    Loop
    SetNothing RS
    Close ff
End If
'//  end general data export
Sql = "Select * from [Job Codes] order by [Last Name] asc"
If ConnectADORS(Sql, Db, RS) = False Then
   HotScheduleExport = "Error Connecting to Employee Records"
   Exit Function
End If
'// start job codes export
If Not RS.EOF Then
    ff = FreeFile
    Open SvrPath.TerminalPath & "Hot Schedules\" & "empl_job" & FileName & ".TXT" For Output As ff
    Do While Not RS.EOF
      EMPID = GetEmpId(RS.fields("First Name").Value & "", RS.fields("Last Name").Value & "")
      Print #ff, CPANYID & "|" & CCEPTID & "|" & LOCID & "|" & EMPID & "|" & ReturnJobCode(RS.fields("Job Code").Value & "") & "|" & Format(Val(RS.fields("Rate").Value & ""), "0.00") & "|1.5|Y"
      RS.MoveNext
    Loop
    DbClose Db, RS
    Close ff
End If
'// end employee export
Time_Reports.WriteHotSchedules CPANYID, CCEPTID, LOCID, FileName
MemoryCheck Time_Reports
''=======================================================
'Dim R   As String
'Dim Zip As MSaleData.Publish
'Set Zip = New MSaleData.Publish
'R = Zip.CreateZipFile(SvrPath.TerminalPath & "Hot Schedules\", SvrPath.TerminalPath & "Hot Schedules\" & FilePrefix & FileName & ".ZIP")
'Set Zip = Nothing

'ZipFile.ArchiveFile = SvrPath.TerminalPath & "Hot Schedules\" & FilePrefix & FileName & ".ZIP"
'ZipFile.RecurseSubdirectories = True
'If FileCheck(SvrPath.TerminalPath & "Hot Schedules\" & "empl_master" & FileName & ".TXT") Then
' ZipFile.IncludeFiles SvrPath.TerminalPath & "Hot Schedules\" & "empl_master" & FileName & ".TXT"
' ZipFile.Compress
'End If
'If FileCheck(SvrPath.TerminalPath & "Hot Schedules\" & "timecard" & FileName & ".TXT") Then
' ZipFile.IncludeFiles SvrPath.TerminalPath & "Hot Schedules\" & "timecard" & FileName & ".TXT"
' ZipFile.Compress
'End If
'If FileCheck(SvrPath.TerminalPath & "Hot Schedules\" & "empl_job" & FileName & ".TXT") Then
' ZipFile.IncludeFiles SvrPath.TerminalPath & "Hot Schedules\" & "empl_job" & FileName & ".TXT"
' ZipFile.Compress
'End If
'If FileCheck(SvrPath.TerminalPath & "Hot Schedules\" & "sales" & FileName & ".TXT") Then
' ZipFile.IncludeFiles SvrPath.TerminalPath & "Hot Schedules\" & "sales" & FileName & ".TXT"
' ZipFile.Compress
'End If
''FileName_20160329_135634
''Set ZipFile = Nothing
'Dim FTP As MSaleData.Publish
'Set FTP = New MSaleData.Publish
'    FTP.SFTP FilePrefix, TXTFileInfo(2), TXTFileInfo(3), SvrPath.TerminalPath & "Hot Schedules\" & FilePrefix & FileName & ".ZIP", FilePrefix & FileName & ".ZIP", SvrPath.TerminalPath & "Hot Schedules DownLoad\Shifts.TXT"
'Set FTP = Nothing

'Set HSFTP = New SFTPLib.SFTP
'With HSFTP
'.RemotePath = ""
'        .TimeOut = 120
'        .RemotePath = ""
'        .RemoteFile = ""
'        .SSHAuthMode = amPassword
'        .SSHUser = TXTFileInfo(2)
'        .SSHPassword = TXTFileInfo(3)
'        .SSHLogon TXTFileInfo(0), 22
'        .RemotePath = TXTFileInfo(1)
'        .RemoteFile = FilePrefix & FileName & ".ZIP"
'        .LocalFile = SvrPath.TerminalPath & "Hot Schedules\" & FilePrefix & FileName & ".ZIP"
'        .Upload
'        .RemoteFile = "Shifts.TXT"
'        .LocalFile = SvrPath.TerminalPath & "Hot Schedules DownLoad\Shifts.TXT"
'        On Error Resume Next
'        .Download
'        .SSHLogoff
'End With
'Set HSFTP = Nothing
HotScheduleExport = "Success"
Exit Function
ERH:
HotScheduleExport = Err.Description
Trans_Log "Error Loading To Hot Schedules " & Err.Description
See_Msg Err.Description, 0, "Error sending hot schedule files"
End Function
Public Function SendCompeat(ExportDate As String) As String
On Error Resume Next
See_Msg "Building Export File", 5, "Compeat"
If Dir(SvrPath.TerminalPath & "Compeat", vbDirectory) = "" Then
 MkDir SvrPath.TerminalPath & "Compeat"
Else
 Kill SvrPath.TerminalPath & "Compeat\*.*"
End If
On Error GoTo ERH
Dim FTPSite As String
Dim CUserN  As String
Dim CUserP  As String
Dim CFile   As String
Dim R       As String
Dim TXTFileInfo() As Variant
Dim CTFile  As Ctuit.CollectSales
Set CTFile = New Ctuit.CollectSales
R = CTFile.BuildSaleExport(ExportDate, 0)
Set CTFile = Nothing
If R = "Done" Then

Else
  Trans_Log "Compeat Export Failed " & R
  Exit Function
End If
CFile = Dir(SvrPath.TerminalPath & "Compeat\*.xml")
'CFile = "[" & GetSetting("Micro$ale", "Register Setup", "Compeat CID") & "]_[" & Format(ExportDate, "YYYYMMDD") & "].xml"
See_Msg "File Built " & R, 5, CFile
On Error GoTo ERH
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "Compeat.INI"), TXTFileInfo()) Then
 FTPSite = TXTFileInfo(0)
 CUserN = TXTFileInfo(1)
 CUserP = TXTFileInfo(2)
End If
'Set HSFTP = New SFTPLib.SFTP
'With HSFTP
'        .RemotePath = ""
'        .RemotePath = ""
'        .RemoteFile = ""
'        .SSHAuthMode = amPassword
'        .SSHUser = CUserN
'        .SSHPassword = CUserP
'        .SSHLogon TXTFileInfo(0), 22
'        .RemotePath = ""
'        .RemoteFile = CFile
'        .LocalFile = SvrPath.TerminalPath & "Compeat\" & CFile
'        .Upload
'        .SSHLogoff
'End With
'Set HSFTP = Nothing
SendCompeat = "Success"
Exit Function
ERH:
SendCompeat = Err.Description
See_Msg Err.Description, 0, "Error sending compeat files"
End Function
Public Sub LoadJobCodes()
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim JC As Integer
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
  Sql = "Select * from [Job Titles] order by [Job Title] asc"
ConnectADORS Sql, Db, Tb
  JC = Tb.RecordCount
  Erase Dept_Pay()
  ReDim Dept_Pay(JC)
   JC = -1
   Do While Not Tb.EOF
      JC = JC + 1
      Dept_Pay(JC).DeptName = Tb.fields("Job Title").Value & ""   ' Dept Name
      Dept_Pay(JC).DeptDollars = 0
      Dept_Pay(JC).TipPostion = Val(Tb.fields("Tipped Position").Value & "")   'Tip Position
      Dept_Pay(JC).TipsOut = Val(Tb.fields("Tip Pool Percent").Value & "")    'Tip Pool Perc
      Dept_Pay(JC).TotatHours = 0
      Tb.MoveNext
   Loop
DbClose Db, Tb
End Sub
Public Sub AutoCloseout()
On Error GoTo ErrorHandler
If Sales_data(0).Caption = 0 Then
  Trans_Log "Auto Close (No Sales Found)"
  See_Msg "Nothing to Close out", 5, ""
  AutoCloseInProcess = False
  Call MemoryCheck(Closing_Report)
  Exit Sub
End If
UpdateOnLineTip
'Dim OverWriteSales As Boolean
'RemoveMenuLink
'If GetConfigInfo("General", "Auto Close Appends Sales", "No", "Functions.INI") = "ENABLED" Then
'   OverWriteSales = False
'   Trans_Log "Auto Close Append = Yes"
'Else
'End If
OverWriteSales = True
Dim DoReports(10)   As Integer
Dim ff              As Integer
Dim InVar           As Integer
Dim IV              As Integer
Dim TXTFileInfo()   As Variant
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "ClosingReport.ini"), TXTFileInfo()) Then
    For IV = 0 To UBound(TXTFileInfo, 1)
        DoReports(IV) = TXTFileInfo(IV)
    Next IV
Else
    For IV = 0 To 10
      DoReports(IV) = 0
    Next IV
End If
Trans_Log "Auto Close Started"
StartClosing
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim DOW         As Integer
Dim Prec        As Integer
'Dim vData()     As Variant
'Dim coupons()   As Variant
Dim FileDate    As String
ClosingDay = True
'CloseDate = GetDate
CloseDate = FindDate
If Not IsDate(CloseDate) Then CloseDate = Format(Date - 1, "MM/DD/YYYY")
Sale_Mix Today, NoPrint, "", "", "", "None", 0
'Call MemoryCheck(Red_Book)
CenterObject Me, Progression
Progression.Visible = True
PAgain:
Prec = 0
'If GetConfigInfo("General", "Sync Sales Data", "OFF", "Functions.INI") = "ENABLED" Then
' MatchClosedChecks
'End If
If Val(S_Data(2).Text) = 0 Then S_Data(2).Text = GetBankDeposits
Sales_Breakdown False, 0
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
Hourly_Sales_Data CloseDate
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
If DoReports(1) = 0 Then Print_Hourly_Sales False, CloseDate
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
If DoReports(4) = 0 Then Print_Meal_Period False
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
Dept_Report "", "Daily " & Format(Date, "M/D/YY"), DoReports(2), "Day", "", 0, False
'Erase vData()
'ReDim vData(2, 1)
If DoReports(5) = 0 Then PrintServeTime False, CloseDate
If DoReports(6) = 0 Then DeptVoid "", True, False
Sales_Report 0, DoReports(7), CloseDate
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
Get_Voids
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
Nightly_Sales DoReports(8), CloseDate
PrintDailyInventory True
DailySalesMix
If SiteLic.Berg = 1 Then
 BergExceptions CloseDate
End If
'If Check_Option(65) = 0 Then
 If System_config = FullService Then
  If AreAuditsOpen("") Then
   See_Msg "Audits Open Cannot Reset Day !!", 5, ""
   Trans_Log "Reset Abort - Audits Open"
   Exit Sub
  End If
 End If
'End If
Prec = Prec + 6: Progress.ProgressBarValue = Prec:
Progression.Visible = False
DOW = 0
SaveRefundsToHistory ListRefunds, CloseDate

StoreTaxedSales CloseDate, OverWriteSales
StoreNoSale CloseDate, OverWriteSales
StoreDayHistory OverWriteSales, CloseDate
SaveDepartment CloseDate, OverWriteSales
Calculate_Labor CloseDate
StoreHeader CloseDate, OverWriteSales
StoreRefunds CloseDate, OverWriteSales
CollectWithHeldTips CloseDate
CollectTipOutWithHeld CloseDate, JobTipOut
Trans_Log "Auto Close Starting Financial Reset"
Progression.Visible = True
CloseWeek = ResetEndOfDay(CloseDate)
Progression.Visible = False
SendPayroll
'Call MemoryCheck(Red_Book)
If Check_Option(61) = 1 Then
 ExportFinancialData
End If
ClosingDay = False
SendEmails(5).Send = True
SendEmails(4).Send = True
If SendEmails(5).Send Then
  Trans_Log "Sending Balance Sheet From Auto Close"
  Display.DSRBalanceSheet CStr(CDate(CloseDate)), CStr(CDate(CloseDate)), SendEmails(5).Send, DoReports(9)
End If
If SendEmails(4).Send Then
  Display.DailySummary CStr(CDate(CloseDate)), SendEmails(4).Send, DoReports(9)
End If
Call MemoryCheck(Display)
EndClosing
Trans_Log "Auto Close Finished"
See_Msg "Close Out Complete....", 5, "System Reset for Next Day"
Trans_Log "Daily Closing Completed..."
Trans_Log "Batch Process Started"
CloseBatch False
Trans_Log "Daily Batch Complete"
SendEMailEOD
If CloseWeek Or Format(CloseDate, "dddd") = Input_Option(60) Then '// make sure cant run longer than 7 days
  See_Msg "Closing Week Sales ", 5, "Resetting Weekly Sales Data"
  Trans_Log "Weekly Closing Finished with auto close "
  Transaction.WeekClose.DoClick
  CloseWeek = False
End If
'// added this line QCR qsr sites not emailing on auto close .
If Val(GetSetting("ProductID", "License File", "General Data 24")) <> 0 Then
 Trans_Log "Export to Hot Schedules " & HotScheduleExport
End If
If Val(GetSetting("Micro$ale", "Register Setup", "Ctuit")) <> 0 Then
   Trans_Log "Export to Compeat " & SendCompeat(CloseDate)
End If
If Val(GetSetting("Micro$ale", "Register Setup", "QuickBooks")) <> 0 Then
  Trans_Log "Export to quick books " & SendQBexport(CloseDate, CloseDate)
End If
Call MemoryCheck(Closing_Report)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Auto Close Out ")
   If Err.Number = 401 Then Resume Next
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub CloseBatch(ShowList As Boolean)
     On Error Resume Next
     Call MemoryCheck(ListCards)
     See_Msg "Batching Charges", 5, "Submitting Batch.."
     Trans_Log "Submitting Batch"
     ListCards.SSCommand3.Enabled = True
     ListCards.SSCommand21.Enabled = True
     ListCards.BatchCharges
     If ShowList Then ListCards.Show 1
     Call MemoryCheck(ListCards)
     Trans_Log "Returned from Batch Send"
End Sub
Private Sub BtnEnh1_Click()
 If See_Msg("Append Sales To Existing File ?", 4, "Adds New Sales With Old Sales") = 6 Then
  CloseDateEdit.Visible = False
  OverWriteSales = False
  Trans_Log "Continue And Append Sales"
 End If
End Sub
Private Sub BtnEnh2_Click()
Dim InDate As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
InDate = PKT.ReturnDate
Set PKT = Nothing
If IsDate(InDate) Then
  If See_Msg("Move Existing Sales To " & InDate, 4, "Moving Sales") = 6 Then
    Display.DateChange Format(CloseDate, "MM/DD/YYYY"), Format(InDate, "MM/DD/YYYY")
    CloseDateEdit.Visible = False
    Trans_Log "Moving Existing Sales"
  End If
End If
End Sub
Private Sub BtnEnh3_Click()
Dim InDate As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
InDate = PKT.ReturnDate
Set PKT = Nothing
If IsDate(InDate) Then
  If See_Msg("Change Close Out Date To " & InDate, 4, "Changing Sales Date") = 6 Then
    CloseDate = Format(InDate, "MM/DD/YYYY")
    CloseDateEdit.Visible = False
    Trans_Log "Changing to a new date Sales"
  End If
End If
End Sub
Private Sub BtnEnh4_Click()
If See_Msg("Over Write Sales  ?", 4, "Over Write Existing History") = 6 Then
OverWriteSales = True
CloseDateEdit.Visible = False
Trans_Log "Over Write Sales"
End If
End Sub
Private Sub BtnEnh5_Click()
CancelClose = True
CloseDateEdit.Visible = False
End Sub
Private Sub CCards_KeyPress(KeyAscii As Integer)
'With CCards
'If KeyAscii = 8 And Len(.Text) > 0 Then
'.TextMatrix(.Row, 1) = Mid(.TextMatrix(.Row, 1), 1, Len(.TextMatrix(.Row, 1)) - 1)
'End If
'If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46 Then
'.TextMatrix(.Row, 1) = .TextMatrix(.Row, 1) & Chr(KeyAscii)
' KeyAscii = 0
'End If
'Totals(2).Caption = "0.00"
'For TP = 0 To .rows - 1
' Totals(2).Caption = Format(Val(Totals(2).Caption) + Val(.TextMatrix(TP, 1)), "0.00")
'Next TP
'Figure_Shortage
'End With
End Sub
Public Sub ClearSevenshiftId()
Dim Db As ADODB.Connection
ConnectADODB SvrPath.TerminalPath & ReceiptDb, Db
 Sql = "delete from [ReceiptData] "
Db.Execute Sql
DbClose Db, Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendAKey VK_TAB, vbKeyTab
    KeyAscii = 0
End If
End Sub
Public Function ResetEndOfDay(CloseDate As String) As Boolean
On Error GoTo ErrorHandler
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim DOW         As Integer
Dim FileDate    As String
Dim Perc        As Integer
See_Msg "Starting Financial Reset", 5, "Working...."
DailyCheckStatBackUp
If Check_Option(61) = 1 Then
  If Check_Option(92) = 1 Then
   Dim Exp As ExportData
   Set Exp = New ExportData
   ' Exp.CreateSalesFile  SvrPath.TerminalPath, SvrPath.TerminalPath, 0, 1
   Set Exp = Nothing
  End If
End If
If System_config = FullService Then
  See_Msg "Saving Server Sales Data", 5, ""
  StoreFinancial
  Trans_Log "Update Server Sales Data"
End If
ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  Sql = "Select * from [closed checks] where [close out day] = 0 order by [check number] asc"
ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   DOW = GetDateToNo(CDate(Tb.fields("Order Date").Value))
   FileDate = Format(Tb.fields("Order Date").Value, "MMDDYYYY")
 End If
 DbClose Db, Tb
 If FileDate = "" Then FileDate = Format(Date, "MMDDYYYY")
 On Error GoTo ErrorHandler
  Dim RSNSale As GFunctions
  Set RSNSale = New GFunctions
   UPdateWhere.Caption = "Closing " & SvrPath.TerminalName
   UPdateWhere.Refresh
   CenterObject Me, Progression
   Progression.Visible = True
   Progression.Refresh
   RSNSale.NoSales SvrPath.TerminalPath, ResetNoSale, ""
   SQL_DailyBackUp DOW & "-" & SvrPath.TerminalName & "-"
   
   If FileCheck(SvrPath.TerminalPath & "QSRKVS.MDB") = True Then
    ConnectADODB SvrPath.TerminalPath & "QSRKVS.MDB", Db
    Sql = "Delete from [Header]"
    Db.Execute Sql
    Sql = "Delete from [QSRItems]"
    Db.Execute Sql
    SetNothing Db
   End If
   Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents

   ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Delete from [Printed Receipts]"
   Db.Execute Sql
    Sql = "Delete from [Whos Here]"  'being use for Book Marks in EnterPrize
   Db.Execute Sql
    Sql = "Delete from [Active Coupons]"
   Db.Execute Sql
    Sql = "Update [Closed Checks] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] = 0;"
   Db.Execute Sql
    Sql = "Update [Check Detail] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] = 0;"
   Db.Execute Sql
    Sql = "Update [Payments] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] = 0;"
   Db.Execute Sql
    Sql = "Update [Error Correct] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] = 0;"
   Db.Execute Sql
    Sql = "Update [Item Refund] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] = 0;"
   Db.Execute Sql
    Sql = "Delete from [Have One On Me]"
   Db.Execute Sql
    Sql = "Delete from [Check Events]"
   Db.Execute Sql
   SetNothing Db
Set RSNSale = Nothing
Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
RecordJobCodeTips CloseDate
On Error Resume Next
DailyDebitBackUp Val(DatePart("d", CloseDate))
DailyBackUp DOW
Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
ConnectADODB SvrPath.TerminalPath & CallInOrderDb, Db
 Sql = "Delete from [Delivery]"
Db.Execute Sql
 Sql = "delete from [Deposits Paid In]"
Db.Execute Sql
 Sql = "Delete from [Check Numbers]"
Db.Execute Sql
 Sql = "Delete from [PrePaid] Where [OrderNumber] <> 0"
Db.Execute Sql

SetNothing Db
Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
  Sql = "Delete from [Check Numbers]"
 Db.Execute Sql
 
If GetConfigInfo("Phone Orders", "Reset Check Numbers Daily", "No", "Functions.INI") = "ENABLED" Then
  Sql = "Delete from [SharedCheckNumbers] "
 Db.Execute Sql '// added in version 1552
Else
  Sql = "Select * from [SharedCheckNumbers] Where Name = 'PhoneOrders'"
 ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   If Val(Tb.fields("Check Number").Value & "") >= 700 Then
      Tb.fields("Check Number").Value = 0
      Tb.Update
   End If
   Tb.Close
 End If
 Sql = "Delete from [SharedCheckNumbers] Where [Name] <> 'PhoneOrders'"
 Db.Execute Sql
End If
SetNothing Db
ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
 Sql = "Delete from [Reservations] "
 Sql = Sql & " Where TRY_CAST([Date] AS datetime) < '" & CDate(Date) & "'"
Db.Execute Sql
SetNothing Db

ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
 Sql = "delete from [CustomerIsWaiting] "
Db.Execute Sql
 Sql = "delete from [LoyaltyId]"
Db.Execute Sql
 Sql = "delete from [GroupAudit]"
Db.Execute Sql
SetNothing Db
ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
   Sql = "Delete from [PMS Charges] "
  Db.Execute Sql
   Sql = "Delete from [Bank Deposits] "
  Db.Execute Sql
   Sql = "Delete from [Pre Auth] "
  Db.Execute Sql
   Sql = "Delete from [Pours]"
  Db.Execute Sql
 '///////////////////////////
   Sql = "Delete from [DayPayments]"
  Db.Execute Sql
   Sql = "Update [Paid Ins] set [Close Out Day] = " & DOW & " where [close out day] = 0"
  Db.Execute Sql
   Sql = "Delete from [Audits]"
  Db.Execute Sql
   Sql = "Delete from [Charge Back]"
  Db.Execute Sql
   Sql = "Update [Paid Outs] set [Close Out Day] = " & DOW & " where [close out day] = 0"
  Db.Execute Sql
    Sql = "Update [Refund] set [close out day] = " & DOW & " where [close out day] = 0"
  Db.Execute Sql

 SetNothing Db
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 On Error GoTo ErrorHandler
   ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
    Sql = "Update [AuditSales] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] < 1" ' -1 is audits on hold
   Db.Execute Sql
    Sql = "Update [Tip Share] Set [Close out day] = " & DOW
    Sql = Sql & " where [Close out day] = 0"
   Db.Execute Sql
   SetNothing Db
   If System_config = FullService Then
     ConnectADODB SvrPath.TerminalPath & "MSaleData", Db
     Sql = "Update [Wait List] Set [day] = " & DOW
     Sql = Sql & " where [day] = 0;"
     Db.Execute Sql
     SetNothing Db
   End If
   SetCloseDayDebit DOW
   Reset_HouseCharges DOW, CloseDate
   DeleteAlerts
   ClearSevenshiftId
   SQLDailyBackUp RemoteBackup_Path
   If Format(CloseDate, "dddd") = Input_Option(60) Or DaysSinceClose > 7 Then
     ResetEndOfDay = True
   Else
     ResetEndOfDay = False
   End If
Exit Function
ErrorHandler:
If Err.Number = 3365 Then
  Trans_Log "Daily Close " & Sql
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reset End Of Day ")
 Select Case ErrorRespose
   Case "Abort"
     Exit Function
   Case Is = "Resume"
     Resume
   Case Is = "Resume Next"
     Resume Next
 End Select
End Function
Public Sub Calculate_Labor(ClosingDate As String)
On Error GoTo ErrorHandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
LoadJobCodes
Dim LabPer          As Double
Dim Ttl_L           As Currency
Dim Ttl_P           As Double
Dim WrkTime         As Single
Dim EmpTwt          As Single
Dim TWRKTIME        As Single
Dim Total_Time      As Single
Dim Db              As ADODB.Connection
Dim Tb              As ADODB.Recordset
Dim PRTb            As ADODB.Recordset
Dim EDb             As ADODB.Connection
Dim Id              As ADODB.Recordset
Dim CHDB            As ADODB.Connection
Dim CHRS            As ADODB.Recordset
Dim PaidHourly      As Integer
Dim Hoursworked     As String
Dim PRate           As Currency
Dim Pl              As Integer
Dim JCS             As Integer
Dim Ck              As Integer
Dim InOverTime      As Boolean
Dim SearchName      As String
For Pl = 0 To UBound(Dept_Pay, 1) - 1
  Dept_Pay(Pl).DeptDollars = 0
Next Pl
If Not IsDate(ClosingDate) Then ClosingDate = Format(Date - 1, "MM/DD/YYYY")
TWRKTIME = 0
ConnectADODB "MSaleData", EDb
 Sql = "Select * From [General Data] Where ISNULL([Terminated], 0) = 0 Order By [Last Name] Asc"
ConnectADORS Sql, EDb, Id
ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
  
  Do While Not Id.EOF
    SearchName = Id(0) & " " & Id(2)
    EmpTwt = 0
    Sql = "Select * From [Job Codes] Where [First Name] = '" & Id(0) & "'"
    Sql = Sql & " and [Last Name] = '" & Id(2) & "'"
    ConnectADORS Sql, EDb, PRTb
        Erase PayRate()
        ReDim PayRate(RecordsCount(PRTb))
            JCS = 0
            Do While Not PRTb.EOF
              PayRate(JCS).JobCodes = PRTb![Job Code] & ""
              PayRate(JCS).JobRate = Val(PRTb![rate] & "")
              PayRate(JCS).JobHourly = Val(PRTb![Hourly] & "")
              JCS = JCS + 1
             PRTb.MoveNext
            Loop
    SetNothing PRTb
    InOverTime = IHaveOverTime(SearchName, ClosingDate)

    If PayRate(0).JobHourly = 0 Then
      Cal_LaborDollars PayRate(0).JobCodes, PayRate(0).JobRate, WrkTime, PayRate(0).JobHourly, 1, 0
    Else
     Sql = "Select * From [Employee Records] Where [Employee Name] = '" & SearchName & "'"
     Sql = Sql & " AND TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(ClosingDate) & " " & GetStartTime & "'"
     Sql = Sql & " and '" & SetEndingDate(CDate(ClosingDate)) & " " & SetEndingTime & "'"
     ConnectADORS Sql, Db, Tb
       Do While Not Tb.EOF
         Hoursworked = CalculateTimeRecord(CStr(Tb(2) & "") & "", CStr(Tb(7) & "") & "", CStr(Tb(3) & ""), CStr(Tb![Break Out] & ""), CStr(Tb![Break In] & ""), CStr(Tb(6) & ""), Total_Time)
         PaidHourly = 0
          If Val(Tb![Rate of Pay] & "") = 0 Then
           For Ck = 0 To JCS - 1
            If Trim(Tb(1)) = Trim(PayRate(Ck).JobCodes) Then
              PRate = PayRate(Ck).JobRate 'PayRate(Ck).JobRate
              PaidHourly = Abs(PayRate(Ck).JobHourly)
             Exit For
            End If
           Next Ck
          Else
           PRate = Val(Tb![Rate of Pay] & "")
           PaidHourly = Abs(Val(Tb![Paid By] & ""))
          End If
          If InOverTime Then PRate = PRate * 1.5
          WrkTime = ConvertTimeToMinutes(Total_Time)
          Cal_LaborDollars Tb(1) & "", PRate, WrkTime, PaidHourly, 1, OvertimeHrs(EmpTwt, WrkTime)
          TWRKTIME = TWRKTIME + Val(WrkTime)
          EmpTwt = EmpTwt + Val(WrkTime)
       Tb.MoveNext
      Loop
    End If
  Id.MoveNext
 Loop
DbClose EDb, Id
DbClose Db, Tb
Ttl_L = 0
ConnectADODB SvrPath.TerminalPath & ClosingHistory, CHDB
  Sql = "delete from [Labor] where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'"
CHDB.Execute (Sql)
  Sql = "Select * from [Labor] where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'"
ConnectADORS (Sql), CHDB, CHRS
 For Pl = 0 To UBound(Dept_Pay, 1) - 1
    If Trim(Dept_Pay(Pl).DeptName) > "" And Val(Dept_Pay(Pl).DeptDollars) > 0 And TotalNetSales > 0 Then
     If Check_Option(91) = 1 Then
      LabPer = Val(Dept_Pay(Pl).DeptDollars) / (TotalNetSales + NetDiscounts) * 100
     Else
      LabPer = Val(Dept_Pay(Pl).DeptDollars) / TotalNetSales * 100
     End If
    Else
      LabPer = 0
    End If
    CHRS.AddNew
    CHRS![Date] = ClosingDate
    CHRS![Name] = Dept_Pay(Pl).DeptName & ""
    CHRS![Amount] = Val(Dept_Pay(Pl).DeptDollars)
    CHRS![Percentage] = Val(LabPer)
    CHRS.Update
    Ttl_L = Ttl_L + Val(Dept_Pay(Pl).DeptDollars)
 Next Pl
 If Ttl_L > 0 And TotalNetSales > 0 Then
   If Check_Option(91) = 1 Then
     Ttl_P = Ttl_L / (TotalNetSales + NetDiscounts) * 100
   Else
     Ttl_P = Ttl_L / TotalNetSales * 100
   End If
 End If
 SetNothing CHRS
 Sql = "Select * from [Sales] where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'"
 ConnectADORS Sql, CHDB, CHRS
   If CHRS.EOF Then CHRS.AddNew
   CHRS![Labor Dollars] = Val(Ttl_L)
   CHRS![Labor Percentage] = Val(Ttl_P)
   CHRS![Labor Hours] = Val(TWRKTIME)
   CHRS.Update
   DbClose CHDB, CHRS
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Calculate Labor Closing Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub Cal_LaborDollars(Dept_Worked As String, Dept_PayRate As Currency, Time_Worked As Single, PaidHourly As Integer, DaysWorked As Integer, OverTime As Single)
    On Error Resume Next
    Dim Cld        As Integer
    Dim SalaryPaid As Currency
    Dim DayRate    As Currency
    
    For Cld = 0 To UBound(Dept_Pay, 1) - 1
       If Trim$(Dept_Worked) = Trim$(Dept_Pay(Cld).DeptName) Then
        If PaidHourly = 1 Then  '// h
            Dept_Pay(Cld).DeptDollars = Dept_Pay(Cld).DeptDollars + (Dept_PayRate * Time_Worked) + CalculateOT(CDbl(OverTime), CCur(Dept_PayRate))
        Else
            If Dept_PayRate > 0 And DaysWorked > 0 Then
                DayRate = Dept_PayRate / 7
                SalaryPaid = DayRate * DaysWorked
                Dept_Pay(Cld).DeptDollars = Dept_Pay(Cld).DeptDollars + SalaryPaid
            End If
        End If
        Exit For
       End If
    Next Cld
End Sub
Private Function CashDeposits() As Currency
On Error Resume Next
Dim fdb As ADODB.Connection
Dim FRs As ADODB.Recordset
    CashDeposits = 0
    ConnectADODB SvrPath.TerminalPath & FinancialDb, fdb
    Sql = "Select distinct sum([Payment Amount]) as PayMentTotal from [DayPayments]  where [check number] = -9999 and [Payment Type] = 'Cash'"
    ConnectADORS Sql, fdb, FRs
    If Not FRs.EOF Then
       CashDeposits = Val(FRs![PaymentTotal] & "")
    End If
    DbClose fdb, FRs
End Function
Private Sub Form_Load()
On Error GoTo ErrorHandler
ResizeScreen1024 Me
RemoveOpenCheck True
UpdateOnLineTip
ResetEODEmail
Set JobTipOut = New FinancialItems
Dim Db  As ADODB.Connection
Dim Tb  As ADODB.Recordset
Dim LPO As Integer

CCards.cols = 4
CCards.HighLight = flexHighlightNever
CCards.rows = 100
CCards.ColAlignment(1) = 0
CCards.ColWidth(0) = 1675 * WidScale1024
CCards.ColWidth(1) = 1200 * WidScale1024
CCards.ColWidth(2) = 0
CCards.ColWidth(3) = 0
With POuts
  .cols = 3
  .HighLight = flexHighlightNever
  .rows = 500
  .ColAlignment(1) = 0
  .ColWidth(0) = 1675 * WidScale1024
  .ColWidth(1) = 1100 * WidScale1024
  .ColWidth(2) = 0
 Dim Pdouts As SortAll
 Set Pdouts = New SortAll
 TotalPO = GetPaidOuts("", Pdouts, 0, False, MyRegisterName, "", 0)
  .TextMatrix(0, 0) = "Open Paidout"
  .TextMatrix(0, 1) = Format(0, "0.00")
  .TextMatrix(0, 2) = "Open Paidout"
  If TotalPO > 0 Then
   For LPO = 1 To Pdouts.count
    .TextMatrix(LPO, 0) = Pdouts.Item(LPO).MenuItem
    .TextMatrix(LPO, 1) = Format(Pdouts.Item(LPO).SaleAmt, "0.00")
    .TextMatrix(LPO, 2) = Pdouts.Item(LPO).MenuItem
    Next LPO
    Totals(4).Caption = Format(TotalPO, "0.00")
  End If
 Set Pdouts = Nothing
End With '// finish loading paid outs

DoEvents  ''/ not sure why this is here

Sales_Breakdown False, 0

If GrossSales = 0 Then
  See_Msg "No Sales for this day", 5, "Please check the day"
End If

SvrTipsCollect = GetTipShare(JobTipOut)

If GetConfigInfo("Misc", "Tip Out Is Reference", "Yes", "Functions.INI") = "ENABLED" Then
 SvrTipsCollect = 0
End If
If SvrTipsCollect > 0 Then
  Titles(27).Caption = "Non Sales Revenue " & vbCrLf & "(includes tip outs)"
Else
  Titles(27).Caption = "Non Sales Revenue"
End If

Sales_data(0).Caption = Align(1, GrossSales)
Sales_data(1).Caption = Align(1, VoidSales) '+ Val(S_Data(0).Text)
Sales_data(2).Caption = Align(1, NetSTax)
Sales_data(3).Caption = Align(1, NetSpTax)
Sales_data(4).Caption = Align(1, NetLiqTax)
Sales_data(5).Caption = Align(1, NetDiscounts)
Sales_data(9).Caption = Align(1, TotalNetSales)
Sales_data(10).Caption = Align(1, ExemptSale + SvrTipsCollect)

NetGratuity = GratuityCharged("Day", "", 0)


Set ListRefunds = New FinancialItems
ListAllReturns ListRefunds, 0
TotalRefunds = GetAllReturns(True, NetRefund).RefundAmount

Totals(0).Caption = Format(VoidSales, "0.00")
Totals(1).Caption = Format(NetDiscounts, "0.00")

TotalTip = 0: CashTenders = 0
    Dim Payments As FinancialItems
    Set Payments = New FinancialItems
    'SortCreditCards Payments, 0, "", 0, "", ""
    SortEndOfDayPayments Payments
    Totals(2).Caption = ""
    Dim RN As Integer
    Dim W  As Integer
    CheckDeposit = 0
    RN = -1
        With CCards
        '.TextMatrix(0, 0) = "Non Cash"
        '.TextMatrix(1, 0) = "Gift Cards"
        For W = 1 To Payments.count
            'If Payments.Item(W).Tender = GetConfigInfo("Credit Cards", "Include Tender In Deposit (Tender Name)", "", "Functions.INI") Then
            '  CheckDeposit = Format(Payments.Item(W).Amount, "0.00")
            If Payments.Item(W).Tender <> "Cash" Then
                RN = RN + 1
                .TextMatrix(RN, 0) = Payments.Item(W).Tender
                .TextMatrix(RN, 1) = Format(Payments.Item(W).Amount, "0.00")
                .TextMatrix(RN, 2) = Payments.Item(W).ItemQty
                .TextMatrix(RN, 3) = Payments.Item(W).TipAmount
                 Totals(2).Caption = Format(Val(Totals(2).Caption & "") + Payments.Item(W).Amount, "0.00")
            End If
        Next W
        End With
        
ChargeBack = 0
For W = 1 To Payments.count
   ChargeBack = ChargeBack + CalChargeBack(Payments.Item(W).Tender, Payments.Item(W).TipAmount)
   TotalTip = TotalTip + Payments.Item(W).TipAmount
Next W

If Check_Option(237) = 1 Then
  CashtoAcctFor = CashtoAcctFor + TotalTip
  Sales_data(12).Caption = Format(TotalTip, "0.00")
  TipWith.Visible = True
Else
  TipWith.Visible = False
  Sales_data(12).Visible = False
End If

Set Payments = Nothing

Dim NonCashPayments As Currency
Dim NonCashDeposits As Currency

CashDep = CashDeposits()
CreditCardDep = CreditCardDeposits("")

Totals(2).Caption = Format(Val(Totals(2).Caption) - CreditCardDep, "0.00")

NonCashDeposits = DepositsCollected - CashDep
TDepositCollected = DepositsCollected

CashtoAcctFor = (GrossSales + (Abs(NetSTax) + (Abs(NetSpTax)) + (Abs(NetLiqTax))))

If Input_Option(58) = "Yes" Then
  CashtoAcctFor = CashtoAcctFor + NetGratuity
Else
  CashtoAcctFor = CashtoAcctFor - NetGratuity
End If

CashtoAcctFor = CashtoAcctFor + (((CashDep + ChargeBack) + SvrTipsCollect) - (AmtOfRefunds))

Sales_data(6).Caption = Format(CashtoAcctFor - (((Val(Totals(2).Caption & "") + Abs(VoidSales)) + Abs(NetDiscounts)) + Abs(Totals(4).Caption)), "0.00")
Sales_data(7).Caption = Format(CashDep, "0.00")
Sales_data(8).Caption = Format(NonCashDeposits, "0.00")
Sales_data(11).Caption = Format(TotalRefunds, "0.00")

Totals(6).Caption = Format(NonCashDeposits, "0.00")

If CheckDeposit > 0 Then
   S_Data(3).Text = Format(CheckDeposit, "0.00")
   Titles(11).Caption = "Checks"
End If

Titles(15).Caption = "Ending Safe Amount"

If GetBankDrops = 0 Then
   Titles(10).Caption = "Cash From Audits"
   S_Data(2).Text = Format(GetBankDeposits, "0.00")
ElseIf GetBankDrops > 0 Then
   Titles(10).Caption = "Day Deposits"
Else
  Titles(10).Caption = "Bank Deposit 1"
End If
S_Data(7).Text = Format(ReturnChangeFundAdj(GetBusinessDay), "0.00")

Figure_Shortage
Set TTRF = Nothing
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Closing Report Load ")
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
Set ListRefunds = Nothing
Set JobTipOut = Nothing
Set Void = Nothing
EndClosing
End Sub
Private Function GetBankDrops() As Currency
On Error GoTo ErrorHandler
GetBankDrops = 0
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & FinancialDb, Db
      Sql = "Select * from [Bank Deposits]"
    ConnectADORS Sql, Db, RS
     If Not RS.EOF Then
      S_Data(2).Text = Format(RS.fields("Bank 1").Value, "0.00")
      S_Data(3).Text = Format(RS.fields("Bank 2").Value, "0.00")
      S_Data(4).Text = Format(RS.fields("Bank 3").Value, "0.00")
      S_Data(5).Text = Format(RS.fields("Bank 4").Value, "0.00")
      GetBankDrops = ((RS.fields("Bank 1").Value + RS.fields("Bank 2").Value) + RS.fields("Bank 3").Value) + RS.fields("Bank 4").Value
     Else
      S_Data(2).Text = "0.00"
      S_Data(3).Text = "0.00"
      S_Data(4).Text = "0.00"
      S_Data(5).Text = "0.00"
     End If
 DbClose Db, RS
 Exit Function
ErrorHandler:
DbClose Db, RS
End Function

Private Sub HSFTP_SSHServerAuthentication(HostKey As String, FingerPrint As String, KeyAlgorithm As String, Accept As Boolean)
Accept = True
End Sub

Private Sub POuts_Click()
On Error GoTo ERH
With POuts
If Misc_Price > 0 Then
 If Trim$(.TextMatrix(.Row, 0)) = "" Then
  Misc_Price = 0
  Exit Sub
 End If
 .TextMatrix(.Row, 1) = Format(Misc_Price, "0.00")
 Misc_Price = 0
End If
If EraseOk = True Then
 SSCommand5.Caption = "Erase False"
 .TextMatrix(.Row, 1) = ""
 EraseOk = False
End If
Totals(4).Caption = "0.00"
 For TP = 0 To POuts.rows - 1
  If .TextMatrix(TP, 0) = "" Then Exit For
  If Val(.TextMatrix(TP, 1)) > 0 Then
   Totals(4).Caption = Format(Val(Totals(4).Caption) + Val(.TextMatrix(TP, 1)), "0.00")
  End If
 Next TP
End With
Figure_Shortage
Exit Sub
ERH:
See_Msg "Error", 5, Err.Description
End Sub
Private Sub POuts_KeyPress(KeyAscii As Integer)
With POuts
 If Trim$(.TextMatrix(.Row, 0)) = "" Then
  Exit Sub
 End If
If KeyAscii = 8 And Len(.TextMatrix(.Row, 1)) > 0 Then
  .TextMatrix(.Row, 1) = Mid(.TextMatrix(.Row, 1), 1, Len(.TextMatrix(.Row, 1)) - 1)
End If
If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46 Or KeyAscii = 45 Then
 .TextMatrix(.Row, 1) = .TextMatrix(.Row, 1) & Chr(KeyAscii)
 KeyAscii = 0
End If
Totals(4).Caption = "0.00"
 For TP = 0 To .rows - 1
  If .TextMatrix(TP, 0) = "" Then Exit For
  Totals(4).Caption = Format(Val(Totals(4).Caption) + Val(.TextMatrix(TP, 1)), "0.00")
 Next TP
End With
Figure_Shortage
End Sub


Private Sub S_Data_Change(Index As Integer)
Select Case Index
   Case Is = 0  'void Total
        Totals(0).Caption = Format(VoidSales, "0.00")
   Case Is = 1  'Discount total
        Totals(1).Caption = Format(NetDiscounts, "0.00")
End Select
AcctFor = 0
'S_Data(0) and S_Data(1) ''no longer used manual voids and discounts
For W = 2 To 7
  AcctFor = AcctFor + Val(S_Data(W).Text)
Next W
Totals(3).Caption = Format(Val(AcctFor), "0.00")
Figure_Shortage
End Sub
Private Sub S_Data_Click(Index As Integer)
If Misc_Price > 0 Then
  S_Data(Index).Text = Format(Misc_Price, "0.00")
  Misc_Price = 0
  S_Data_Change Index
  Figure_Shortage
End If
If EraseOk = True Then
SSCommand5.Caption = "Erase False"
S_Data(Index).Text = ""
EraseOk = False
End If
End Sub
Private Sub S_Data_DblClick(Index As Integer)
S_Data(Index).Text = Format(Val(ShowNumberPad("Enter Amount", False, 285, 0)), "0.00")
Figure_Shortage
End Sub
Private Sub S_Data_GotFocus(Index As Integer)
Highlighton S_Data(Index), "On"
End Sub
Private Sub S_Data_KeyPress(Index As Integer, KeyAscii As Integer)
If Clear_Text = 0 Then
    S_Data(Index).Text = ""
    Clear_Text = 1
End If
If KeyAscii = 13 Then Exit Sub
If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46 Or KeyAscii = 45 Or KeyAscii = 8 Then
Select Case Index
   Case Is = 0  'void Total
        Totals(0).Caption = Format(VoidSales, "0.00")
   Case Is = 1  'Discount total
        Totals(1).Caption = Format(NetDiscounts, "0.00")
End Select
Figure_Shortage
Else
KeyAscii = 0
End If
End Sub
Private Sub S_Data_LostFocus(Index As Integer)
Clear_Text = 0
Highlighton S_Data(Index), "Off"
S_Data(Index).Text = Format(S_Data(Index).Text, "0.00")
Figure_Shortage
End Sub
Private Function FindDate() As String
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim FS As Integer
FindDate = ""
  Sql = "Select Distinct TRY_CAST([Order Date] AS datetime) from [Closed Checks] Where [Close Out Day] = 0"
  Sql = Sql & " Group by [Order Date] Order by TRY_CAST([Order Date] AS datetime) asc"
  ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
  ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
     If AutoCloseInProcess Then  '// make sure with auto close that only use the oldest date found
       FindDate = Format(RS(0), "MM/DD/YYYY")
     Else
      If RecordsCount(RS) > 1 Then
        If See_Msg("Is This The Correct Close Out Date ? " & Format(RS(0), "MM/DD/YYYY"), 4, "If not select no for date selection") = 6 Then
            FindDate = Format(RS(0), "mm/DD/YYYY")
        Else
            Load_What = "FindDate"
            Find_What = "Select Date"
            Select_What.Show 1
            FindDate = GetSelected
        End If
      Else
        FindDate = Format(RS(0), "MM/DD/YYYY")
      End If
     End If
     DbClose Db, RS
     Exit Function
    End If
  DbClose Db, RS
If Not IsDate(FindDate) Then FindDate = Format(Date, "MM/DD/YYYY")
Exit Function
ErrorHandler:
FindDate = Format(Date, "MM/DD/YYYY")
Err.Clear
End Function
Private Sub SSCommand1_Click()
On Error GoTo ErrorHandler
Dim IV As Integer
If Sales_data(0) = 0 Then
  See_Msg "Nothing to Close out", 5, "No Sales Found"
  Exit Sub
End If
If Val(S_Data(7).Text) = 0 Then
  If GetConfigInfo("Daily Close", "Safe Amount Required", "No", "Functions.INI") = "ENABLED" Then
     See_Msg "Safe Amount Required..", 5, ""
     Exit Sub
  End If
End If
Dim FileDate            As String
Dim DoReports(10)       As Integer
Dim ff                  As Integer
Dim InVar               As Integer
Dim BkUpFile            As String
Dim GetReportName       As String
Dim TXTFileInfo()       As Variant
Dim Db                  As ADODB.Connection
Dim Tb                  As ADODB.Recordset
Dim DOW                 As Integer
Dim Prec                As Integer
OverWriteSales = True
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "INI Files\" & "ClosingReport.ini"), TXTFileInfo()) Then
    For IV = 0 To UBound(TXTFileInfo, 1)
      DoReports(IV) = TXTFileInfo(IV)
    Next IV
Else
    For IV = 0 To 10
      DoReports(IV) = 0
    Next IV
End If
If ClosingProgess Then
 If See_Msg("Closing In Progress", 4, "Start Close Again") <> 6 Then
   Exit Sub
 End If
End If
If See_Msg("Are You Sure ?", 4, "Close Out Will Reset all Sales Data for Next Day") <> 6 Then Exit Sub
 StartClosing  'mark records for close
 Sleep 0
 ClosingDay = True
TryNewDate:
 CloseDate = FindDate  'check for multi dates in the database
 If Not IsDate(CloseDate) Then
  See_Msg "Not A Valid Date", 5, "Try Again"
  GoTo TryNewDate
 End If
 CancelClose = False
 If CloseDateFound(CloseDate) = True Then
  CenterObject Me, CloseDateEdit
  DateEdit.Caption = "Date For Close Out Is " & CloseDate
  CloseDateEdit.Visible = True
  CloseDateEdit.ZOrder 0
  Do While CloseDateEdit.Visible = True
    DoEvents
    Sleep 250
  Loop
  S_Data(7).Text = Format(ReturnChangeFundAdj(CloseDate), "0.00")
 End If
 If CancelClose = True Then Exit Sub
 If Not IsDate(CloseDate) Then
  CloseDate = Format(Date, "MM/DD/YYYY") 'GetDate  'if returned date is nothing then get date from datebase
 End If
'Set Flag for history file updates
 SSCommand3.Enabled = False
 SSCommand1.Enabled = False
 If DoReports(0) = 0 And See_Msg("Print A Daily Sales Mix ?", 4, "") = 6 Then  'will not print sales Mix
   Sale_Mix Today, PrintSMix, "", "", "", "None", 0
 Else
   Sale_Mix Today, NoPrint, "", "", "", "None", 0
 End If
 PageHeader = "Daily Close For " & CloseDate
 CenterObject Me, Progression
 Progression.Visible = True
PAgain:
 'If GetConfigInfo("General", "Sync Sales Data", "OFF", "Functions.INI") = "ENABLED" Then
 ' MatchClosedChecks
 'End If
 Prec = 0
 Sales_Breakdown False, 0
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 Hourly_Sales_Data CloseDate
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 If DoReports(1) = 0 Then Print_Hourly_Sales False, CloseDate
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 If DoReports(4) = 0 Then Print_Meal_Period False
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 Dept_Report "", "Daily " & Format(Date, "M/D/YY"), DoReports(2), "Day", "", 0, False
 If DoReports(5) = 0 Then PrintServeTime False, CloseDate
 If DoReports(6) = 0 Then DeptVoid "", True, False
 Sales_Report 0, DoReports(7), CloseDate
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 Get_Voids
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
 Nightly_Sales DoReports(8), CloseDate
 PrintDailyInventory True
 DailySalesMix
 If SiteLic.Berg = 1 Then
  BergExceptions CloseDate
 End If
 If Check_Option(32) = 1 Then
   GetOrderHistory CStr(CDate(CloseDate) + 1), CStr(CDate(CloseDate) + 1)
 End If
 If GetConfigInfo("General", "Prompt Order Type", "Off", "Functions.INI") = "ENABLED" Then
  GetOrderTypes CStr(CloseDate), 0
 End If
 Prec = Prec + 6: Progress.ProgressBarValue = Prec: DoEvents
   If AreAuditsOpen("") Then
     See_Msg "Audits Open Cannot Reset Day !!", 5, ""
     Trans_Log "Reset Abort - Audits Open"
     SSCommand3.Enabled = True
     SSCommand1.Enabled = True
     Exit Sub
   End If
  Trans_Log "Daily Reset Started"
  Progression.Visible = False
  DoEvents
  DOW = 0
  SaveRefundsToHistory ListRefunds, CloseDate
  See_Msg "Saving Taxed Sales", 5, "Daily Reset"
  StoreTaxedSales CloseDate, OverWriteSales
  StoreNoSale CloseDate, OverWriteSales
  See_Msg "Saving Day History", 5, "Daily Reset"
  Sleep 0
  StoreDayHistory OverWriteSales, CloseDate
  See_Msg "Saving Department Sales", 5, "Daily Reset"
  Sleep 0
  SaveDepartment CloseDate, OverWriteSales
  See_Msg "Saving Labor", 5, "Daily Reset"
  Sleep 0
  Calculate_Labor CloseDate
  See_Msg "Saving Product History", 5, "Daily Reset"
  Sleep 0
  StoreHeader CloseDate, OverWriteSales
  See_Msg "Saving Refunds", 5, "Daily Reset"
  Sleep 0
  StoreRefunds CloseDate, OverWriteSales
  See_Msg "Saving Payments", 5, "Daily Reset"
  Sleep 0
  CollectWithHeldTips CloseDate
  CollectTipOutWithHeld CloseDate, JobTipOut
  Progression.Visible = True
  CloseWeek = ResetEndOfDay(CloseDate)
  Progression.Visible = False
  SendPayroll
  On Error GoTo ErrorHandler
   Progression.Visible = False
   'Call MemoryCheck(Red_Book)
   If Check_Option(61) = 1 Then
     ExportFinancialData
   End If
ClosingDay = False
SSCommand3.Enabled = True
SSCommand1.Enabled = True
If SendEmails(5).Send Then
 See_Msg "Building Balance Sheet", 5, ""
 Display.DSRBalanceSheet CStr(CDate(CloseDate)), CStr(CDate(CloseDate)), SendEmails(5).Send, DoReports(9)
End If
If SendEmails(4).Send Then
 See_Msg "Building Daily Audit Report", 5, ""
 Display.DailySummary CStr(CDate(CloseDate)), SendEmails(4).Send, DoReports(9)
End If
Call MemoryCheck(Display)
EndClosing
See_Msg "Close Out Complete....", 5, "System Reset for Next Day"
Trans_Log "Daily Closing Completed..."
Trans_Log "Batch Process Started"
CloseBatch False
Trans_Log "Daily Batch Complete"
SendEMailEOD
EPServiceIsUp
' If FileCheck(SvrPath.TerminalPath & "INI Files\" & "ExChange.INI") Then
'  ExChange.ExRate = wConvertCurrency(ExChange.ExName)
'  If ExChange.ExRate > 0 Then
'   ff = FreeFile
'   Open SvrPath.TerminalPath & "INI Files\" & "Exchange.INI" For Output As ff
'   Write #ff, Trim$(ExChange.ExName), ExChange.ExRate
'   Close
'  End If
' End If
 If CloseWeek Or Format(CloseDate, "dddd") = Input_Option(60) Then
  See_Msg "Closing Week Sales ", 5, "Resetting Weekly Sales Data"
  Trans_Log "Weekly Closing "
  Transaction.WeekClose.DoClick
  CloseWeek = False
 End If
 If Val(GetSetting("ProductID", "License File", "General Data 24")) <> 0 Then
  Trans_Log "Export to Hot Schedules " & HotScheduleExport
 End If
 If Val(GetSetting("Micro$ale", "Register Setup", "Ctuit")) <> 0 Then
   Trans_Log "Export to Compeat " & SendCompeat(CloseDate)
 End If
 If Val(GetSetting("Micro$ale", "Register Setup", "QuickBooks")) <> 0 Then
  Trans_Log "Export to quick books " & SendQBexport(CloseDate, CloseDate)
 End If
UnSuspendAllOpenChecks
Call MemoryCheck(Closing_Report)
Exit Sub
ErrorHandler:
If Err.Number = 3365 Then
  Trans_Log "Daily Close " & Sql
End If
SSCommand3.Enabled = True
SSCommand1.Enabled = True
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Closing Report SS1 ")
   Err.Clear
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub Reset_HouseCharges(DOW As Integer, ClosingDate As String)
On Error GoTo ErrorHandler
    Dim R As String
    Dim GRS As GeneralLib.Utilities
    Set GRS = New GeneralLib.Utilities
    Sql = "Update [HouseCharges]"
    Sql = Sql & " Set [Day of Charge] = " & DatePart("d", CDate(ClosingDate), WeekStartsOn, vbFirstJan1)
    Sql = Sql & " where [Day of Charge] = 0"
    If Check_Option(165) = 1 Then
      Sql = Sql & " and [Location Name] = '" & LocationReset & "'"
    End If
    R = GRS.ExecuteSql(Sql, ReturnAdoConnection(HouseAcct))
    Set GRS = Nothing
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Closing House Charges ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub BergExceptions(ClosingDate As String)
On Error GoTo ErrorHandler
If SiteLic.Berg = 0 Then Exit Sub
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim FR          As Integer
Dim BergEx()    As Variant
Dim TEX         As Integer
Dim RN          As Integer
Dim PrinterPort As String
Dim ItemPrice   As Currency
Dim PF          As Integer
Dim LRF         As Integer
Erase BergEx()
ReDim BergEx(4, 0)
TEX = 0: RN = 0
    OpenDb SvrPath.TerminalPath & FinancialDb, Db
     Sql = "Select * from [Pours] Where [Open] = 1 and [Check No] = 0"
    ConnectADORS Sql, Db, RS
     TEX = RecordsCount(RS)
     ReDim BergEx(4, TEX)
     Do While Not RS.EOF
      RN = RN + 1
      BergEx(1, RN) = RS![Register Name] & ""
      BergEx(0, RN) = RS![Time] & ""
      BergEx(2, RN) = Left(PLUtoMenu(RS![Plu No] & "", ItemPrice), 15)
      BergEx(3, RN) = Format(ItemPrice, "0.00")
      RS.MoveNext
    Loop
  DbClose Db, RS
If UBound(BergEx, 2) = 0 Then Exit Sub

If Check_Option(96) = 0 Then
   Printer_Drivers "Receipt"
   PF = FreeFile
   PrinterPort = "Berg Ex"
tryagain:
    Open PrinterPort For Output Shared As #PF
     CenterText = EC(8)
     LeftText = EC(9)
     'RightText = Chr(27) + Chr(97) + Chr(2)
    Print #PF, EC(1); "Drink Not Accounted For"; EC(4); EC(7)
    Print #PF, String(40, 45); EC(7)
    Print #PF, "Time"; Tab(12); "Register"; Tab(28); "PLU NO"
    Print #PF, String(40, 45); EC(7)
    For LRF = 1 To UBound(BergEx(), 2)
      Print #PF, BergEx(0, LRF); Tab(12); BergEx(1, LRF); Tab(31); BergEx(2, LRF) & EC(7)
    Next LRF
    Print #PF, String(40, 45); EC(7)
    Print #PF, EC(6); EC(6); EC(5)
    Close #PF
    SendJobToPrinter PrinterPort, "Receipt", False
Else
   SetRegular 0
   SetPage ReportView.View, "Liquor Exception Report " & Format(ClosingDate, "MM/DD/YYYY")
   DrawLine
   PrintLine "Time", 2500
   PrintLine "Register", 4500
   PrintLine "PLU NO", 6000
   PrintLine "Price", -1
   DrawLine
   For LRF = 1 To UBound(BergEx(), 2)
     PrintLine BergEx(0, LRF), 2500
     PrintLine BergEx(1, LRF), 4500
     PrintLine BergEx(2, LRF), 6000
     PrintLine BergEx(3, LRF), -1
   Next LRF
   DrawLine
   ReportView.View.EndDoc
   ReportView.View.PrintDoc
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Berg Exceptions ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub SSCommand3_Click()

If ClosingProgess Then
 If See_Msg("Wait for batch to complete", 4, "No to exit") = 6 Then
  Exit Sub
 End If
End If
'Call MemoryCheck(Red_Book)
Call MemoryCheck(Closing_Report)
End Sub
Private Sub SSCommand4_Click()
Dim DrwCt As AllCounts
Set DrwCt = New AllCounts
Dim MYC As GeneralLib.Utilities
Set MYC = New GeneralLib.Utilities
Misc_Price = MYC.GetCalculator(DrwCt)
Set MYC = Nothing
Set DrwCt = Nothing
End Sub
Private Sub SSCommand5_Click()
If EraseOk = False Then
SSCommand5.Caption = "Erase Yes"
EraseOk = True
Exit Sub
Else
SSCommand5.Caption = "Erase No"
EraseOk = False
End If
End Sub
Public Sub Count_OverShort()
Totals(5).Caption = ""
AcctFor = 0
For Ts = 0 To 4
AcctFor = AcctFor + Val(Totals(Ts).Caption)
Next Ts
AcctFor = AcctFor = Val(Totals(6).Caption)
Totals(5).Caption = Format(Val(AcctFor - CashtoAcctFor), "0.00")
If Val(Totals(5).Caption) < 0 Then
Totals(5).ForeColor = vbRed
Else
Totals(5).ForeColor = vbBlack
End If
End Sub
Public Sub Nightly_Sales(PrintReport As Integer, ClosingDate As String)
On Error GoTo ErrorHandler
Dim OverShort       As Currency
Dim PrinterPort     As String
Dim Db              As ADODB.Connection
Dim Tb              As ADODB.Recordset
Dim RS              As ADODB.Recordset
Dim RN              As Integer
Dim Reg             As Integer
Dim DCT             As Currency
Dim DCQ             As Integer
Dim DP              As Integer
Dim CCardTotal      As Currency
Dim LRF             As Integer
Dim BankDeposit     As Currency
Dim CashCollected   As Currency
Dim FGT             As Integer
Dim QtyCb           As Integer
Dim WtdNotTaxed     As Currency
Dim TaxedSales      As Currency
Dim PC              As Integer
Dim DailyFee        As Currency
Dim ChargeTip       As Currency
Dim TotalCharges    As Currency

Dim RLIST           As FinancialItems
Set RLIST = New FinancialItems

Dim DiscountTotals  As StoreDiscounts
Set DiscountTotals = New StoreDiscounts

Dim DeptTotals      As StoreDiscounts
Set DeptTotals = New StoreDiscounts

Trans_Log "Load Coupons"

List_Coupons DiscountTotals, DeptTotals, True, 1, ClosingDate

BankDeposit = 0
  For DP = 2 To 6
   BankDeposit = BankDeposit + Val(S_Data(DP).Text & "")
  Next DP
'Trans_Log "Load Refunds"   '1

If PrintReport = 0 Then
   If Check_Option(96) = 0 Then
    Printer_Drivers "Receipt"
    PF = FreeFile
    PrinterPort = "Nightly Sales"
tryagain:
   Close
   Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    If ListRefunds.count > 0 Then
        Print #PF, EC(1); "Refund List"; EC(4); EC(7)
        Print #PF, "Time"; Tab(16); "Manager"; Tab(31); "Amount"
        Print #PF, String(40, 45); EC(7)
        For LRF = 1 To ListRefunds.count
          Print #PF, ListRefunds.Item(LRF).TimeEnter; Tab(15); ListRefunds.Item(LRF).Manager; Tab(35); Format(ListRefunds.Item(LRF).Amount, "0.00")
          Print #PF, "Refund Reason "; ListRefunds.Item(LRF).Reason
        Next LRF
        Print #PF, String(40, 45); EC(7)
     End If
     If BankDeposit > 0 Then
        Print #PF, EC(1); EC(9); "Bank Deposits"; EC(4); EC(7)
        Print #PF, String(40, 45); EC(7)
     End If
    Else
      StartPage
      If ListRefunds.count > 0 Then
       DrawLine
       PrintLine "Refund List", -1
       PrintLine "Time", 2500: PrintLine "Manager", 4500: PrintLine "Amount", 6500: PrintLine "Reason", -1
       DrawLine
       For LRF = 1 To ListRefunds.count
         PrintLine ListRefunds.Item(LRF).TimeEnter, 2500: PrintLine ListRefunds.Item(LRF).Manager, 4500: PrintLine Format(ListRefunds.Item(LRF).Amount, "0.00"), 6500
         PrintLine ListRefunds.Item(LRF).Reason, -1
       Next LRF
       DrawLine
      End If
      If BankDeposit > 0 Then
        PrintLine "Bank Deposits", -1
        DrawLine
      End If
    End If
End If

If PrintReport = 0 Then
   For DP = 2 To 6
     If Check_Option(96) = 0 Then
       If Val(S_Data(DP).Text) > 0 Then Print #PF, "Bank Deposit            : "; Align(7, Val(S_Data(DP))); EC(7)
     Else
       If Val(S_Data(DP).Text) > 0 Then PrintLine "Bank Deposit ", 4500: PrintLine Format(Val(S_Data(DP)), CurrencyPreFix), -1
     End If
   Next DP
   If Check_Option(96) = 0 Then
     If Val(S_Data(7).Text) <> 0 Then Print #PF, "Safe Cash               : "; Align(7, Val(S_Data(DP))); EC(7)
   Else
     If Val(S_Data(7).Text) <> 0 Then PrintLine "Safe Cash            ", 4500: PrintLine Format(Val(S_Data(DP)), CurrencyPreFix), -1
   End If
   If BankDeposit > 0 Then
     If Check_Option(96) = 0 Then
       Print #PF, String(40, 45); EC(7)
       Print #PF, "Total Bank Deposits     : "; Align(7, BankDeposit); EC(7)
     Else
       DrawLine
       PrintLine "Total Bank Deposits", 4500
       PrintLine Format(BankDeposit, CurrencyPreFix), -1
     End If
   End If
End If
DailyFee = 0

If ProcessingFeesOn Then
   DailyFee = DailyProcessFee
End If
If PrintReport = 0 Then
  If TDepositCollected > 0 Then
    If Check_Option(96) = 0 Then
      Print #PF, String(40, 45); EC(7)
      Print #PF, "Deposits Collected"; EC(7)
      Print #PF, String(40, 45); EC(7)
      Print #PF, "Total Cash In Deposits  : "; Align(7, CashDeposits); EC(7)
      Print #PF, "Total NonCash Deposits  : "; Align(7, TDepositCollected - CashDeposits); EC(7)
      Print #PF, "Total Deposits          : "; Align(7, TDepositCollected); EC(7)
    Else
      DrawLine
      PrintLine "Deposits Collected", -1
      DrawLine
      PrintLine "Total Cash Deposits   ", 4500
      PrintLine Format(CashDeposits, "0.00"), -1
      PrintLine "Total Non Cash Deposits   ", 4500
      PrintLine Format(TDepositCollected - CashDeposits, "0.00"), -1
      PrintLine "Total Deposits   ", 4500
      PrintLine Format(TDepositCollected, "0.00"), -1
    End If
  End If
End If

TotalTip = 0
CashCollected = CashDue(TotalTip)

If PrintReport = 0 Then
   If Check_Option(96) = 0 Then
      Print #PF, String(40, 45); EC(7)
      Print #PF, EC(1); "Tender Types"; EC(4); EC(7)
      Print #PF, String(40, 45); EC(7)
   Else
      PrintLine "", -1
      PrintLine "Tender Types", -1
      DrawLine
   End If
   If Check_Option(96) = 0 Then   'added for jerry NY 12/05/2001
      If 24 - Len("Cash Collected") > 0 Then
        Print #PF, Trim("Cash Collected"); String(24 - Len("Cash Collected"), 32); ": ";
      Else
        Print #PF, Trim("Cash Collected");
      End If
      Print #PF, Align(7, CashCollected); EC(7)
   Else
      PrintLine ("Cash Collected"), 4500
      PrintLine Format(CashCollected, CurrencyPreFix), -1
   End If
   CCardTotal = 0
   With CCards
      For PC = 0 To .rows - 1
         If .TextMatrix(PC, 0) = "" Then Exit For
         Select Case .TextMatrix(PC, 0)
             Case Is = "Visa", "MasterCard", "Amex", "Discover", "Diners", "Master Card"
                  CCardTotal = CCardTotal + Val(.TextMatrix(PC, 1))
         End Select
         If Check_Option(96) = 0 Then
            If 24 - Len(.TextMatrix(PC, 0)) > 0 Then
               Print #PF, .TextMatrix(PC, 0); String(24 - Len(.TextMatrix(PC, 0)), 32); ": ";
            Else
               Print #PF, .TextMatrix(PC, 0);
            End If
            Print #PF, Align(7, Val(.TextMatrix(PC, 1))); EC(7)
         Else
            PrintLine (.TextMatrix(PC, 0)), 4500
            PrintLine Format(Val(.TextMatrix(PC, 1)), CurrencyPreFix), -1
         End If
      Next PC
   End With

 Trans_Log "Load Refunds"
      
 If Check_Option(96) = 0 Then
   Print #PF, String(40, 45); EC(7)
   Print #PF, "Total Credit Cards      : "; Align(7, CCardTotal); EC(7)
   If DailyFee > 0 Then
     Print #PF, "Total Processing Fees   : "; Align(7, DailyFee); EC(7)
     Print #PF, "Net CreditCard Deposit  : "; Align(7, CCardTotal - DailyFee); EC(7)
   End If
   Print #PF, "Total Non Cash Tenders  : "; Align(7, Val(Totals(2).Caption)); EC(7)
   If TotalTip > 0 Then Print #PF, "Total Tips Reported     : "; Align(7, TotalTip); EC(7)
   Print #PF, "Credit Card Returns     : "; Align(7, GetReturns(RLIST)); EC(7)
   Print #PF, "Returns Included In Total"; EC(7)
   Print #PF, String(40, 45); EC(7)
   Print #PF, "List Of Returns"; EC(7)
   Print #PF, String(40, 45); EC(7)
   For RN = 1 To RLIST.count
       Print #PF, RLIST.Item(RN).ItemName; Tab(15); RLIST.Item(RN).ItemName1; Tab(30); Align(7, RLIST.Item(RN).Amount); EC(7)
   Next RN
   Print #PF, String(40, 45); EC(7)
 Else
   DrawLine
   PrintLine "Total Credit Cards", 4500
   PrintLine Format(CCardTotal, CurrencyPreFix), -1
   If DailyFee > 0 Then
     PrintLine "Total Processing Fees", 4500
     PrintLine Format(DailyFee, CurrencyPreFix), -1
     PrintLine "Net CreditCard Deposit", 4500
     PrintLine Format(CCardTotal - DailyFee, CurrencyPreFix), -1
   End If
   PrintLine "Total Non Cash Tenders ", 4500
   PrintLine Format(Val(Totals(2)), CurrencyPreFix), -1
   If TotalTip > 0 Then
     PrintLine "Total Tips Reported ", 4500
     PrintLine Format(TotalTip, CurrencyPreFix), -1
   End If
   PrintLine "Total Credit Card Returns", 4500
   PrintLine Format(GetReturns(RLIST), CurrencyPreFix), -1
   PrintLine "Returns Included In Total", -1
   DrawLine
   PrintLine "List of Returns", -1
   DrawLine
   For RN = 1 To RLIST.count
      PrintLine RLIST.Item(RN).ItemName, 2500: PrintLine RLIST.Item(RN).ItemName1, 4500: PrintLine Format(RLIST.Item(RN).Amount, "0.00"), -1
   Next RN
   DrawLine
 End If
End If
' Add to handler in house charges
If PrintReport = 0 Then
 If SvrTipsCollect > 0 Then   'And PrintReport = 0
  Trans_Log "Load Svr Tips Collected"
  ChargeTip = 0
  If Check_Option(96) = 0 Then
     Print #PF, String(40, 45); EC(7)
     Print #PF, EC(1); "Todays Tip Out Collected"; EC(4); EC(7)
     Print #PF, EC(1); "Job Code               Amount"; EC(4); EC(7)
     Print #PF, String(40, 45); EC(7)
  Else
     PrintLine "", -1
     PrintLine "Todays Tip Out Collected", -1
     DrawLine
     PrintLine "", -1
  End If
  For PC = 1 To JobTipOut.count   'was -1 turned off May be dropping last entry
    If Check_Option(96) = 0 Then
      If 20 - Len(JobTipOut.Item(PC).ItemName) > 0 Then
        Print #PF, JobTipOut.Item(PC).ItemName; String(20 - Len(JobTipOut.Item(PC).ItemName), 32);
      Else
        Print #PF, Left$(JobTipOut.Item(PC).ItemName, 20);
      End If
      Print #PF, Tab(26); Format(JobTipOut.Item(PC).Amount, "0.00"); EC(7)
    Else
      PrintLine JobTipOut.Item(PC).ItemName, 6500
      PrintLine Format(JobTipOut.Item(PC).Amount, CurrencyPreFix), -1
    End If
    ChargeTip = ChargeTip + JobTipOut.Item(PC).Amount
  Next PC
  If Check_Option(96) = 0 Then
     Print #PF, String(40, 45); EC(7)
     Print #PF, "Total Tips Collected    : "; Align(7, ChargeTip); EC(7)
  Else
     PrintLine "", -1
     PrintLine "Todays Total Tip Out Collected ", 4500
     PrintLine Format(Val(ChargeTip), CurrencyPreFix), -1
     DrawLine
  End If
 End If
End If

'***************************************
If PrintReport = 0 And System_config = FullService Then
   If Check_Option(96) = 0 Then
      Print #PF, String(40, 196); EC(7)
      Print #PF, EC(1); "Cash Sales /W Tax  " & Format(CashSales, CurrencyPreFix); EC(4); EC(7)
      Print #PF, String(40, 196); EC(7)
   Else
      PrintLine "", -1
      PrintLine "Cash Sales with Tax    " & Format(CashSales, CurrencyPreFix), -1
      DrawLine
   End If
End If
If PrintReport = 0 Then
  If Val(Totals(4).Caption) > 0 Then
   If Check_Option(96) = 0 Then
      Print #PF, String(40, 45); EC(7)
      Print #PF, EC(1); "Paid Outs"; EC(4); EC(7)
      Print #PF, String(40, 45); EC(7)
   Else
      PrintLine "", -1
      PrintLine "Paid Outs", -1
      DrawLine
   End If
   With POuts
     For PC = 0 To .rows - 1
      If .TextMatrix(PC, 0) = "" Then Exit For
      If Val(.TextMatrix(PC, 1)) > 0 Then
         If Check_Option(96) = 0 Then
          If 24 - Len(Trim(.TextMatrix(PC, 0))) > 0 Then 'makes sure number is alway > 0 else err 5
              Print #PF, Trim(.TextMatrix(PC, 0)); String(24 - Len(Trim(.TextMatrix(PC, 0))), 32); ": ";
          Else
              Print #PF, Trim(.TextMatrix(PC, 0));
          End If
          Print #PF, Align(7, Val(.TextMatrix(PC, 1))); EC(7)
         Else
          PrintLine .Text, 4500
          PrintLine Format(Val(.TextMatrix(PC, 1)), CurrencyPreFix), -1
         End If
      End If
     Next PC
   End With
   If Check_Option(96) = 0 Then
     Print #PF, String(40, 45); EC(7)
     Print #PF, "Total Paid Outs         : "; Align(7, Val(Totals(4).Caption)); EC(7)
   Else
     PrintLine "Total Paid Outs ", 4500
     PrintLine Format(Val(Totals(4).Caption), CurrencyPreFix), -1
   End If
  End If
End If

Trans_Log "Load Department Totals"

If PrintReport = 0 Then
  If DeptTotals.count > 0 Then
    Dim FDL As Integer
    For FDL = 1 To DeptTotals.count
        If Check_Option(96) = 0 Then
           Print #PF, String(40, 45); EC(7)
           Print #PF, EC(1); DeptTotals.Item(FDL).DiscountDept; EC(4); EC(7)
           Print #PF, "Name                     Disc       #"; EC(7)
           Print #PF, String(40, 45); EC(7)
        Else
           DrawLine
           PrintLine DeptTotals.Item(FDL).DiscountDept, 4500: PrintLine "Discount Amt", 6500
           PrintLine "Quantity", -1
           DrawLine
        End If
        For CP = 1 To DiscountTotals.count
            If DiscountTotals.Item(CP).DiscountDept = DeptTotals.Item(FDL).DiscountDept Then
               If Check_Option(96) = 0 Then
                  If Len(DiscountTotals.Item(CP).DiscountName) > 20 Then DiscountTotals.Item(CP).DiscountName = Left(DiscountTotals.Item(CP).DiscountName, 20)
                  Print #PF, Trim$(DiscountTotals.Item(CP).DiscountName); String(20 - Len(Trim(DiscountTotals.Item(CP).DiscountName)), 32); ":";
                  Print #PF, Tab(26); Align(4, DiscountTotals.Item(CP).DiscountAmt); Tab(36); Format(DiscountTotals.Item(CP).DiscountQty, "##"); EC(7)
               Else
                  PrintLine DiscountTotals.Item(CP).DiscountName, 4500
                  PrintLine Format(DiscountTotals.Item(CP).DiscountAmt, CurrencyPreFix), 6500: PrintLine Format(DiscountTotals.Item(CP).DiscountQty, "##"), -1
               End If
            End If
        Next CP
    Next FDL
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select * from [Have One On Me]"
    ConnectADORS Sql, Db, RS
    If Not RS.EOF Then
       If Check_Option(96) = 0 Then
          Print #PF, String(40, 45); EC(7)
          Print #PF, EC(1); "Attached Coupons"; EC(4); EC(7)
          Print #PF, "Coupon Name          Check No   Mgr"; EC(7)
          Print #PF, "Reason"; EC(7)
          Print #PF, String(40, 45); EC(7)
       Else
          PrintLine "Coupon Name", 4500
          PrintLine "Check No ", 6500: PrintLine "Manager", 8500: PrintLine "Reason", -1
       End If
    End If
    Do While Not RS.EOF
       If Check_Option(96) = 0 Then
          Print #PF, Trim(RS![coupon Name] & "");
          Print #PF, Tab(22); Val(RS![Check Number] & ""); Tab(30); RS![Manager Name] & ""; EC(7)
          Print #PF, RS![Reason] & ""; EC(7)
       Else
          PrintLine Trim(RS![coupon Name] & ""), 4500
          PrintLine Val(RS![Check Number] & ""), 6500: PrintLine RS![Manager Name] & "", 8500: PrintLine RS![Reason] & "", -1
       End If
     RS.MoveNext
    Loop
    DbClose Db, RS
    If Check_Option(96) = 0 Then
       Print #PF, String(40, 45); EC(7)
    Else
       DrawLine
    End If
    DCT = 0: DCQ = 0
    If Check_Option(96) = 0 Then
       For FDL = 1 To DeptTotals.count
           DCT = DCT + DeptTotals.Item(FDL).DiscountAmt
           DCQ = DCQ + DeptTotals.Item(FDL).DiscountQty
           If 20 - Len("Total " & DeptTotals.Item(FDL).DiscountDept) > 0 Then
             Print #PF, "Total " & DeptTotals.Item(FDL).DiscountDept; String(20 - Len("Total " & DeptTotals.Item(FDL).DiscountDept), 32); ":";
           Else
             Print #PF, "Total " & DeptTotals.Item(FDL).DiscountDept; ":";
           End If
           Print #PF, Align(4, DeptTotals.Item(FDL).DiscountAmt); EC(7)
       Next FDL
       Print #PF, String(40, 45); EC(7)
       Print #PF, "Total"; String(20 - Len("Total"), 32); ":";
       Print #PF, Align(4, DCT); EC(7)
    Else
       For FDL = 1 To DeptTotals.count
           DCT = DCT + DeptTotals.Item(FDL).DiscountAmt
           DCQ = DCQ + DeptTotals.Item(FDL).DiscountQty
           PrintLine "Total " & DeptTotals.Item(FDL).DiscountDept, 4500
           PrintLine Format(DeptTotals.Item(FDL).DiscountAmt, CurrencyPreFix), 6500
           PrintLine Format(DeptTotals.Item(FDL).DiscountQty, "###"), -1
       Next FDL
       DrawLine
       PrintLine "Total ", 4500
       PrintLine Format(DCT, CurrencyPreFix), 6500
       PrintLine Format(DCQ, "###"), -1
    End If
  End If
End If
'end discount section
If PrintReport = 0 Then
  If Void.count > 0 Then
   VSale = 0
   VDisc = 0
   If Check_Option(96) = 0 Then
      Print #PF, EC(1); "Closed Check Voids"; EC(4); EC(7)
      Print #PF, "Check Number    Void Sale   Void Disc"; EC(7)
      Print #PF, String(40, 45); EC(7)
   Else
      PrintLine "Closed Check Voids", 4500
      PrintLine "Void Sale", 6500: PrintLine "Void Disc", -1
      DrawLine
   End If
   For CP = 1 To Void.count
       VSale = VSale + Abs(Void.Item(CP).Amount)
       VDisc = VDisc + Abs(Void.Item(CP).VoidAmount)
       If Check_Option(96) = 0 Then
          Print #PF, "Chk # "; CStr(Void.Item(CP).CheckNo); String(9 - Len(Trim(Str(Void.Item(CP).CheckNo))), 32); "  ";
          Print #PF, Align(7, Void.Item(CP).Amount); Tab(34); Align(5, Void.Item(CP).VoidAmount); EC(7)
       Else
          PrintLine "Check No. " & Void(0, CP), 4500
          PrintLine Format(Void.Item(CP).Amount, CurrencyPreFix), 6500: PrintLine Format(Void.Item(CP).VoidAmount, "0.00"), -1
       End If
   Next CP
   If Check_Option(96) = 0 Then
      Print #PF, String(40, 45); EC(7)
      If VSale <> 0 Then Print #PF, "Total Void Sales         "; Align(7, VSale); EC(7)
      If VDisc <> 0 Then Print #PF, "Total Void Disc/Coupons  "; Align(7, VDisc); EC(7)
   Else
      DrawLine
      If VSale <> 0 Then PrintLine "Total Void Sales ", 4500: PrintLine Format(VSale, CurrencyPreFix), -1
      If VDisc <> 0 Then PrintLine "Total Void Disc/Coup ", 4500: PrintLine Format(VDisc, CurrencyPreFix), -1
   End If
  End If
End If

If PrintReport = 0 Then
  If Check_Option(96) = 0 Then
     Print #PF, String(40, 45); EC(7)
  Else
     DrawLine
  End If
End If

AcctFor = 0
For All = 2 To 4
  AcctFor = AcctFor + Val(Totals(All).Caption)
Next All
AcctFor = AcctFor + Val(Totals(6).Caption)

NetGratuity = GratuityCharged("Day", "", 0)

ADJ_SAles = Val(S_Data(0).Text) + Val(S_Data(1).Text)

Gross_Total = ((((((TotalNetSales + ExemptSale) + NetSTax) + NetSpTax) + NetLiqTax) + ChargeBack) + SvrTipsCollect) - (Val(S_Data(0).Text) + Val(S_Data(1).Text))

If Input_Option(58) = "Yes" Then

Else
 Gross_Total = Gross_Total - NetGratuity
End If
Gross_Total = Gross_Total + TDepositCollected

If PrintReport = 0 Then
   If Check_Option(96) = 0 Then
      Print #PF, EC(1); "Cash To Account For      "; Align(7, Val(Sales_data(6).Caption)); EC(4); EC(7)
      Print #PF, EC(1); "Adjusted Gross Total     "; Align(7, Val(Gross_Total)); EC(4); EC(7)
      If ChargeBack > 0 Then
         Print #PF, "(includes tip fee income ) "; Format(ChargeBack, "0.00"); EC(7)
      End If
      If SvrTipsCollect > 0 Then
         Print #PF, "(includes tip share) "; Format(SvrTipsCollect, "0.00"); EC(7)
      End If
      If Input_Option(58) = "Yes" Then
        Print #PF, "Gratuity Included"; EC(7)
      End If
   Else
      PrintLine "Cash to Account For      ", 4500
      PrintLine Format(Val(Sales_data(6).Caption), CurrencyPreFix), -1
      PrintLine "Adjusted Gross Total     ", 4500
      PrintLine Format(Val(Gross_Total), CurrencyPreFix), -1
      If ChargeBack > 0 Then
        PrintLine "(includes tip fee income of) " & Format(ChargeBack, "0.00"), -1
      End If
      If ChargeBack > 0 Then
        PrintLine "(includes tip share of) " & Format(ChargeBack, "0.00"), -1
      End If
      If Input_Option(58) = "Yes" Then
        PrintLine "Gratuity Included", -1
      End If
   End If
End If

If Val(ADJ_SAles) > 0 And PrintReport = 0 Then
   If Check_Option(96) = 0 Then
     Print #PF, "manual disc & voids are removed"
   Else
     PrintLine "manual disc & voids are removed", -1
   End If
End If

OverShort = AcctFor - Gross_Total

If PrintReport = 0 Then
   If Check_Option(96) = 0 Then
      Print #PF, EC(1); "Total Accounted For      "; Align(7, Val(AcctFor)); EC(4); EC(7)
      Print #PF, String(40, 45); EC(7)                                                       'Credit Card   Paidouts
   Else
      PrintLine "Total Accounted For", 4500: PrintLine Format(Val(AcctFor), CurrencyPreFix), -1
      DrawLine
   End If
End If
If PrintReport = 0 Then
   If Check_Option(96) = 0 Then
      If OverShort > 0 Then
         Print #PF, EC(1); "Over"; Tab(27); Align(7, OverShort); EC(4); EC(7)
      ElseIf OverShort < 0 Then
         Print #PF, EC(1); "Short"; Tab(27); Align(7, OverShort); EC(4); EC(7)
      Else
         Print #PF, EC(1); "Perfect"; Tab(27); Align(7, OverShort); EC(4); EC(7)
      End If
      Print #PF, String(40, 45); EC(7) 'Credit Card   Paidouts
   Else
      If OverShort > 0 Then
         PrintLine "Over", 4500: PrintLine Format(OverShort, CurrencyPreFix), -1
      ElseIf OverShort < 0 Then
         PrintLine "Short", 4500: PrintLine Format(OverShort, CurrencyPreFix), -1
      Else
         PrintLine "Perfect", 4500: PrintLine Format(OverShort, CurrencyPreFix), -1
      End If
      DrawLine
    End If
End If

AdjNet = Val(TotalNetSales) - (Val(S_Data(0)) + Val(S_Data(1)))


If PrintReport = 0 Then
  If Check_Option(96) = 0 Then
     Print #PF, "Adjusted Net Sales  "; Align(7, AdjNet); EC(4); EC(7)
     If NewDiners > 0 Then
       Print #PF, "Please Update Frequent Diners !!"
       Print #PF, "You Have New Entries " & Format(NewDiners, "0")
     End If
  Else
    PrintLine "Adjusted Net Sales", 4500: PrintLine Format(AdjNet, CurrencyPreFix), -1
    If NewDiners > 0 Then
     PrintLine "Please Update Frequent Diner Accounts ", 4500: PrintLine "New Entries " & Format(NewDiners, "#"), -1
    End If
  End If
End If
NotTaxed = GetNonTaxSales(WtdNotTaxed, TaxedSales)
If NotTaxed > 0 And PrintReport = 0 Then
  If Check_Option(96) = 0 Then
    Print #PF, "Taxed Sales       : "; Tab(23); Align(6, Val(TaxedSales)); EC(7)
    Print #PF, "Non Taxed Sales   : "; Tab(23); Align(6, Val(NotTaxed)); EC(7)
  Else
    PrintLine "Taxed Sales     ", 4500: PrintLine Format(Val(TaxedSales), CurrencyPreFix), -1
    PrintLine "Non Taxed Sales ", 4500: PrintLine Format(Val(NotTaxed), CurrencyPreFix), -1
  End If
End If

If ChargeBack > 0 And PrintReport = 0 Then
  If Check_Option(96) = 0 Then
    Print #PF, "Tip Fee Income    : "; Tab(23); Align(6, Val(ChargeBack)); EC(7)
  Else
    PrintLine "Tip Fee Income", 4500: PrintLine Format(Val(ChargeBack), CurrencyPreFix), -1
  End If
End If

If PrintReport = 0 Then
   If Check_Option(96) = 0 Then
      Print #PF, EC(7)
      Print #PF, EC(7)
      Print #PF, "Sign ________________________________"; EC(7)
      Print #PF, EC(6); EC(6); EC(5)
      Close #PF
      SendJobToPrinter PrinterPort, "Receipt", False
   Else
      PrintLine "", -1
      PrintLine "", -1
      PrintLine "Sign ________________________________________________", -1
      PrintLine "Sales Report For " & CloseDate, -1
      ReportView.View.EndDoc
      ReportView.View.PrintDoc
   End If
End If
SaveRecap:
Set Void = Nothing
Set DiscountTotals = Nothing
Set DeptTotals = Nothing
Set RLIST = Nothing

Exit Sub
ErrorHandler:
If Err.Number = 438 Then
  Trans_Log "Error On Close Nightly Sales Report " & Err.Number
  Err.Clear
  Exit Sub
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Nightly Sales ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Function NewDiners() As Long
On Error Resume Next
NewDiners = 0
        Dim FD As ADODB.Connection, FR As ADODB.Recordset
        If ConnectADODB(SvrPath.TerminalPath & FrequentDinerDb, FD) = False Then Exit Function
        Sql = "Select * from [FrequentDiners] where [Name] = '<Name>'"
        If ConnectADORS(Sql, FD, FR) = False Then Exit Function
           NewDiners = RecordsCount(FR)
        DbClose FD, FR
End Function
Private Function CashSales() As Currency
On Error Resume Next
        Dim FD As ADODB.Connection, FR As ADODB.Recordset
        ConnectADODB SvrPath.TerminalPath & CheckStatDb, FD
        Sql = "Select Distinct [Payment Type], Sum([Payment Amount]) as TCash from [Payments] where [payment type] = 'Cash' and [Close Out Day] = 0"
        Sql = Sql & " Group By [Payment Type]"
        ConnectADORS Sql, FD, FR
         If Not FR.EOF Then
           CashSales = Val(FR![TCash] & "")
         Else
           CashSales = 0
         End If
        DbClose FD, FR
End Function
Public Sub StoreDayHistory(OverWriteSales As Boolean, ClosingDate As String)
On Error GoTo ErrorHandler
If Not IsDate(ClosingDate) Then ClosingDate = Date
If GrossSales = 0 Then Exit Sub
Dim WD   As Integer
Dim CHDB As ADODB.Connection
Dim CHRS As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & ClosingHistory, CHDB
  Sql = "Delete from [Sales] where Not Isdate([date])"
  CHDB.Execute Sql
  If OverWriteSales Then
    Sql = "Delete from [Sales] where TRY_CAST([date] AS datetime) = '" & Format(CDate(CDate(ClosingDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    CHDB.Execute (Sql)
  End If
  Sql = "Select * from [Sales] where TRY_CAST([date] AS datetime) = '" & Format(CDate(CDate(ClosingDate)), "yyyy-mm-dd hh:nn:ss") & "'"
  ConnectADORS (Sql), CHDB, CHRS
    If CHRS.EOF Then CHRS.AddNew
    CHRS![Date] = ClosingDate
    CHRS![Net Sales] = Val(CHRS![Net Sales] & "") + TotalNetSales 'Net Sales
    CHRS![Gross Sales] = Val(CHRS![Gross Sales] & "") + GrossSales 'Gross Sales
    CHRS![Sales Tax] = Val(CHRS![Sales Tax] & "") + NetSTax 'Net Sales Tax
    CHRS![Special Tax] = Val(CHRS![Special Tax] & "") + NetSpTax 'Net Special Tax
    CHRS![Liquor Tax] = Val(CHRS![Liquor Tax] & "") + NetLiqTax 'Net Liq Tax
    CHRS![Voids] = Val(CHRS![Voids] & "") + VoidSales
    CHRS![Discounts] = Val(CHRS![Discounts] & "") + NetDiscounts 'net discounts
    CHRS![Dining Room Sales] = Val(CHRS![Dining Room Sales] & "") + RegisterSales((HereSale)).NetSaleAmount - RegisterSales((HereSale)).NetDiscAmount  ' Val(SalesType(0, 0, 0)) 'Here_Sales(0)
    CHRS![Dining Room Qty] = Val(CHRS![Dining Room Qty] & "") + RegisterSales((HereSale)).TransactionCount
    CHRS![Go Sales] = Val(CHRS![Go Sales] & "") + RegisterSales((GoSale)).NetSaleAmount - RegisterSales((GoSale)).NetDiscAmount 'Go_Sales(0)
    CHRS![Go Qty] = Val(CHRS![Go Qty] & "") + RegisterSales((GoSale)).TransactionCount  'Go_Trans(0)
    CHRS![Drive Thru Sales] = Val(CHRS![Drive Thru Sales] & "") + RegisterSales((DTSale)).NetSaleAmount - RegisterSales((DTSale)).NetDiscAmount 'DT_Sales(0)
    CHRS![Drive Thru Qty] = Val(CHRS![Drive Thru Qty] & "") + RegisterSales((DTSale)).TransactionCount  'DT_Trans(0)
    CHRS![Call In Sales] = Val(CHRS![Call In Sales] & "") + (RegisterSales((PKUPSale)).NetSaleAmount + RegisterSales(DelSale).NetSaleAmount) - (RegisterSales((PKUPSale)).NetDiscAmount + RegisterSales(DelSale).NetDiscAmount)
    CHRS![Call In Qty] = Val(CHRS![Call In Qty] & "") + RegisterSales((PKUPSale)).TransactionCount + RegisterSales(DelSale).TransactionCount
    CHRS![Deposits] = Val(CHRS![Deposits] & "") + Val(Totals(3)) 'total Deposits
    CHRS![Tenders] = Val(CHRS![Tenders] & "") + Val(Totals(2).Caption) + CreditCardDep 'Total Credit Cards
    CHRS![Paid Outs] = Val(CHRS![Paid Outs] & "") + Val(Totals(4).Caption)  'Total Paid Outs
    CHRS![Accounted For] = Val(CHRS![Accounted For] & "") + Val(Totals(3).Caption) + Val(Totals(2).Caption) + CreditCardDep + Val(Totals(4).Caption)
    CHRS![Adjusted Sales] = Val(CHRS![Adjusted Sales] & "") + (NetSTax + TotalNetSales + NetSpTax + NetLiqTax + ChargeBack)
    CHRS![Over Short] = Val(CHRS![Over Short] & "") + Val(Totals(5).Caption) '+ Val(Totals(4).Caption)) - ((NetSTax + TotalNetSales(0) + NetSpTax(0) + NetLiqTax(0) + ChargeBack))
    CHRS![None Taxed Sales] = Val(CHRS![None Taxed Sales] & "") + Val(NotTaxed)
    CHRS![Tip Charge Back] = Val(CHRS![Tip Charge Back] & "") + Val(ChargeBack)
    CHRS![House Charges] = Val(CHRS![House Charges] & "") + Val(TotalCharges)
    CHRS![Deposits Collected] = Val(CHRS![Deposits Collected] & "") + Val(TDepositCollected)
    CHRS![Table Average] = Val(CHRS![Table Average] & "") + Val(Table_average)
    CHRS![Credit Card Tips] = Val(CHRS![Credit Card Tips] & "") + Val(TotalTip)
    CHRS![Charged Gratuity] = Val(CHRS![Charged Gratuity] & "") + Val(NetGratuity)
    CHRS![Exempt Sales] = ((Val(CHRS![Exempt Sales] & "") + Val(ExemptSale)) + SvrTipsCollect)  '- Val(NetGratuity(0))
    CHRS![Refunds] = Val(CHRS![Refunds] & "") + Format(Val(TotalRefunds & ""), "0.00")
   CHRS.Update
   SetNothing CHRS
   Sql = "Select * from [Deposits]"
   ConnectADORS Sql, CHDB, CHRS
    If OverWriteSales Then CHDB.Execute ("Delete from [Deposits] where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'")
    For DP = 2 To 7
      If Val(S_Data(DP)) <> 0 Then
        CHRS.AddNew
        CHRS![Date] = ClosingDate
        CHRS![Amount] = Val(S_Data(DP).Text)
        CHRS.Update
      End If
    Next DP
    SetNothing CHRS
    If Val(S_Data(7).Text) <> 0 Then
    ChangeFundAdj CloseDate, Val(S_Data(7).Text), CHDB
    End If
    Sql = "Select * from [Paid Outs]"
    ConnectADORS Sql, CHDB, CHRS
      If OverWriteSales Then CHDB.Execute ("Delete from [Paid Outs] where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'")
      With POuts
        For WD = 0 To .rows - 1
         If .TextMatrix(WD, 0) = "" Then Exit For
         If Val(.TextMatrix(WD, 1)) > 0 Then
          CHRS.AddNew
          CHRS![Date] = ClosingDate
          CHRS![Type] = .TextMatrix(WD, 0)
          CHRS![Name] = .TextMatrix(WD, 2)  'Paid Out Name
          CHRS![Amount] = Val(.TextMatrix(WD, 1) & "")
          CHRS![Quantity] = 0
          CHRS.Update
         End If
        Next WD
       End With
     SetNothing CHRS
     Sql = "Select * from [Tenders]"  ''proble Here
     ConnectADORS Sql, CHDB, CHRS
     CHDB.Execute ("Update [Tenders] Set [Tip Total] = 0 Where IsNull([Tip Total])")
      If OverWriteSales Then CHDB.Execute ("Delete from [Tenders] where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'")
         With CCards
           For PC = 0 To .rows - 1
            If .TextMatrix(PC, 0) = "" Then Exit For
            If Val(.TextMatrix(PC, 1)) <> 0 Then
              CHRS.AddNew
              CHRS![Date] = ClosingDate
              CHRS![Name] = .TextMatrix(PC, 0)
              CHRS![Amount] = Val(.TextMatrix(PC, 1))
              CHRS![Quantity] = Val(.TextMatrix(PC, 2))
              CHRS![Tip Total] = Val(.TextMatrix(PC, 3))
              CHRS.Update
            End If
           Next PC
         End With
       Dim OrderDestType As AllDestinations
       Set OrderDestType = New AllDestinations
       ReturnOrderDestination OrderDestType, 1
       If OverWriteSales Then
        Sql = "Delete from [Order Destination] Where TRY_CAST([date] AS datetime) = '" & Format(CDate(ClosingDate), "yyyy-mm-dd hh:nn:ss") & "'"
        CHDB.Execute Sql
       End If
       Sql = "Select * from [Order Destination]"
       ConnectADORS Sql, CHDB, CHRS
       For PC = 1 To OrderDestType.count
          CHRS.AddNew
          CHRS![Date] = ClosingDate
          If OrderDestType.Item(PC).OrderType = "" Then OrderDestType.Item(PC).OrderType = "Missing"
          CHRS![Order Destination] = OrderDestType.Item(PC).OrderType
          CHRS![Quantity] = OrderDestType.Item(PC).OrderQty
          If OrderDestType.Item(PC).OrderType = "Dine In" Then
           CHRS![Amount] = OrderDestType.Item(PC).OrderAmt - TotalRefunds    'Remove from Dine Total
          Else
           CHRS![Amount] = OrderDestType.Item(PC).OrderAmt
          End If
          CHRS.Update
       Next PC
       Set OrderDestType = Nothing
    DbClose CHDB, CHRS
    If Val(Input_Option(89)) > 0 And TotalRefunds >= Val(Input_Option(89)) Then
      EmailAlerts "Excess Refund", "Refund total for today is " & Format(TotalRefunds, "0.00")
    End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Store Day History ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub List_Coupons(DiscountTotals As StoreDiscounts, DeptTotals As StoreDiscounts, WriteHistory As Boolean, WhichDay As Integer, ClosingDate As String)
On Error GoTo ErrorHandler
Dim Db  As ADODB.Connection
Dim Tb  As ADODB.Recordset
Dim CC  As ADODB.Recordset
Dim RTS As Integer
'// no discounts are recorded if you are running dual pricing with the software calculating the percentage

ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select distinct [Menu Item],[Department Report], Sum(Price) as CouponTTl,"
    Sql = Sql & " Count([Menu Item]) as CouponCt, Sum([Inclusive Amount]) as TTLIncl"
    Sql = Sql & " From [check detail] where [Voided] = 0"
    Sql = Sql & " and [Department Report] > ''"
    Sql = Sql & " and [Price] < 0"
    Sql = Sql & " and [Department Report] <> 'Non Sales Discount'"
    If WhichDay = 1 Then
       Sql = Sql & " and [Close Out Day] = 0"
    End If
    Sql = Sql & " Group By [menu item],[Department Report]"
    Sql = Sql & " Order by [menu item] asc"
ConnectADORS Sql, Db, CC
     Do While Not CC.EOF
       If CC![Menu Item] > "" Then
        If DiscountTotals.Item(CC![Menu Item]) Is Nothing Then
            DiscountTotals.add CC![Menu Item]
        End If
        If DeptTotals.Item(CC![Department Report]) Is Nothing Then
          DeptTotals.add CC![Department Report]
        End If
        DiscountTotals.Item(CC![Menu Item]).DiscountName = CC![Menu Item]
        DiscountTotals.Item(CC![Menu Item]).DiscountAmt = DiscountTotals.Item(CC![Menu Item]).DiscountAmt + (Abs(CC!CouponTTl) - Abs(CC!TTLIncl))
        DiscountTotals.Item(CC![Menu Item]).DiscountQty = DiscountTotals.Item(CC![Menu Item]).DiscountQty + CC!Couponct
        DiscountTotals.Item(CC![Menu Item]).DiscountDept = CC![Department Report]
        DeptTotals.Item(CC![Department Report]).DiscountDept = CC![Department Report]
        DeptTotals.Item(CC![Department Report]).DiscountAmt = DeptTotals.Item(CC![Department Report]).DiscountAmt + (Abs(CC!CouponTTl) - Abs(CC!TTLIncl))
        DeptTotals.Item(CC![Department Report]).DiscountQty = DeptTotals.Item(CC![Department Report]).DiscountQty + CC!Couponct
       End If
     CC.MoveNext
   Loop
DbClose Db, CC
If IsDate(ClosingDate) And WriteHistory Then
    Dim CHDB As ADODB.Connection
    Dim CHRS As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & ClosingHistory, CHDB
    Dim WD As Integer
    Trans_Log "Closing Flag for Over Write " & OverWriteSales
    If OverWriteSales Then
      Sql = "delete from [Discounts] where TRY_CAST([date] AS datetime) = '" & Format(CDate(CDate(ClosingDate)), "yyyy-mm-dd hh:nn:ss") & "'"
      CHDB.Execute (Sql)
    End If
    Sql = "Select * from [Discounts] where TRY_CAST([date] AS datetime) = '" & Format(CDate(CDate(ClosingDate)), "yyyy-mm-dd hh:nn:ss") & "'"
    ConnectADORS (Sql), CHDB, CHRS
    For WD = 1 To DiscountTotals.count
      CHRS.AddNew
      CHRS![Date] = ClosingDate
      CHRS![Name] = DiscountTotals.Item(WD).DiscountName
      CHRS![Amount] = DiscountTotals.Item(WD).DiscountAmt
      CHRS![Quantity] = DiscountTotals.Item(WD).DiscountQty
      CHRS.Update
    Next WD
    DbClose CHDB, CHRS
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reset Database ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub Get_Voids()
On Error Resume Next
Set Void = New FinancialItems
Dim TVoids As Integer
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
    TVoids = 0
    ConnectADODB SvrPath.TerminalPath & CheckStatDb, Db
    Sql = "Select * from [closed checks] where [close out day] = " & 0
    Sql = Sql & " and [voided sales] <> " & 0
    Sql = Sql & " or [close out day] = " & 0
    Sql = Sql & " and [voided discounts] <> " & 0
    ConnectADORS Sql, Db, Tb
      Do While Not Tb.EOF
       TVoids = TVoids + 1
       If Void.Item(TVoids) Is Nothing Then
         Void.add (TVoids)
       End If
       Void.Item(TVoids).Amount = Val(Tb.fields("Voided Sales").Value & "")
       Void.Item(TVoids).CheckNo = Tb.fields("Check Number").Value
       Void.Item(TVoids).VoidAmount = Val(Tb.fields("Voided Discounts").Value & "")   'void discount
       Tb.MoveNext
    Loop
    DbClose Db, Tb
End Sub
Public Sub ResetDataBase(path As String, Cd As Integer)
Dim Adb As ADODB.Connection
On Error GoTo ErrorHandler
ConnectADODB path & FinancialDb, Adb
    Sql = "Delete from [DayPayments]"
    Adb.Execute Sql
    Sql = "Update [Paid Ins] set [Close Out Day] = " & Cd & " where [close out day] = 0"
    Adb.Execute Sql
    Sql = "Delete from [Audits]"
    Adb.Execute Sql
    'make sure all audit records are removed from system
    'will stop lost records from register Name changes
    Sql = "Delete from [Charge Back]"
    Adb.Execute Sql

    Sql = "Update [Paid Outs] set [Close Out Day] = " & Cd & " where [close out day] = 0"
    Adb.Execute Sql
        
    Sql = "Update [Refund] set [close out day] = " & Cd & " where [close out day] = 0"
    Adb.Execute Sql
    SetNothing Adb
SetNothing Adb
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reset Database ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub Figure_Shortage()
Dim ActF   As Currency
Dim Act    As Integer
Totals(5).Caption = "": ActF = 0
For Act = 0 To 4
  ActF = ActF + Val(Totals(Act).Caption)
Next Act
If ActF >= Val(CashtoAcctFor) Then
  Totals(5).ForeColor = vbBlack
Else
  Totals(5).ForeColor = vbRed
End If
Totals(5).Caption = Align(7, ActF - Val(CashtoAcctFor))
End Sub
                                                        