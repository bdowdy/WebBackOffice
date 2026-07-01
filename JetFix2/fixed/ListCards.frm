VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form ListCards 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Credit Card List"
   ClientHeight    =   10950
   ClientLeft      =   270
   ClientTop       =   705
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox BatchUtility 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   5295
      Left            =   -4080
      ScaleHeight     =   5265
      ScaleWidth      =   4665
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command6 
         Caption         =   "DataCap SAF Statistics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Batch Pax Devices"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   38
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton ExitButton 
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   37
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CommandButton BtnEnh7 
         Caption         =   "Capture Batch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   2150
      End
      Begin VB.CommandButton BtnEnh5 
         Caption         =   "Voice Auth Pax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   35
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CommandButton SSCommand7 
         Caption         =   "Clear Batch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   34
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Program to Read PDF Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   30
         Top             =   2520
         Width           =   2150
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Utilities"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2295
         Left            =   120
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Frame MakeASelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Make a Selection"
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
      Height          =   5055
      Left            =   1080
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   4150
         Width           =   3375
      End
      Begin VB.ListBox ListNames 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   3375
      End
   End
   Begin MSComDlg.CommonDialog Touch 
      Left            =   240
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox EditCharge 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   4680
      ScaleHeight     =   5025
      ScaleWidth      =   7305
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox AcctData 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "0"
         Top             =   4440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox AcctData 
         Alignment       =   2  'Center
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "0"
         Top             =   4080
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox AcctData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox AcctData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "My Name"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox AcctData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox AcctData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   840
         Width           =   2655
      End
      Begin BTNENHLib4.BtnEnh Quit 
         Height          =   735
         Left            =   1800
         TabIndex        =   17
         Top             =   4080
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
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
         CornerFactor    =   20
         Surface         =   15
         BackColorContainer=   12632256
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "ListCards.frx":0000
         textLT          =   "ListCards.frx":0068
         textCT          =   "ListCards.frx":0080
         textRT          =   "ListCards.frx":0098
         textLM          =   "ListCards.frx":00B0
         textRM          =   "ListCards.frx":00C8
         textLB          =   "ListCards.frx":00E0
         textCB          =   "ListCards.frx":00F8
         textRB          =   "ListCards.frx":0110
         colorBack       =   "ListCards.frx":0128
         colorIntern     =   "ListCards.frx":0152
         colorMO         =   "ListCards.frx":017C
         colorFocus      =   "ListCards.frx":01A6
         colorDisabled   =   "ListCards.frx":01D0
         colorPressed    =   "ListCards.frx":01FA
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
         Office2003ColorSouthMouseOver=   128
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh SSCommand1 
         Height          =   1215
         Left            =   3360
         TabIndex        =   18
         Top             =   840
         Width           =   3615
         _Version        =   589828
         _ExtentX        =   6376
         _ExtentY        =   2143
         _StockProps     =   66
         Caption         =   $"ListCards.frx":0224
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
         textCaption     =   "ListCards.frx":0234
         textLT          =   "ListCards.frx":02AC
         textCT          =   "ListCards.frx":02C4
         textRT          =   "ListCards.frx":02DC
         textLM          =   "ListCards.frx":02F4
         textRM          =   "ListCards.frx":030C
         textLB          =   "ListCards.frx":0324
         textCB          =   "ListCards.frx":033C
         textRB          =   "ListCards.frx":0354
         colorBack       =   "ListCards.frx":036C
         colorIntern     =   "ListCards.frx":0396
         colorMO         =   "ListCards.frx":03C0
         colorFocus      =   "ListCards.frx":03EA
         colorDisabled   =   "ListCards.frx":0414
         colorPressed    =   "ListCards.frx":043E
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
         Office2003ColorBorder=   16777215
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
      Begin BTNENHLib4.BtnEnh SSCommand12 
         Height          =   735
         Left            =   3480
         TabIndex        =   13
         Top             =   3000
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1296
         _StockProps     =   66
         Caption         =   $"ListCards.frx":0468
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
         textCaption     =   "ListCards.frx":047B
         textLT          =   "ListCards.frx":04F9
         textCT          =   "ListCards.frx":0511
         textRT          =   "ListCards.frx":0529
         textLM          =   "ListCards.frx":0541
         textRM          =   "ListCards.frx":0559
         textLB          =   "ListCards.frx":0571
         textCB          =   "ListCards.frx":0589
         textRB          =   "ListCards.frx":05A1
         colorBack       =   "ListCards.frx":05B9
         colorIntern     =   "ListCards.frx":05E3
         colorMO         =   "ListCards.frx":060D
         colorFocus      =   "ListCards.frx":0637
         colorDisabled   =   "ListCards.frx":0661
         colorPressed    =   "ListCards.frx":068B
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
      Begin VB.Label Media 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Media"
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
         Left            =   2280
         TabIndex        =   25
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number"
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
         Left            =   360
         TabIndex        =   9
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Servers Name"
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
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Amount"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Amount"
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
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3855
         Left            =   120
         Top             =   120
         Width           =   7095
      End
   End
   Begin BTNENHLib4.BtnEnh exit 
      Height          =   735
      Left            =   13560
      TabIndex        =   14
      Top             =   9960
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
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
      CornerFactor    =   20
      Surface         =   15
      BackColorContainer=   8438015
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "ListCards.frx":06B5
      textLT          =   "ListCards.frx":071D
      textCT          =   "ListCards.frx":0735
      textRT          =   "ListCards.frx":074D
      textLM          =   "ListCards.frx":0765
      textRM          =   "ListCards.frx":077D
      textLB          =   "ListCards.frx":0795
      textCB          =   "ListCards.frx":07AD
      textRB          =   "ListCards.frx":07C5
      colorBack       =   "ListCards.frx":07DD
      colorIntern     =   "ListCards.frx":0807
      colorMO         =   "ListCards.frx":0831
      colorFocus      =   "ListCards.frx":085B
      colorDisabled   =   "ListCards.frx":0885
      colorPressed    =   "ListCards.frx":08AF
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
      Office2003ColorSouth=   6118749
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   128
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
      Height          =   735
      Left            =   10200
      TabIndex        =   15
      Top             =   9960
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Batch Utilities"
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
      BackColorContainer=   8438015
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "ListCards.frx":08D9
      textLT          =   "ListCards.frx":0957
      textCT          =   "ListCards.frx":096F
      textRT          =   "ListCards.frx":0987
      textLM          =   "ListCards.frx":099F
      textRM          =   "ListCards.frx":09B7
      textLB          =   "ListCards.frx":09CF
      textCB          =   "ListCards.frx":09E7
      textRB          =   "ListCards.frx":09FF
      colorBack       =   "ListCards.frx":0A17
      colorIntern     =   "ListCards.frx":0A41
      colorMO         =   "ListCards.frx":0A6B
      colorFocus      =   "ListCards.frx":0A95
      colorDisabled   =   "ListCards.frx":0ABF
      colorPressed    =   "ListCards.frx":0AE9
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
      Office2003ColorBorder=   16777215
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
   Begin BTNENHLib4.BtnEnh SSCommand21 
      Height          =   735
      Left            =   11880
      TabIndex        =   16
      Top             =   9960
      Width           =   1575
      _Version        =   589828
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   $"ListCards.frx":0B13
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
      BackColorContainer=   8438015
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "ListCards.frx":0B25
      textLT          =   "ListCards.frx":0BA1
      textCT          =   "ListCards.frx":0BB9
      textRT          =   "ListCards.frx":0BD1
      textLM          =   "ListCards.frx":0BE9
      textRM          =   "ListCards.frx":0C01
      textLB          =   "ListCards.frx":0C19
      textCB          =   "ListCards.frx":0C31
      textRB          =   "ListCards.frx":0C49
      colorBack       =   "ListCards.frx":0C61
      colorIntern     =   "ListCards.frx":0C8B
      colorMO         =   "ListCards.frx":0CB5
      colorFocus      =   "ListCards.frx":0CDF
      colorDisabled   =   "ListCards.frx":0D09
      colorPressed    =   "ListCards.frx":0D33
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
      Office2003ColorSouth=   6118749
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   32768
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
      Height          =   735
      Left            =   8040
      TabIndex        =   27
      Top             =   9960
      Width           =   2055
      _Version        =   589828
      _ExtentX        =   3625
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Print Batch"
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
      CornerFactor    =   12
      Surface         =   3
      BackColorContainer=   8438015
      ButtonRaiseFactor=   0
      CaptionWordWrapPerc=   85
      LogPixels       =   96
      SpecialEffectFactor=   2
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "ListCards.frx":0D5D
      textLT          =   "ListCards.frx":0DD3
      textCT          =   "ListCards.frx":0DEB
      textRT          =   "ListCards.frx":0E03
      textLM          =   "ListCards.frx":0E1B
      textRM          =   "ListCards.frx":0E33
      textLB          =   "ListCards.frx":0E4B
      textCB          =   "ListCards.frx":0E63
      textRB          =   "ListCards.frx":0E7B
      colorBack       =   "ListCards.frx":0E93
      colorIntern     =   "ListCards.frx":0EBD
      colorMO         =   "ListCards.frx":0EE7
      colorFocus      =   "ListCards.frx":0F11
      colorDisabled   =   "ListCards.frx":0F3B
      colorPressed    =   "ListCards.frx":0F65
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
      Office2003ColorNorth=   2233099
      Office2003ColorSouth=   2233099
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   2233099
      Office2003ColorSouthMouseOver=   2233099
      Office2003ColorBorderMouseOver=   2233099
      Office2003ColorNorthPressed=   4465941
      Office2003ColorSouthPressed=   4465941
      Office2003ColorBorderPressed=   8209702
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh4 
      Height          =   735
      Left            =   5880
      TabIndex        =   28
      Top             =   9960
      Width           =   2055
      _Version        =   589828
      _ExtentX        =   3625
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Reprint Batch"
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
      CornerFactor    =   12
      Surface         =   3
      BackColorContainer=   8438015
      ButtonRaiseFactor=   0
      CaptionWordWrapPerc=   85
      LogPixels       =   96
      SpecialEffectFactor=   2
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "ListCards.frx":0F8F
      textLT          =   "ListCards.frx":1009
      textCT          =   "ListCards.frx":1021
      textRT          =   "ListCards.frx":1039
      textLM          =   "ListCards.frx":1051
      textRM          =   "ListCards.frx":1069
      textLB          =   "ListCards.frx":1081
      textCB          =   "ListCards.frx":1099
      textRB          =   "ListCards.frx":10B1
      colorBack       =   "ListCards.frx":10C9
      colorIntern     =   "ListCards.frx":10F3
      colorMO         =   "ListCards.frx":111D
      colorFocus      =   "ListCards.frx":1147
      colorDisabled   =   "ListCards.frx":1171
      colorPressed    =   "ListCards.frx":119B
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
      Office2003ColorNorth=   2233099
      Office2003ColorSouth=   2233099
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   2233099
      Office2003ColorSouthMouseOver=   2233099
      Office2003ColorBorderMouseOver=   2233099
      Office2003ColorNorthPressed=   4465941
      Office2003ColorSouthPressed=   4465941
      Office2003ColorBorderPressed=   8209702
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   3
   End
   Begin VSFlex8DAOCtl.VSFlexGrid ShowCharges 
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   14775
      _cx             =   1987601869
      _cy             =   1987591074
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12632256
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
      Rows            =   1000
      Cols            =   8
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
      FrozenRows      =   1
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
   Begin VB.Label NoTipOn 
      BackStyle       =   0  'Transparent
      Caption         =   "Checks without tip 0"
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
      Left            =   240
      TabIndex        =   39
      Top             =   9400
      Width           =   3855
   End
   Begin VB.Label Pathtoreader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4800
      TabIndex        =   29
      Top             =   10920
      Width           =   6015
   End
   Begin VB.Label ItemCount 
      BackStyle       =   0  'Transparent
      Caption         =   "Sending 1000 1 or 1"
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
      Left            =   240
      TabIndex        =   26
      Top             =   9120
      Width           =   3855
   End
   Begin VB.Label BToT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   13440
      TabIndex        =   24
      Top             =   9120
      Width           =   1695
   End
   Begin VB.Label BToT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   23
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label BToT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   22
      Top             =   9120
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Total"
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
      Left            =   11760
      TabIndex        =   21
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tip Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   5640
      Top             =   9840
      Width           =   9615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "List of Current Charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   9495
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   9615
      Left            =   120
      Top             =   120
      Width           =   15135
   End
End
Attribute VB_Name = "ListCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BatchTotal      As Currency
Public CheckBelongsTo  As String
Public PreAuthSale     As Boolean
Public PauseButton     As Double
Public BatchInProgress As Boolean

Public Function PrintBatch(BatchResult As String, RecCount As Long) As Boolean
Dim PrinterPort         As String
Dim PrintBatchName      As String
Dim DepositsCollected   As Currency
PrintBatch = True
If Check_Option(96) = 0 And Check_Option(130) = 0 Then
   Printer_Drivers "Receipt"
   CenterText = EC(8)
   LeftText = EC(9)
   PF = FreeFile
   PrinterPort = "Batch Report"
tryagain:
End If
SvBatch:
Dim vDb         As ADODB.Connection
Dim vRc         As ADODB.Recordset
Dim BatchDate   As String
Dim VSale       As Currency
Dim TCharge     As Currency
Dim TTip        As Currency
Dim DataStr     As String

If ConnectADODB(SvrPath.TerminalPath & DebitDb, vDb) = False Then Exit Function
  Sql = "Select * from [Debit] Where [Tender] = 'EMV' or [Tender] = 'EMV Refund' order by [Server Name],[Check Number] asc"
ConnectADORS Sql, vDb, vRc
RecCount = RecordsCount(vRc)
If vRc.EOF Then
    Call See_Msg("Nothing to Batch", 5, "")
    DbClose vDb, vRc
    PrintBatch = False
    Exit Function
End If
DbClose vDb, vRc
DepositsCollected = 0: TCharge = 0: TTip = 0
   If ConnectADODB(SvrPath.TerminalPath & DebitDb, vDb) = False Then Exit Function
    Sql = "Select * from [Debit] Where [Tender] = 'EMV' or [Tender] = 'EMV Refund' or [Tender] = 'QRPayment' order by [Server Name],[Check Number] asc"
    ConnectADORS Sql, vDb, vRc
   DataStr = ""
   If Not vRc.EOF Then
    BatchDate = Format(vRc.fields("Date").Value, "MM/DD/YY")
   End If
   If Not IsDate(BatchDate) Then BatchDate = Date
   If Check_Option(96) = 0 And Check_Option(130) = 0 And GetConfigInfo("Printing Options", "EMV Batch 80 Column Only", "No", "FUNCTIONS.INI") = "DISABLED" Then
     Open PrinterPort For Output Shared As #PF
      Print #PF, EC(1); CenterText; Heading; EC(7)
      Print #PF, LeftText; EC(7)
      Print #PF, ; EC(4); EC(7)
      Print #PF, Format(Time, "HH:MM AMPM");
      Print #PF, Tab(24); "Date  "; Trim(Date); EC(7)
      Print #PF, String(40, 45); EC(7)
      Do While Not vRc.EOF
         Print #PF, vRc![Account Name]; Tab(19); "Card xxxx-xxxx-xxxx-"; vRc.fields("Card Number").Value & ""; EC(7)
         Print #PF, "Time "; vRc![Time]; Tab(24); "Date "; vRc![Date]; EC(7)
         Print #PF, "Server  "; vRc![Server Name]; Tab(24); "Check No "; ReturnCardStatus(vRc![Check Number]); EC(7)
         Print #PF, "Sale Amt "; Format(vRc![Amount], "0.00"); Tab(24); "Tip Amt "; Format(vRc![tip], "0.00"); EC(7)
         Print #PF, "Total Charge "; Format(vRc![Amount] + Val(vRc![tip] & ""), "0.00")
         Print #PF, "Auth. No. "; vRc![Account Number]; EC(7)
         Print #PF, String(40, 45); EC(7)
         TCharge = TCharge + Format(vRc.fields("Amount").Value, "0.00")
         TTip = TTip + Format(vRc.fields("Tip").Value, "0.00")
        vRc.MoveNext
      Loop
      Print #PF, "Sales : "; Format(TCharge, "0.00"); Tab(25); "Tips :"; Format(TTip, "0.00") & EC(7)
      Print #PF, "Charge Total "; Format(TCharge + TTip, "0.00") & EC(7)
      Print #PF, EC(7)
      Print #PF, EC(7)
      Print #PF, "Please Wait For Batch Result"
      Print #PF, EC(6) & EC(5)
      Close #PF
      SendJobToPrinter PrinterPort, "Receipt", False
      '130 disable batch print out
    End If
    TCharge = 0: TTip = 0
    If Not vRc.BOF Then vRc.MoveFirst
     Do While Not vRc.EOF
       DataStr = DataStr & Format(vRc.fields("Date").Value, "MM/DD/YY") & "|"
       DataStr = DataStr & Format(vRc.fields("Time").Value, "HH:MM ampm") & "|"
       DataStr = DataStr & vRc.fields("Check Number").Value & "|"
       DataStr = DataStr & vRc.fields("Account Name").Value & "|"
       DataStr = DataStr & vRc.fields("Account Number").Value & "" & "|"
       DataStr = DataStr & vRc.fields("Server Name").Value & "|"
       DataStr = DataStr & Format(vRc.fields("Amount").Value, "0.00") & "|"
       TCharge = TCharge + Format(vRc.fields("Amount").Value, "0.00")
       TTip = TTip + Format(vRc.fields("Tip").Value, "0.00")
       DataStr = DataStr & Format(vRc.fields("Tip").Value, "0.00") & "|"
       DataStr = DataStr & Format(vRc.fields("Amount").Value + vRc.fields("Tip").Value, "0.00") & "|"
       DataStr = DataStr & vRc.fields("Card Number").Value & "" & ";"
      vRc.MoveNext
     Loop
     DbClose vDb, vRc
     ReportView.View.Orientation = orPortrait
     SetPage ReportView.View, "Credit Cards " & Format(BatchDate, "MM/DD/YY")
     Dim F As String
     Dim H As String
     H = "Date|Time|Check #|Card Type|Auth|Server Name|Sale Amt|Tip Amt|Total|Last 4;"
     F = "+<950|+<950|^1000|^1500|^1200|^1500|+>1000|+>1000|+>1000|+>1000;"
     ReportView.View.TableBorder = tbBoxColumns
     ReportView.View.addTable F, H, DataStr, vbBlack, vbWhite
     DataStr = Format(TCharge, "0.00") & "|"
     DataStr = DataStr & Format(TTip, "0.00") & "|"
     DataStr = DataStr & Format(TCharge + TTip, "0.00") & ";"
     H = "Total Charges|Total Tips|Batch Total;"
     F = "+>2000|+>2000|+>2000;"
     ReportView.View.TableBorder = tbBoxRows
     ReportView.View.addTable F, H, DataStr, vbBlack, vbWhite
'End If
ReportView.View.EndDoc
Dim Wl As Integer
Dim TXTFileInfo() As Variant
Dim Email         As String
If SendEmails(6).Send Then
   ReportView.EmailEOD "BatchReport"
End If
ReportView.VSPDF1.Compress = vspfdCompressAll
ReportView.VSPDF1.ConvertDocument ReportView.View, SvrPath.TerminalPath & "PDF Files\" & "Batch" & Format(Now, "MMDDYYYY-HHMMSS") & ".PDF"
If Check_Option(96) = 1 And Check_Option(130) = 0 Then
   ReportView.View.PrintDoc
End If
Call MemoryCheck(ReportView)
BatchTotal = TCharge + TTip
Exit Function
ErrorHandler:
DbClose vDb, vRc
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Print Batch ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Private Function ReturnCardStatus(DataIn As Long) As String
If DataIn = -9999 Then
    ReturnCardStatus = "Deposit"
ElseIf DataIn = -8888 Then
    ReturnCardStatus = "Return"
Else
    ReturnCardStatus = CStr(DataIn)
End If
End Function
Private Sub PrtCCSlip(HeaderData As String, HeaderReason As String, AuthorNo As String)

 Dim CCD As CCData
 With CCD
  .HeaderData = HeaderData
  .HeaderReason = HeaderReason
  .AuthNo = AuthorNo
  .SaleAmt = Format(AcctData(1).Text, "0.00")
  .TipAmt = Format(AcctData(2).Text, "0.00")
  .AccNo = AcctData(0).Text
  .CustName = ""
  .CType = Media.Caption & " Copy"
  .ServerName = AcctData(3).Text
  .TBLName = ""
  .ChkNO = CStr(AcctData(5).Text)
  .ProcessorResponse = "Approved"
  .Datetime = UserVoid.Time
  .ckTotal = Format(AcctData(1).Text, "0.00")
  .ExpDate = "" 'AcctData(4).Text
 End With
 PrintSlip CCD
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Print Slip ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub LoadGrid(OrderBy As String)
On Error GoTo ErrorHandler
Dim NoTips      As Integer
Dim Chk_Total   As Currency
Dim Tip_Total   As Currency
Dim Chg_Total   As Currency
Chk_Total = 0: Tip_Total = 0: Chg_Total = 0
Dim TSName      As String
Dim TsAddress   As String
Dim TXML        As String
With ShowCharges
  .Clear flexClearScrollable, flexClearData
  .Row = 0
  .Col = 0
  .Text = "Auth. No."
  .ColWidth(0) = 1500 * WidScale1024
  .Col = 1
  .Text = "Status"
  .ColWidth(1) = 1500 * WidScale1024
  .Col = 2
  .Text = "Check No"
  .ColWidth(2) = 1500 * WidScale1024
  .Col = 3
  .Text = "Servers"
  .ColWidth(3) = 3000 * WidScale1024
  .Col = 4
  .Text = "Account Number"
  .ColWidth(4) = 0   '2080 * WidScale
  .Col = 5
  .Text = "Charge Amt"
  .ColWidth(5) = 1500 * WidScale1024
  .Col = 6
  .Text = "Gratuity"
  .ColWidth(6) = 1500 * WidScale1024
  .Col = 7
  .Text = "Account Number"
  .ColWidth(7) = 4000 * WidScale1024
End With
Dim Cdb As ADODB.Connection, Tb As ADODB.Recordset
    If ConnectADODB(SvrPath.TerminalPath & DebitDb, Cdb) = False Then Exit Sub
    If OrderBy = "Gratuity" Then
     Sql = "Select * from [Debit] where [Tender] = 'QRPayment' or [Tender] = 'EMV' or [Tender] = 'EMV Refund' order by [Tip] asc"
    ElseIf OrderBy = "Amount" Then
     Sql = "Select * from [Debit] where [Tender] = 'QRPayment' or [Tender] = 'EMV' or [Tender] = 'EMV Refund' order by [Amount] asc"
    ElseIf OrderBy = "Check Number" Then
     Sql = "Select * from [Debit] where [Tender] = 'QRPayment' or [Tender] = 'EMV' or [Tender] = 'EMV Refund' order by [Check Number] asc"
    Else
     Sql = "Select * from [Debit] where [Tender] = 'QRPayment' or [Tender] = 'EMV' or [Tender] = 'EMV Refund' order by [Server Name],[Check Number] asc"
    End If
    ConnectADORS Sql, Cdb, Tb
ItemCount.Caption = "Item Count " & RecordsCount(Tb)
Tip_Total = 0: Chk_Total = 0: Chg_Total = 0

With ShowCharges
.rows = RecordsCount(Tb) + 1  ''Changed 08.20.2008
.Row = 0
NoTips = 0
  Do While Not Tb.EOF
   .Col = 0
   .Row = .Row + 1
   .Text = Check_Null(Tb.fields("Account Number").Value)
   .Col = 1
   If Tb![Check Number] = -9999 Then
   .TextMatrix(.Row, 1) = "Deposit"
   Else
   .Text = Tb.fields("Account Number").Value & ""   ' "EMV Sale"
   End If
   .Col = 2
   .Text = Val(Tb![Check Number] & "")
   .Col = 3
   .Text = Check_Null(Tb.fields("Server Name").Value)
   .Col = 4
    .Text = ""
   .Col = 5
   .Text = Format(Check_Null(Tb!Amount), "0.00")
     Chk_Total = Chk_Total + Format(Check_Null(Tb!Amount), "0.00")
   .Col = 6
   If Check_Null(Tb.fields("Tip").Value) = 0 Then NoTips = NoTips + 1
   .Text = Format(Check_Null(Tb.fields("Tip").Value), "0.00")
   Tip_Total = Tip_Total + Format(Check_Null(Tb.fields("Tip").Value), "0.00")
   .Col = 7
   .Text = "XXXX-XXXX-XXXX-" & Tb.fields("Card Number").Value & ""
  Tb.MoveNext
  Loop
End With
DbClose Cdb, Tb
BToT(0).Caption = Format(Chk_Total, "0.00")
BToT(1).Caption = Format(Tip_Total, "0.00")
BToT(2).Caption = Format(Chk_Total + Tip_Total, "0.00")
NoTipOn.Caption = "Charges without tips " & NoTips
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Description & Sql, Err.Number, "Load Grid List Cards ")
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
'FindAllCharges
End Sub
Private Sub AcctData_Click(Index As Integer)
SSCommand1_Click
End Sub

Private Sub BtnEnh2_Click()
End Sub

Private Sub BtnEnh3_Click()

End Sub
Private Sub BtnEnh4_Click()
 On Error Resume Next
 If LoggedOnAs(0).EmpAccessLevel < Access_Required("Credit Card Tips") Then
   See_Msg "No Access", 5, ""
   Exit Sub
 End If
 If GetSetting("Micro$ale", "Register Setup", "Acrobat reader") = "" Then
    See_Msg "Please Select a Reader", 5, ""
    Exit Sub
 End If
 Load_What = "BATCH"
 Find_What = "Print Batch"
 Select_What.Show 1
 If GetSelected > "" Then
  Shell GetSetting("Micro$ale", "Register Setup", "Acrobat reader") & " " & SvrPath.TerminalPath & "PDF Files\" & GetSelected, vbMaximizedFocus
 End If
End Sub
Private Sub BtnEnh5_Click()
            Dim Results As String
            'Dim PXFO As Object
            Dim sales As AllEMVSales
            Set sales = New AllEMVSales
            Check_no = 1
            If sales.Item(CStr(Check_no)) Is Nothing Then
              sales.add CStr(Check_no)
            End If
            Dim FAmt As Currency
            FAmt = ShowNumberPad("Enter Amount", False, 0, 0)
            sales.Item(CStr(Check_no)).Cashier = LoggedOnAs(0).EMPName
            sales.Item(CStr(Check_no)).Term = MyRegisterName
            sales.Item(CStr(Check_no)).IpAddress = ""
            sales.Item(CStr(Check_no)).SalesAmt = FAmt 'Format(Val(RefundData(1).Caption) + (Val(RefundData(2).Caption) + Val(RefundData(5).Caption)), "0.00")
            sales.Item(CStr(Check_no)).TipAmt = 0
            sales.Item(CStr(Check_no)).AccountNo = "Prompt"
            sales.Item(CStr(Check_no)).XMLString = ""
            sales.Item(CStr(Check_no)).TransNo = "1"  ' CStr(Check_no)
            sales.Item(CStr(Check_no)).TranType = "FORCEAUTH"
            See_Msg "Please Swipe Card", 5, ""
            Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
            If LoadEMVDevice(MyRegisterName, EmvEquip) = "Failed" Then
              See_Msg "No Device Found", 5, "Check Device Setup"
              Exit Sub
            End If
            Dim PMG As PAXMSG
            Set PMG = New PAXMSG
            Dim PXFO As EMVProcessLib.Process
            Set PXFO = New Process
            Results = PXFO.ProcessSale(sales, 6, PMG, EmvEquip)    'voice auth
            Set sales = Nothing
            If UCase(PMG.Item(1).ResultTxt) = "APPROVED" Then
               See_Msg "Voice Approved.", 5, ""
               Dim apychg As ver11DebitCards.Accounting
               Set apychg = New ver11DebitCards.Accounting
               Dim RA As ver11DebitCards.AccountInfo
               With RA
                 .AccountName = PMG.Item(1).CardType
                 .AuthCode = PMG.Item(1).AuthCode
                 If GetSelected = "Debit" Then
                 .ChargeType = ctDebit
                 ElseIf GetSelected = "EBT" Then
                 .ChargeType = ctEBT
                 Else
                 .ChargeType = ctSale
                 End If
                 .CheckNo = Check_no
                 .ckTotal = Format(CCur(PMG.Item(1).ApprovedAmount), "0.00")
                 .ServerName = LoggedOnAs(0).EMPName
                 .TenderName = "EMV Voice Auth"
                 .TipAmt = 0
                 .XMLDataString = PMG.Item(1).ExtraData
                 .RegisterName = MyRegisterName
                 .CashierName = ReturnMyCashier
               End With
               apychg.ChargeAccount RA, SvrPath.TerminalPath
               SaveEMVPayments RA.AuthCode, RA.CheckNo, RA.ServerName, RA.ckTotal, RA.TipAmt, RA.AccountName
               Set apychg = Nothing
               Set PMG = Nothing
               Set PXFO = Nothing
            Else
                See_Msg "Voice Auth Failed", 5, PMG.Item(1).Message
                Set PMG = Nothing
                Set PXFO = Nothing
                Exit Sub
            End If
End Sub

Private Sub BtnEnh6_Click()

End Sub

Private Sub BtnEnh7_Click()
On Error Resume Next
If Check_Option(79) = 1 Then
 See_Msg "Tips Not Enabled", 5, ""
 Exit Sub
End If
Dim Cdb         As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim NotGrat     As Integer
Dim WTime       As Long
Dim Rst         As String
Dim R           As String
Dim BC          As Object
Dim result      As String
ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
 Sql = "Select * from [Debit] Where [Tender] = 'EMV' or [Tender] = 'EMV Refund'"
ConnectADORS Sql, Cdb, Tb
  If Tb.EOF Then
   See_Msg "Batch Already Closed", 5, "No transaction(s) found in batch"
   DbClose Cdb, Tb
   Exit Sub
  End If
  DbClose Cdb, Tb
 
 If (IsDataCapInterface(EMVINTERFACE) And EMVSUPPORTPREAUTH = 1) Then
   If IsPadPresent = False Then Exit Sub
 Else
   See_Msg "Not set for preauth", 5, ""
 End If

ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
  Sql = "Select * from [Debit] where [Tender] = 'EMV' and [Charge] = -1"  ' where [Check Number] > 0"
ConnectADORS Sql, Cdb, Tb
Do While Not Tb.EOF
  ItemCount.Caption = "Sending for Capture " & Tb.fields("Check Number").Value
  ItemCount.Refresh
   result = ApplyPAXTip(Tb.fields("Check Number").Value, Tb.fields("Server Name").Value, Format(Tb.fields("Tip").Value, "0.00"), Tb.fields("Account Number").Value, DoEMVCapture, Tb.fields("Card Number").Value, Tb.fields("Account Name").Value)
   If result = "Approved" Then
     
   Else
     Trans_Log "Batch Capture Error " & result
     PrintFailedCapture "EMV Capture Failed " & result, Tb.fields("Check Number").Value, Tb.fields("Card Number").Value, Format(Tb.fields("Amount").Value, "0.00"), Format(Tb.fields("Tip").Value, "0.00")
   End If
  Tb.MoveNext
Loop
See_Msg "EMV Batch Capture Complete", 5, "Closed By Processor."

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
 With Touch
     .ShowOpen
     SaveSetting "Micro$ale", "Register Setup", "Acrobat reader", .FileName
 End With
 See_Msg "Saved...", 5, ""
 Pathtoreader.Caption = "Path to PDF Reader " & GetSetting("Micro$ale", "Register Setup", "Acrobat reader")
End Sub

Private Sub Command4_Click()
MakeASelection.Visible = False
MakeASelection.Left = -5000
End Sub

Private Sub Command5_Click()
If IsPAXInterface(EMVINTERFACE) Then
   Trans_Log "Sending Closing pax equipment batch"
   'Dim BC          As Object
   'Set BC = New EMVProcessLib.Process
   BatchPaxDevices Date
Else
   See_Msg "No Pax Equipment", 5, ""

End If

End Sub

Private Sub Command6_Click()
Dim EmvEquip        As EMVProcessLib.EmvDeviceHardWare
If LoadEMVDevice(MyRegisterName, EmvEquip) = "Failed" Then
   See_Msg "No Device Found", 5, "Check Device Setup"
   Exit Sub
End If
Dim R As String
Dim DCD As EMVProcessLib.Process
Set DCD = New EMVProcessLib.Process
'R = DCD.ReturnSAFStats(EmvEquip)
Trans_Log "SAF_Stat " & R
MsgBox (R)
R = DCD.SAFForward(EmvEquip)
MsgBox (R)
Trans_Log "SAF_Forward " & R
Set DCD = Nothing
End Sub

Private Sub Exit_Click()
If SSCommand21.Caption = "Sending....." Then
   If See_Msg("You Should Not quit Batch In Progress.", 4, "Sure you want to quit") = 6 Then
    Trans_Log "Credit Card Batch Quit "
    Exit Sub
   End If
End If
SSCommand3.Enabled = True
SSCommand21.Enabled = True
SSCommand6.Enabled = True
Trans_Log "Credit Card Exit "
Call MemoryCheck(ListCards)
End Sub
Private Sub ExitButton_Click()
BatchUtility.Visible = False
BatchUtility.Left = -4000
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ERH
ResizeScreen1024 Me
If SSCommand3.Enabled = False Then
  Exit Sub
Else
  LoadGrid ""
End If
BatchInProgress = False
Pathtoreader.Caption = "Path to PDF Reader " & GetSetting("Micro$ale", "Register Setup", "Acrobat reader")
BtnEnh5.Visible = True
SSCommand21.Caption = "Batch Charges"
SSCommand3.Enabled = True
Exit Sub
ERH:
Err.Clear
Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set NetEpay = Nothing
End Sub
Private Sub Quit_Click()
EditCharge.Visible = False
EditCharge.Left = -4000
EditCharge.Refresh
End Sub
Public Sub LoadEMVChecks(MyCheckNo As Long)
SSCommand3.Enabled = False
SSCommand21.Enabled = False
SSCommand6.Enabled = False
BtnEnh4.Enabled = False
Dim Cdb As ADODB.Connection
'Dim Db  As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
  If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" And MyCheckNo > -1 Then
      Sql = "Update [Debit] set [cashier Name] = '" & ReturnMyCashier & "' where [Close out day] = 0 "
      Sql = Sql & " and [Check Number] = " & MyCheckNo
      Sql = Sql & " and [Tender] = 'EMV'"
      Sql = Sql & " and [Server Name] = [Cashier Name]"
     Cdb.Execute Sql
     '// us this for sites that are allowing server to close checks to credit but tips are paid by cashier
     '// Lewis 12.04.2023
  End If
  Sql = "Select * from [Debit] where [Close out day] = 0 "
  If MyCheckNo = -1 Then
    If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" And CheckBelongsTo > "" Then
      Sql = Sql & " and [Server Name] = '" & CheckBelongsTo & "'"
    Else
      Sql = Sql & " and [Server Name] = '" & LoggedOnAs(0).EMPName & "'"
    End If
  Else
   Sql = Sql & " and [Check Number] = " & MyCheckNo
  End If
  Sql = Sql & " and [Tender] = 'EMV'"
ConnectADORS Sql, Cdb, Tb
ItemCount.Caption = "Item Count " & RecordsCount(Tb)
With ShowCharges
.Clear (flexClearScrollable)
.Row = 0
  .TextMatrix(0, 0) = "Auth. No."
  .TextMatrix(0, 1) = "Status"
  .TextMatrix(0, 2) = "Check No"
  .TextMatrix(0, 3) = "Servers"
  .TextMatrix(0, 4) = "Account Number"
  .TextMatrix(0, 5) = "Charge Amt"
  .TextMatrix(0, 6) = "Gratuity"
  .TextMatrix(0, 7) = "Account Number"
   Tip_Total = 0: Chk_Total = 0: Chg_Total = 0
  .rows = RecordsCount(Tb) + 1  ''Changed 08.20.2008
  .Row = 0
  
  
  Do While Not Tb.EOF
   
   .Col = 0
   .Row = .Row + 1
   .Text = Check_Null(Tb.fields("Account Number").Value)
   .Col = 1
   If Tb![Check Number] = -9999 Then
   .TextMatrix(.Row, 1) = "Deposit"
   Else
   .Text = "EMV Sale"
   End If
   .Col = 2
   .Text = Val(Tb![Check Number] & "")
   .Col = 3
   .Text = Check_Null(Tb.fields("Server Name").Value)
   .Col = 4
    .Text = ""
   .Col = 5
   .Text = Format(Check_Null(Tb!Amount), "0.00")
     Chk_Total = Chk_Total + Format(Check_Null(Tb!Amount), "0.00")
   .Col = 6
   .Text = Format(Check_Null(Tb.fields("Tip").Value), "0.00")
   Tip_Total = Tip_Total + Format(Check_Null(Tb.fields("Tip").Value), "0.00")
   .Col = 7
   .Text = Tb.fields("Account Name").Value & ""
  Tb.MoveNext
  Loop
End With
DbClose Cdb, Tb
BToT(0).Visible = False
BToT(1).Visible = True
BToT(1).Caption = Format(Tip_Total, "0.00")
BToT(2).Visible = False
End Sub
Private Sub ShowCharges_Click()
On Error GoTo ErrorHandler
'If SSCommand21.Enabled = False Then Exit Sub
Dim ChkNO       As String
Dim ReturnAmt   As String
Dim Auth        As String
Dim AcctNo      As String
Dim Cdb         As ADODB.Connection
Dim Tb          As ADODB.Recordset
If ShowCharges.Row = 0 And SSCommand21.Enabled = False Then Exit Sub
If ShowCharges.Row = 0 Then
   If ShowCharges.Col = 6 Then
     LoadGrid "Gratuity"
   ElseIf ShowCharges.Col = 5 Then
     LoadGrid "Amount"
   ElseIf ShowCharges.Col = 2 Then
     LoadGrid "Check Number"
   Else
     LoadGrid ""
   End If
    Exit Sub
End If
ShowCharges.Col = 1
If ShowCharges.Text = "Void" Then
   See_Msg "Charge voided", 5, ""
   SSCommand1.Enabled = False
Else
   SSCommand1.Enabled = True
End If
ShowCharges.Col = 0
Auth = ShowCharges.Text & ""
ShowCharges.Col = 5
ReturnAmt = ShowCharges.Text
ShowCharges.Col = 4
AcctNo = ShowCharges.Text
ShowCharges.Col = 2
ChkNO = ShowCharges.TextMatrix(ShowCharges.Row, 2)
If Val(ChkNO) = 0 Then Exit Sub
PreAuthSale = False
ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
    Sql = "Select top 1 * from [debit] where [Account Number] = '" & Auth & "'"
    Sql = Sql & " and [check Number] = " & Val(ChkNO)
    Sql = Sql & " and convert(decimal(10,2),[Amount]) = " & Format(CCur(ReturnAmt), "0.00")
ConnectADORS Sql, Cdb, Tb
       ItemCount.Caption = "Item Count " & RecordsCount(Tb)
       If Not Tb.EOF Then
        If Tb.fields("Account Name").Value = "Debit" Then
           See_Msg "Cannot Modify Debit Charge", 5, ""
           DbClose Cdb, Tb
           Exit Sub
        End If
        If Tb.fields("Charge").Value = -1 Then
          PreAuthSale = True
        Else
          PreAuthSale = False
        End If
        AcctData(0).Text = Tb.fields("Card Number").Value & ""
        AcctData(1).Text = Format(Tb!Amount, "0.00")
        AcctData(2).Text = Format(Tb!tip, "0.00")
        AcctData(5).Text = Tb![Check Number] & ""
        'AcctData(4).Text = ""
        Media.Caption = Tb.fields("Account Name").Value & ""
        AcctData(3).Text = Tb.fields("server Name").Value & ""
        AcctData(6).Text = Auth
        If Check_Option(34) = 1 And Input_Option(10) = "Open/Close" Then
          CheckBelongsTo = Tb.fields("server Name").Value & ""
        Else
          CheckBelongsTo = ""
        End If
        If Tb.fields("Tender").Value = "QRPayment" Then
           Media.Caption = "QRPayment"
        End If
       End If
DbClose Cdb, Tb
CenterObject Me, EditCharge
EditCharge.Visible = True
EditCharge.Refresh
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Show Charges ")
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
On Error Resume Next
Dim RTN_Response As String
Dim TipAmount As Currency
If Media.Caption = "QRPayment" Then
 See_Msg "Cannot Modify QR Payments", 5, ""
 Exit Sub
End If
If Can_Continue(SSCommand1.Caption) = False Then Exit Sub
   If System_config = FullService Then
      If Trim$(AcctData(3).Text) <> "Phone Order" Then
       If AreAuditsOpen(Trim$(AcctData(3).Text)) = False Then  'Used to check for closed audit return true for closed
         If See_Msg("Server Audit Closed please reopen server audit ", 4, "Yes to continue with audit closed") <> 6 Then Exit Sub
            Trans_Log "Audit Closed : Tip Adjust"
         End If
       End If
      End If
      If PreAuthSale And IsDejavooInterface(EMVINTERFACE) Then
       If IsThisCheckOpen(Val(AcctData(5).Text)) Then
        See_Msg "This check is open please close the check first", 5, ""
        Exit Sub
       End If
      End If
      AcctData(2).Text = ShowNumberPad("Enter Tip Amount", False, 0, 0)
      
      If Nz(AcctData(2).Text, 0) = 0 Then Exit Sub
      
      AcctData(2).Text = PlaceDec(AcctData(2).Text)
      TipAmount = Round2Money(Val(AcctData(2).Text))
      
      If TipAmount >= Val(AcctData(1).Text) Then
        If See_Msg("Tip amount is greater than the Check Total " & Format(AcctData(2).Text, "0.00"), 4, "Exccessive Tip Amount, Continue?") <> 6 Then Exit Sub
      End If
      If See_Msg("Tip Amount Is " & Format(AcctData(2).Text, "0.00"), 4, "OK to Continue") <> 6 Then Exit Sub
      If Val(Input_Option(66)) > 0 And Nz(AcctData(1).Text, 0) > 0 Then
         If Val(AcctData(2).Text) > (Val(AcctData(1).Text) * (Val(Input_Option(66)) / 100)) Then
            See_Msg "Exccessive Tip Amount", 5, "Manager Access Required"
            If Can_Continue("Managers Menu") = False Then Exit Sub
         End If
      End If
      If IsDataCapInterface(EMVINTERFACE) Then
         ApplyPAXTip Val(AcctData(5).Text), Trim$(AcctData(3).Text), Format(TipAmount, "0.00"), AcctData(6).Text, SkipEMVCapture, AcctData(0).Text, ""
      Else
       If PreAuthSale And IsPAXInterface(EMVINTERFACE) Then
          ApplyPAXTip Val(AcctData(5).Text), Trim$(AcctData(3).Text), Format(TipAmount, "0.00"), AcctData(6).Text, SkipEMVCapture, AcctData(0).Text, ""
       ElseIf PreAuthSale And IsDejavooInterface(EMVINTERFACE) Then
          ApplyPAXTip Val(AcctData(5).Text), Trim$(AcctData(3).Text), Format(TipAmount, "0.00"), AcctData(6).Text, DoEMVCapture, AcctData(0).Text, ""
          If Nz(TipAmount, 0) > 0 Then
           ApplyPAXTip Val(AcctData(5).Text), Trim$(AcctData(3).Text), Format(TipAmount, "0.00"), AcctData(6).Text, ApplyEMVTip, AcctData(0).Text, ""
          End If
       Else
          ApplyPAXTip Val(AcctData(5).Text), Trim$(AcctData(3).Text), Format(TipAmount, "0.00"), AcctData(6).Text, ApplyEMVTip, AcctData(0).Text, ""
       End If
      End If
      If SSCommand3.Enabled = False Then
        LoadEMVChecks -1
      Else
        LoadGrid ""
      End If
      Quit_Click
End Sub
Private Sub SSCommand12_Click()
'CardType = CardName(AcctData(0))
If Val(AcctData(5).Text) = -8888 Then
  PrtCCSlip "Yes", "Return", AcctData(6).Text
Else
 If AcctData(6).Text = "Void" Then
  PrtCCSlip "Yes", "Void", AcctData(6).Text
 Else
  PrtCCSlip "No", "Sale", AcctData(6).Text
 End If
End If
EditCharge.Visible = False
EditCharge.Left = -4000
EditCharge.Refresh
End Sub
Private Sub SSCommand13_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
End Sub
Private Sub SSCommand2_Click()
End Sub
Private Sub SSCommand2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
End Sub
Public Sub SSCommand21_Click()
If BatchInProgress Then
  If See_Msg("Batch In Progress", 4, "Close Batch Again?") <> 6 Then Exit Sub
  Trans_Log "Batch In Progress Over Ride"
End If
BatchInProgress = True
BatchCharges
BatchInProgress = False
End Sub
Private Sub BatchStatus(Status As String, BatchDate As String)
  On Error Resume Next
  Dim CHDB As ADODB.Connection
  Dim CHRS As ADODB.Recordset
  If Not IsDate(BatchDate) Then BatchDate = Format(Date, "MM/DD/YYYY")
  ConnectADODB SvrPath.TerminalPath & ClosingHistory, CHDB
  ConnectADORS "Select * from [Tenders] where TRY_CAST([date] AS datetime) = '" & Format(CDate(BatchDate), "yyyy-mm-dd hh:nn:ss") & "'", CHDB, CHRS
       CHRS.AddNew
       CHRS![Date] = Format(BatchDate, "MM/DD/YYYY")
       CHRS![Name] = "1-Batch:" & Trim$(Left$(Status, 43))
       CHRS![Amount] = 0
       CHRS![Quantity] = 0
       CHRS![Tip Total] = 0
       CHRS.Update
  DbClose CHDB, CHRS
End Sub
Private Sub BatchDejavooDevices(BatchDate As String)
On Error GoTo ErrorHandler
Dim BStation    As Integer
Dim Rst         As String
Dim Cdb         As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim FailedBatch As Boolean
Dim BC          As EMVProcessLib.Process
FailedBatch = False
BStation = 1
ConnectADODB DebitDb, Cdb
    Sql = "Select * from [EMVDeviceList] Where [DeviceName] > ''"
ConnectADORS Sql, Cdb, Tb
 If Tb.EOF Then
   See_Msg "No Device Installed", 5, "Cannot Batch"
   Trans_Log "No Equipment Found"
   BatchStatus "Failed No Equipment Found", BatchDate
   DbClose Cdb, Tb
   Exit Sub
 End If
 Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
 Set BC = New EMVProcessLib.Process
 Dim djvMsg() As String
 Do While Not Tb.EOF
  
  If LoadEMVDevice(Tb.fields("DeviceName").Value, EmvEquip) = "Failed" Then
    See_Msg "Device Not Found " & Tb.fields("DeviceName").Value, 5, "Check Dejavoo Setup"
    Trans_Log "Device Not Found " & Tb.fields("DeviceName").Value
  Else
   Trans_Log "Searching for " & Tb.fields("PinPadIpAddress").Value & " " & Tb.fields("DeviceName").Value
   If Tb.fields("MerchantId").Value > "" Then
      '// test and make sure pax device has not lost its ip address or has gone to sleep.
      '// will not attempt to batch device.
     Rst = BC.SettleDejavoo("", EmvEquip)
     If InStr(Rst, "Object reference not set") > 0 Then
      ReDim djvMsg(2)
      djvMsg(2) = "Batch Error " & Rst
     Else
      djvMsg = Split(Rst, vbCrLf)
     End If
     If djvMsg(1) = "Batch Error : 95" Or InStr(djvMsg(0), "Batch Approved") > 0 Then
      SSCommand21.Caption = "Batch Approved."
      BatchCompleted "", "Batch " & djvMsg(2), 1
      BatchStatus "AP " & djvMsg(2), BatchDate
     Else
      SSCommand21.Caption = "Batch Failed.."
      EmailAlerts "BATCHFAILED.Alt", "Batch Failed; Message Returned from processor " & djvMsg(0)
      BatchStatus "Fail " & djvMsg(2), BatchDate
     End If
     Trans_Log "Results from Send Batch " & Rst & " Device " & Tb.fields("DeviceName").Value
     BStation = BStation + 1
     BatchCompleted Tb.fields("DeviceName").Value, BRt.Item(1).ResultTxt & " Batch Approved for " & Format(BRt.Item(1).CreditAmt, "0.00"), BStation
     BatchStatus "AP " & BRt.Item(1).ResultTxt & " " & Format(BRt.Item(1).CreditAmt, "0.00"), BatchDate
     Trans_Log "AP " & BRt.Item(1).ResultTxt & " " & Format(BRt.Item(1).CreditAmt, "0.00")
     BRt.Item(1).CreditAmt = 0
   Else
    '// cannot ping device
    Trans_Log "Device Not Found " & Tb.fields("PinPadIpAddress").Value
    See_Msg "Device Not Found " & Tb.fields("DeviceName").Value, 5, "Check Pax Setup"
    'See_Msg "Pax Device Not Found", 5, "Cannot Batch Terminal"
    FailedBatch = True
   End If
  End If
  Tb.MoveNext
 Loop
DbClose Cdb, Tb

If FailedBatch Then
  SSCommand21.Caption = "Batch Failed"
  See_Msg "dejavoo Batch Did Not Complete", 5, "Batch Failed"
  Trans_Log "Batch Send Failed"
Else
 SSCommand21.Caption = "Batch Approved."
 Trans_Log "Batch Send Finished..."
End If
Set BRt = Nothing
Set BC = Nothing
Exit Sub
ErrorHandler:
If Err.Number = 91 Then
  Trans_Log "Error 91 " & batching
  Exit Sub  '// stop batch process with pax machine returning an object error 91
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Batch Pax ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Private Sub BatchPaxDevices(BatchDate As String)
On Error GoTo ErrorHandler
Dim BStation    As Integer
Dim Rst         As String
Dim Cdb         As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim FailedBatch As Boolean
Dim BC          As EMVProcessLib.Process
Dim BRt         As PAXMSG
FailedBatch = False
BStation = 1
ConnectADODB DebitDb, Cdb
    Sql = "Select * from [EMVDeviceList] Where [DeviceName] > ''"
ConnectADORS Sql, Cdb, Tb
 If Tb.EOF Then
   See_Msg "No Device Installed", 5, "Cannot Batch"
   Trans_Log "No Equipment Found"
   BatchStatus "Failed No Equipment Found", BatchDate
   DbClose Cdb, Tb
   Exit Sub
 End If
 Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
 Set BC = New EMVProcessLib.Process
 Set BRt = New PAXMSG
 
 Do While Not Tb.EOF
  
  If LoadEMVDevice(Tb.fields("DeviceName").Value, EmvEquip) = "Failed" Then
    See_Msg "Device Not Found " & Tb.fields("DeviceName").Value, 5, "Check Pax Setup"
    Trans_Log "Device Not Found " & Tb.fields("DeviceName").Value
  Else
   Trans_Log "Searching for " & Tb.fields("PinPadIpAddress").Value & " " & Tb.fields("DeviceName").Value
   If Tb.fields("PinPadIpAddress").Value > "" And IsTermialThere(Tb.fields("PinPadIpAddress").Value) Then
      '// test and make sure pax device has not lost its ip address or has gone to sleep.
      '// will not attempt to batch device.
      Rst = BC.Send_Batch(BRt, EmvEquip)
      
      Trans_Log "Results from Send Batch " & Rst & " Device " & Tb.fields("DeviceName").Value
      If Rst <> "Success" Then
       If BRt Is Nothing Then
        See_Msg "Device Failed To Connect", 5, "Check Pax Device"
        Trans_Log "Device Failed To Connect"
       ElseIf BRt.Item(1) Is Nothing Then
        See_Msg "Device Failed To Connect", 5, "Check Pax Device"
        Trans_Log "Device Failed To Connect"
       End If
       Trans_Log BRt.Item(1).ResultTxt
       BatchStatus "Fail " & BRt.Item(1).ResultTxt, BatchDate
       Call See_Msg("Please Check Pax Termial", 5, BRt.Item(1).ResultTxt)
       FailedBatch = True '// will not clear a failed batch
      Else
       If BRt.Item(1).ResultTxt = "OK" Or BRt.Item(1).ResultTxt = "NOT FOUND" Then
        Trans_Log "Batch Results " & BRt.Item(1).ResultTxt & " Response " & Rst
        If BRt.Item(1).ResultTxt = "NOT FOUND" Then
          See_Msg "Batch Already Closed...", 5, ""
        Else
          See_Msg "Batch " & BRt.Item(1).ResultTxt, 5, "Approved for " & Format(BRt.Item(1).CreditAmt, "0.00")
        End If
        BStation = BStation + 1
        BatchCompleted Tb.fields("DeviceName").Value, BRt.Item(1).ResultTxt & " Batch Approved for " & Format(BRt.Item(1).CreditAmt, "0.00"), BStation
        BatchStatus "AP " & BRt.Item(1).ResultTxt & " " & Format(BRt.Item(1).CreditAmt, "0.00"), BatchDate
        Trans_Log "AP " & BRt.Item(1).ResultTxt & " " & Format(BRt.Item(1).CreditAmt, "0.00")
        BRt.Item(1).CreditAmt = 0
        If Tb.fields("DeviceNumberPadType1").Value & "" > "" Then
         Rst = BC.SetBOA(Tb.fields("DeviceNumberPadType1").Value, EmvEquip)
         Trans_Log "Update MSCN " & Rst
        End If '// this MSCN is used for update after batch complete to bank of america
       Else
        Trans_Log "Results from Send Batch " & BRt.Item(1).ResultTxt & " Device " & Tb.fields("DeviceName").Value
        Call See_Msg("Batch Failed, Check Pax Terminal", 5, BRt.Item(1).ResultTxt)
        FailedBatch = True
       End If
      End If
   Else
    '// cannot ping device
    Trans_Log "Device Not Found " & Tb.fields("PinPadIpAddress").Value
    See_Msg "Device Not Found " & Tb.fields("DeviceName").Value, 5, "Check Pax Setup"
    'See_Msg "Pax Device Not Found", 5, "Cannot Batch Terminal"
    FailedBatch = True
   End If
  End If
  Tb.MoveNext
 Loop
DbClose Cdb, Tb

If FailedBatch Then
  SSCommand21.Caption = "Batch Failed"
  See_Msg "Pax Batch Did Not Complete", 5, "Batch Failed"
  Trans_Log "Batch Send Failed"
Else
 SSCommand21.Caption = "Batch Approved."
 Trans_Log "Batch Send Finished..."
End If
Set BRt = Nothing
Set BC = Nothing
Exit Sub
ErrorHandler:
If Err.Number = 91 Then
  Trans_Log "Error 91 " & batching
  Exit Sub  '// stop batch process with pax machine returning an object error 91
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Batch Pax ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub BatchCharges()
On Error GoTo ErrorHandler
Dim TDb         As ADODB.Connection
Dim Trs         As ADODB.Recordset
Dim BatchDate   As String
Dim RN          As Integer
Dim CHDB        As ADODB.Connection
Dim CHRS        As ADODB.Recordset
Dim Cdb         As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim NotGrat     As Integer
Dim WTime       As Long
Dim Rst         As String
Dim R           As String
Dim result      As String
Dim BC          As EMVProcessLib.Process
Dim ClearBatch  As String
Trans_Log "Starting The Batch Process"
'If Timer - PauseButton < 0.25 Then
' Exit Sub
'End If
'PauseButton = Timer

If AutoCloseInProcess = False Then
   ConnectADODB SvrPath.TerminalPath & CheckStatDb, TDb
     Sql = "Select * from [Check Detail] where [close out day] = 0"
   ConnectADORS Sql, TDb, Trs
    If Not Trs.EOF Then
      If See_Msg("Do you want to continue ?", 4, "Close out daily sales first, Daily Sales will no longer have credit card payments.") <> 6 Then
        DbClose TDb, Trs
        Exit Sub
      End If
    End If
    DbClose TDb, Trs
End If

If GetConfigInfo("Credit Cards", "EMV Clear Records", "No", "FUNCTIONS.INI") = "ENABLED" Then
  ClearBatch = "Yes"
Else
  ClearBatch = "No"
End If

If EMVSUPPORTPREAUTH = 1 Or PAXBARTABPREAUTH = 1 Then ClearBatch = "No"
If AutoCloseInProcess And IsEMVRapidConnect Then GoTo CloseRC

'send to capture tips rapid connect platform
If AutoCloseInProcess Then GoTo ClBatch
If ClearBatch = "Yes" Then GoTo ClBatch


Call See_Msg("Sending Charges To Processor", 5, "Closing Batch for Day")
   
   '// test for batch open
   ListCards.Refresh
   BatchDate = ""
   ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
    Sql = "Select * from [Debit] Where [Tender] = 'EMV' or [Tender] = 'EMV Refund'"
   ConnectADORS Sql, Cdb, Tb
   If Tb.EOF Then
    See_Msg "Batch Already Closed", 5, "No transcation found in batch"
    DbClose Cdb, Tb
    Exit Sub
   Else
    BatchDate = Tb![Date] & ""
   End If
   DbClose Cdb, Tb
'*********************************************************************************
 If IsDate(BatchDate) Then
  ConnectADODB SvrPath.TerminalPath & ClosingHistory, CHDB
  ConnectADORS ("Select * from [Tenders] where TRY_CAST([date] AS datetime) = '" & Format(CDate(BatchDate), "yyyy-mm-dd hh:nn:ss") & "'"), CHDB, CHRS
  If CHRS.EOF Then
    Dim Payments As FinancialItems
    Set Payments = New FinancialItems
    SortCreditCards Payments, 0, "", 0, "", ""
    RN = -1
    For RN = 1 To Payments.count
     If Payments.Item(RN).Tender <> "Cash" Then
       CHRS.AddNew
       CHRS![Date] = Format(BatchDate, "MM/DD/YYYY")
       CHRS![Name] = Payments.Item(RN).Tender
       CHRS![Amount] = Format(Payments.Item(RN).Amount & "", "0.00")
       CHRS![Quantity] = Val(Payments.Item(RN).ItemQty & "")
       CHRS![Tip Total] = Payments.Item(RN).TipAmount
       CHRS.Update
     End If
    Next RN
    Set Payments = Nothing
  End If
  DbClose CHDB, CHRS
 End If
'********************************************************************************
CloseRC:
If AutoCloseInProcess = False Then
   If (IsDataCapInterface(EMVINTERFACE) And EMVSUPPORTPREAUTH = 1) Then
     If IsPadPresent = False Then Exit Sub
     'need pin pad to capture transactions
     'without a pin pad you will not get paid
   End If
   ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
     Sql = "Select * from [Debit] where [close out day] = 0 and [Tender] = 'EMV' and [Tip] = 0"
   ConnectADORS Sql, Cdb, Tb
    NotGrat = RecordsCount(Tb)
    If NotGrat > 0 Then
     If See_Msg("With no gratuity. Ok to continue ??", 4, "There are " & CStr(NotGrat) & " Checks") <> 6 Then
      DbClose Cdb, Tb
      Exit Sub
     End If
    End If
End If
'end of manual close
ClBatch:  'Close batch

If IsDataCapInterface(EMVINTERFACE) Or IsDejavooInterface(EMVINTERFACE) Then
 Dim EmvEquip As EMVProcessLib.EmvDeviceHardWare
 If LoadEMVDevice(MyRegisterName, EmvEquip) = "Failed" Then
   See_Msg "No Device Found", 5, "Check Device Setup"
   Exit Sub
 End If
 Set BC = New EMVProcessLib.Process
End If
'// print batch
Call PrintBatch("Closing Batch", WTime)
SSCommand21.Caption = "Sending....."
Dim TotalRecords As Long
Dim LastCheckNo  As Long
Dim Cr           As Long
DbClose Cdb, Tb
   ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
     Sql = "Select * from [Debit] where [Charge] = -1 and [Tender] = 'EMV' Order by [Register Name],[Check Number] asc "
     '// order by register name and check number improve the speed for pax since it has to capture by device.
     '// makes no difference for datacap or Dejavoo.
   ConnectADORS Sql, Cdb, Tb
   TotalRecords = RecordsCount(Tb)
   If TotalRecords > 0 Then
    Cr = 0
    See_Msg "Sending all preauth charges for capture", 5, "Capture Batch"
    
    Do While Not Tb.EOF
      Cr = Cr + 1
      ItemCount.Caption = "Sending " & Tb.fields("Check Number").Value & " " & Cr & " of " & TotalRecords
      ItemCount.Refresh
      Trans_Log "Sending " & Tb.fields("Check Number").Value & " " & Cr & " of " & TotalRecords
      result = ApplyPAXTip(Tb.fields("Check Number").Value, Tb.fields("Server Name").Value, Format(Tb.fields("Tip").Value, "0.00"), Tb.fields("Account Number").Value, DoEMVCapture, Tb.fields("Card Number").Value, Tb.fields("Account Name").Value & "")
      If result = "Approved" Then
        Tb.fields("Tip Paid").Value = 1
        'On Error Resume Next
        Tb.Update
        Trans_Log "Sending preauth captures " & Tb.fields("Check Number").Value & " captured"
      Else
        Trans_Log "Sending preauth captures failed " & result
        PrintFailedCapture "EMV Capture Failed " & result, Tb.fields("Check Number").Value, Tb.fields("Card Number").Value, Format(Tb.fields("Amount").Value, "0.00"), Format(Tb.fields("Tip").Value, "0.00")
      End If
      Tb.MoveNext
    Loop
   End If
    
   
   If IsPAXInterface(EMVINTERFACE) Then
      GoTo DoBatch
   End If
   
 '  If IsDataCapInterface(EMVINTERFACE) Then
 '    If EMVSUPPORTPREAUTH = 1 Then
 '      Rst = BC.SAFForward(EmvEquip)
 '      Trans_Log "Store and Forward " & Rst
 '      See_Msg "Checking Store and Forward", 5, Rst
 '    End If
 '  End If
   
  If IsEMVRapidConnect Or IsDejavooInterface(EMVINTERFACE) Then
    See_Msg "EMV Batch Capture Complete", 5, "Closed By Processor."
    If EMVSUPPORTPREAUTH = 1 Then
      Sql = "Delete from [Debit] Where [Tip Paid] = 1"
    Else
      Sql = "Delete from [Debit]"
    End If
    Cdb.Execute Sql
     Sql = "Delete from [Debit] Where [Tender] = 'EMV Refund'"
    Cdb.Execute Sql
     Sql = "Delete from [Debit] Where [Amount] = 0"
    Cdb.Execute Sql
    DbClose Cdb, Tb
    SSCommand21.Caption = "Batch Closed."
    See_Msg "EMV Batch Records Closed", 5, "Captured All."
    BatchStatus "All Records Captured...", Format(Date, "mm/dd/yyyy")
    LoadGrid ""
    Set BC = Nothing
    Trans_Log "All Records Captured "
    Exit Sub
  End If
  
DoBatch:  'Close batch

If ClearBatch = "Yes" Then  'IsEMVRapidConnect Or
   ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
   Sql = "Delete from [Debit] Where [Tender] IN ('QRPayment','EMV','EMV Refund');"
   Cdb.Execute Sql
   SetNothing Cdb
   SSCommand21.Caption = "Batch Closed."
   See_Msg "EMV Batch Records Closed", 5, "Closed By Time Batch"
   LoadGrid ""
   SSCommand21.Caption = "Batch Accepted"
   BatchCompleted "", EMVINTERFACE, 1
  Exit Sub
End If

Trans_Log " Sending Batch"
If IsPAXInterface(EMVINTERFACE) Then
   Trans_Log "Sending Pax Batch"
   BatchPaxDevices BatchDate
ElseIf IsDejavooInterface(EMVINTERFACE) Then
    'BatchDejavooDevices BatchDate
    Set BC = New EMVProcessLib.Process
    Rst = BC.SettleDejavoo("", EmvEquip)
    '//////////
    Trans_Log "Batch Data Returned From Dejavoo " & Rst
    
    Dim djvMsg() As String
    If InStr(Rst, "Object reference not set") > 0 Then
     ReDim djvMsg(2)
     djvMsg(2) = "Batch Error " & Rst
    Else
      djvMsg = Split(Rst, vbCrLf)
    End If
    If djvMsg(1) = "Batch Error : 95" Or InStr(djvMsg(0), "Batch Approved") > 0 Then
      SSCommand21.Caption = "Batch Approved."
      BatchCompleted "", "Batch " & djvMsg(2), 1
      BatchStatus "AP " & djvMsg(2), BatchDate
    Else
      SSCommand21.Caption = "Batch Failed.."
      EmailAlerts "BATCHFAILED.Alt", "Batch Failed; Message Returned from processor " & djvMsg(0)
      BatchStatus "Fail " & djvMsg(2), BatchDate
    End If
Else
   If GetConfigInfo("Credit Cards", "Store And Foward Is", "DISABLED", "FUNCTIONS.INI") = "ENABLED" Then
     Rst = BC.SAFForward(EmvEquip)
     See_Msg "Checking Store and Forward", 5, Rst
     Dim OffLineCaptures As Currency
     If xml.token("CmdStatus") = "Error" Then
      OffLineCaptures = 0
     Else
      OffLineCaptures = Val(xml.token("StoredSaleAuthorized"))
     End If
     Trans_Log "Store and Foward Results " & Rst
   End If
   Rst = BC.BatchClose(EmvEquip)
   Trans_Log Rst
   See_Msg Rst, 5, "Batch"
   If InStr(Rst, "Approved") > 0 Then
      SSCommand21.Caption = "Batch Approved."
      BatchCompleted "", "Batch " & Rst, 1
      BatchStatus "AP " & Rst, BatchDate
   Else
      SSCommand21.Caption = "Batch Failed.."
      EmailAlerts "BATCHFAILED.Alt", "Batch Failed; Message Returned from processor " & Rst
      BatchStatus "Fail " & Rst, BatchDate
   End If
End If
Set BC = Nothing
Set BRt = Nothing
If System_config = FullService Then
 FailedBatchRecordReport BatchDate
End If
Trans_Log "Batch Charges Complete Send"
Call UpdateWindow(ListCards.hwnd)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Batch Charges ")
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
CenterObject Me, BatchUtility
'If Dir(App.path & "\MakeMeVisible.dat") > "" Then BtnEnh5.Visible = True
If GetConfigInfo("Misc", "Show Batch Utility Button", "", "FUNCTIONS.INI") = "ENABLED" Then BtnEnh5.Visible = True
Dim Cdb As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & DebitDb, Cdb
  Sql = "Select * from [Debit] Where [Tender] = 'EMV' or [Tender] = 'EMV Refund'"
  ConnectADORS Sql, Cdb, Tb
  If Tb.EOF Then
     Command5.Visible = True
  Else
     Command5.Visible = False
End If
DbClose Cdb, Tb
BatchUtility.Visible = True
Sleep 0
End Sub
Private Sub SSCommand6_Click()
If LoggedOnAs(0).EmpAccessLevel < Access_Required("Credit Card Tips") Then Exit Sub
PrintBatch "Batch is Open", 0
End Sub
Private Sub SSCommand7_Click()
On Error GoTo ERH
If Can_Continue("Clear Batch") = False Then Exit Sub
If AdminIsOn Then
  If GetAdministrator("Clear Batch") = False Then Exit Sub
End If
Dim Db As ADODB.Connection
If See_Msg("Are You Sure ? You are about to clear the batch", 4, "Resets Batch All Data Will Be Lost !!") = 6 Then
  If See_Msg("All Credit Card Will Be Erased. YES to Continue", 4, "Resets Batch All Data Will Be Lost !!") <> 6 Then Exit Sub
  FileCopy SvrPath.TerminalPath & "Debit.Mdb", SvrPath.TerminalPath & "Daily BackUp Files\" & "Clear" & "-Debit.mdb"
  SqlToAccessBackUp SvrPath.TerminalPath & "Daily BackUp Files\" & "Clear" & "-Debit.mdb", "Clear" & "-Debit.mdb", SvrPath.TerminalPath & DebitDb, DebitDb
  Sql = "Delete from [Debit] Where [Tender] IN ('QRPayment','EMV','EMV Refund');"
  ConnectADODB SvrPath.TerminalPath & DebitDb, Db
  Db.Execute Sql
  SetNothing Db
  Trans_Log "Batch Clear"
End If
LoadGrid ""
Exit Sub
ERH:
 See_Msg "Error Clearing Batch " & Err.Description, 0, ""
End Sub
Private Sub BatchCompleted(DeviceName As String, BatchResults As String, BatchNo As Integer)
On Error GoTo ErrorHandler
Dim Wl            As Integer
Dim TXTFileInfo() As Variant
Dim Db            As ADODB.Connection
Dim PrinterPort   As String
Dim Email         As String
        Trans_Log "Batch Closed " & DeviceName & BatchResults
        ConnectADODB SvrPath.TerminalPath & DebitDb, Db
        If IsPAXInterface(EMVINTERFACE) Then
         Sql = "Delete from [Debit] Where [Register Name] = '" & DeviceName & "' and [Tender] IN ('QRPayment','EMV','EMV Refund');"
        Else
         Sql = "Delete from [Debit] Where [Tender] IN ('QRPayment','EMV','EMV Refund');"
        End If
        Db.Execute Sql
        SetNothing Db
        If Check_Option(96) = 0 Then
              If GetConfigInfo("Printing Options", "Disable Batch Chits", "No", "FUNCTIONS.INI") = "DISABLED" Then
               PF = FreeFile
               PrinterPort = "Batch Completed"
               Open PrinterPort For Output As PF
               Print #PF, EC(2); "POS Batch Total  " & Format(BatchTotal, "0.00") & EC(7)
               Print #PF, EC(2); "Please Check Batch Totals" & EC(7)
               Print #PF, EC(2); "Batch Close Time "; Format(Time, "HH:MM ampm") & EC(7)
               Print #PF, EC(2); "Batch Close Date "; Date & EC(7)
               Print #PF, EC(2); "Returned From Processor"; EC(7)
               Print #PF, EC(2); DeviceName & " " & BatchResults & EC(7)
               Print #PF, EC(6); EC(5)
               Close
               SendJobToPrinter PrinterPort, "Receipt", False
              End If
              If SendEmails(6).Send Then
                SetPrinter 0
                PrintLine "POS Batch Total        " & Format(BatchTotal, "0.00"), -1
                PrintLine "Please Check Batch Total", -1
                PrintLine "Batch Close Time       " & Format(Time, "HH:MM ampm"), -1
                PrintLine "Batch Close Date       " & Format(Date, "mm/dd/yyyy"), -1
                PrintLine "Batch Completed " & DeviceName & " " & BatchResults, -1
                ReportView.View.EndDoc
                ReportView.EmailEOD CStr(BatchNo) & " Batch Results"
              End If
        Else
              SetPrinter 0
              PrintLine "POS Batch Total        " & Format(BatchTotal, "0.00"), -1
              PrintLine "Please Check Batch Total", -1
              PrintLine "Batch Close Time       " & Format(Time, "HH:MM ampm"), -1
              PrintLine "Batch Close Date       " & Format(Date, "mm/dd/yyyy"), -1
              PrintLine "Batch Completed " & DeviceName & " " & BatchResults, -1
              ReportView.View.EndDoc
              If SendEmails(6).Send Then
                Trans_Log "Sending Batch Chit"
                ReportView.EmailEOD CStr(BatchNo) & " Batch Results"
              End If
              If Check_Option(130) = 0 Then ReportView.View.PrintDoc
        End If
        SSCommand21.Caption = "Batch Closed."
        See_Msg "Batch Approved...", 5, DeviceName & " " & BatchResults
        LoadGrid ""
      Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Batch Complete ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub


