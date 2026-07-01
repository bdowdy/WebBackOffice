VERSION 5.00
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Time_Reports 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Time Records"
   ClientHeight    =   10740
   ClientLeft      =   15
   ClientTop       =   420
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Reportday 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   11880
      ScaleHeight     =   5385
      ScaleWidth      =   3150
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   3180
      Begin VB.ListBox DateList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin BTNENHLib4.BtnEnh SSCommand5 
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   4560
         Width           =   2415
         _Version        =   589828
         _ExtentX        =   4260
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
         textCaption     =   "Time_Reports.frx":0000
         textLT          =   "Time_Reports.frx":0068
         textCT          =   "Time_Reports.frx":0080
         textRT          =   "Time_Reports.frx":0098
         textLM          =   "Time_Reports.frx":00B0
         textRM          =   "Time_Reports.frx":00C8
         textLB          =   "Time_Reports.frx":00E0
         textCB          =   "Time_Reports.frx":00F8
         textRB          =   "Time_Reports.frx":0110
         colorBack       =   "Time_Reports.frx":0128
         colorIntern     =   "Time_Reports.frx":0152
         colorMO         =   "Time_Reports.frx":017C
         colorFocus      =   "Time_Reports.frx":01A6
         colorDisabled   =   "Time_Reports.frx":01D0
         colorPressed    =   "Time_Reports.frx":01FA
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Available Dates"
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
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   4335
         Left            =   120
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox EList 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   -3000
      ScaleHeight     =   8025
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   -120
      Visible         =   0   'False
      Width           =   4530
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   960
         TabIndex        =   9
         Top             =   6360
         Width           =   2415
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5070
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   720
         Width           =   3735
      End
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   975
         Left            =   3000
         TabIndex        =   6
         Top             =   6960
         Width           =   1335
         _Version        =   589828
         _ExtentX        =   2355
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
         textCaption     =   "Time_Reports.frx":0224
         textLT          =   "Time_Reports.frx":028C
         textCT          =   "Time_Reports.frx":02A4
         textRT          =   "Time_Reports.frx":02BC
         textLM          =   "Time_Reports.frx":02D4
         textRM          =   "Time_Reports.frx":02EC
         textLB          =   "Time_Reports.frx":0304
         textCB          =   "Time_Reports.frx":031C
         textRB          =   "Time_Reports.frx":0334
         colorBack       =   "Time_Reports.frx":034C
         colorIntern     =   "Time_Reports.frx":0376
         colorMO         =   "Time_Reports.frx":03A0
         colorFocus      =   "Time_Reports.frx":03CA
         colorDisabled   =   "Time_Reports.frx":03F4
         colorPressed    =   "Time_Reports.frx":041E
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
      Begin BTNENHLib4.BtnEnh BtnEnh5 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   6960
         Width           =   2775
         _Version        =   589828
         _ExtentX        =   4895
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Print Individual Time Card"
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
         textCaption     =   "Time_Reports.frx":0448
         textLT          =   "Time_Reports.frx":04DC
         textCT          =   "Time_Reports.frx":04F4
         textRT          =   "Time_Reports.frx":050C
         textLM          =   "Time_Reports.frx":0524
         textRM          =   "Time_Reports.frx":053C
         textLB          =   "Time_Reports.frx":0554
         textCB          =   "Time_Reports.frx":056C
         textRB          =   "Time_Reports.frx":0584
         colorBack       =   "Time_Reports.frx":059C
         colorIntern     =   "Time_Reports.frx":05C6
         colorMO         =   "Time_Reports.frx":05F0
         colorFocus      =   "Time_Reports.frx":061A
         colorDisabled   =   "Time_Reports.frx":0644
         colorPressed    =   "Time_Reports.frx":066E
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Employee List"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4215
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   6735
         Left            =   120
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Labor Reports"
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
      Height          =   8295
      Left            =   3480
      TabIndex        =   10
      Top             =   600
      Width           =   7455
      Begin BTNENHLib4.BtnEnh SSCommand4 
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Punch Report"
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
         textCaption     =   "Time_Reports.frx":0698
         textLT          =   "Time_Reports.frx":0710
         textCT          =   "Time_Reports.frx":0728
         textRT          =   "Time_Reports.frx":0740
         textLM          =   "Time_Reports.frx":0758
         textRM          =   "Time_Reports.frx":0770
         textLB          =   "Time_Reports.frx":0788
         textCB          =   "Time_Reports.frx":07A0
         textRB          =   "Time_Reports.frx":07B8
         colorBack       =   "Time_Reports.frx":07D0
         colorIntern     =   "Time_Reports.frx":07FA
         colorMO         =   "Time_Reports.frx":0824
         colorFocus      =   "Time_Reports.frx":084E
         colorDisabled   =   "Time_Reports.frx":0878
         colorPressed    =   "Time_Reports.frx":08A2
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
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Hourly Labor Report"
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
         textCaption     =   "Time_Reports.frx":08CC
         textLT          =   "Time_Reports.frx":0952
         textCT          =   "Time_Reports.frx":096A
         textRT          =   "Time_Reports.frx":0982
         textLM          =   "Time_Reports.frx":099A
         textRM          =   "Time_Reports.frx":09B2
         textLB          =   "Time_Reports.frx":09CA
         textCB          =   "Time_Reports.frx":09E2
         textRB          =   "Time_Reports.frx":09FA
         colorBack       =   "Time_Reports.frx":0A12
         colorIntern     =   "Time_Reports.frx":0A3C
         colorMO         =   "Time_Reports.frx":0A66
         colorFocus      =   "Time_Reports.frx":0A90
         colorDisabled   =   "Time_Reports.frx":0ABA
         colorPressed    =   "Time_Reports.frx":0AE4
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
         Height          =   840
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   3480
         _Version        =   589828
         _ExtentX        =   6138
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Employee Payroll Summary / Labor Report"
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
         PictureCustom   =   "C:\Users\ben\Downloads\flat_button_no_text.png"
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPictureCustom=   "Time_Reports.frx":0B0E
         textCaption     =   "Time_Reports.frx":0BBC
         textLT          =   "Time_Reports.frx":0C6A
         textCT          =   "Time_Reports.frx":0C82
         textRT          =   "Time_Reports.frx":0C9A
         textLM          =   "Time_Reports.frx":0CB2
         textRM          =   "Time_Reports.frx":0CCA
         textLB          =   "Time_Reports.frx":0CE2
         textCB          =   "Time_Reports.frx":0CFA
         textRB          =   "Time_Reports.frx":0D12
         colorBack       =   "Time_Reports.frx":0D2A
         colorIntern     =   "Time_Reports.frx":0D54
         colorMO         =   "Time_Reports.frx":0D7E
         colorFocus      =   "Time_Reports.frx":0DA8
         colorDisabled   =   "Time_Reports.frx":0DD2
         colorPressed    =   "Time_Reports.frx":0DFC
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
         Left            =   3720
         TabIndex        =   14
         Top             =   2280
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Who's On the Clock"
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
         textCaption     =   "Time_Reports.frx":0E26
         textLT          =   "Time_Reports.frx":0EAA
         textCT          =   "Time_Reports.frx":0EC2
         textRT          =   "Time_Reports.frx":0EDA
         textLM          =   "Time_Reports.frx":0EF2
         textRM          =   "Time_Reports.frx":0F0A
         textLB          =   "Time_Reports.frx":0F22
         textCB          =   "Time_Reports.frx":0F3A
         textRB          =   "Time_Reports.frx":0F52
         colorBack       =   "Time_Reports.frx":0F6A
         colorIntern     =   "Time_Reports.frx":0F94
         colorMO         =   "Time_Reports.frx":0FBE
         colorFocus      =   "Time_Reports.frx":0FE8
         colorDisabled   =   "Time_Reports.frx":1012
         colorPressed    =   "Time_Reports.frx":103C
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
         Left            =   3720
         TabIndex        =   15
         Top             =   1320
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Over Time Watch"
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
         textCaption     =   "Time_Reports.frx":1066
         textLT          =   "Time_Reports.frx":10E4
         textCT          =   "Time_Reports.frx":10FC
         textRT          =   "Time_Reports.frx":1114
         textLM          =   "Time_Reports.frx":112C
         textRM          =   "Time_Reports.frx":1144
         textLB          =   "Time_Reports.frx":115C
         textCB          =   "Time_Reports.frx":1174
         textRB          =   "Time_Reports.frx":118C
         colorBack       =   "Time_Reports.frx":11A4
         colorIntern     =   "Time_Reports.frx":11CE
         colorMO         =   "Time_Reports.frx":11F8
         colorFocus      =   "Time_Reports.frx":1222
         colorDisabled   =   "Time_Reports.frx":124C
         colorPressed    =   "Time_Reports.frx":1276
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
      Begin BTNENHLib4.BtnEnh SSCommand11 
         Height          =   855
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Break Report"
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
         textCaption     =   "Time_Reports.frx":12A0
         textLT          =   "Time_Reports.frx":1318
         textCT          =   "Time_Reports.frx":1330
         textRT          =   "Time_Reports.frx":1348
         textLM          =   "Time_Reports.frx":1360
         textRM          =   "Time_Reports.frx":1378
         textLB          =   "Time_Reports.frx":1390
         textCB          =   "Time_Reports.frx":13A8
         textRB          =   "Time_Reports.frx":13C0
         colorBack       =   "Time_Reports.frx":13D8
         colorIntern     =   "Time_Reports.frx":1402
         colorMO         =   "Time_Reports.frx":142C
         colorFocus      =   "Time_Reports.frx":1456
         colorDisabled   =   "Time_Reports.frx":1480
         colorPressed    =   "Time_Reports.frx":14AA
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
         Left            =   3720
         TabIndex        =   17
         Top             =   4200
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Phone List"
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
         textCaption     =   "Time_Reports.frx":14D4
         textLT          =   "Time_Reports.frx":1554
         textCT          =   "Time_Reports.frx":156C
         textRT          =   "Time_Reports.frx":1584
         textLM          =   "Time_Reports.frx":159C
         textRM          =   "Time_Reports.frx":15B4
         textLB          =   "Time_Reports.frx":15CC
         textCB          =   "Time_Reports.frx":15E4
         textRB          =   "Time_Reports.frx":15FC
         colorBack       =   "Time_Reports.frx":1614
         colorIntern     =   "Time_Reports.frx":163E
         colorMO         =   "Time_Reports.frx":1668
         colorFocus      =   "Time_Reports.frx":1692
         colorDisabled   =   "Time_Reports.frx":16BC
         colorPressed    =   "Time_Reports.frx":16E6
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
         Left            =   3720
         TabIndex        =   18
         Top             =   3240
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Access Codes"
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
         textCaption     =   "Time_Reports.frx":1710
         textLT          =   "Time_Reports.frx":1794
         textCT          =   "Time_Reports.frx":17AC
         textRT          =   "Time_Reports.frx":17C4
         textLM          =   "Time_Reports.frx":17DC
         textRM          =   "Time_Reports.frx":17F4
         textLB          =   "Time_Reports.frx":180C
         textCB          =   "Time_Reports.frx":1824
         textRB          =   "Time_Reports.frx":183C
         colorBack       =   "Time_Reports.frx":1854
         colorIntern     =   "Time_Reports.frx":187E
         colorMO         =   "Time_Reports.frx":18A8
         colorFocus      =   "Time_Reports.frx":18D2
         colorDisabled   =   "Time_Reports.frx":18FC
         colorPressed    =   "Time_Reports.frx":1926
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
      Begin BTNENHLib4.BtnEnh SSCommand8 
         Height          =   855
         Left            =   3720
         TabIndex        =   19
         Top             =   6120
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Individual Time Card Review"
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
         textCaption     =   "Time_Reports.frx":1950
         textLT          =   "Time_Reports.frx":19E6
         textCT          =   "Time_Reports.frx":19FE
         textRT          =   "Time_Reports.frx":1A16
         textLM          =   "Time_Reports.frx":1A2E
         textRM          =   "Time_Reports.frx":1A46
         textLB          =   "Time_Reports.frx":1A5E
         textCB          =   "Time_Reports.frx":1A76
         textRB          =   "Time_Reports.frx":1A8E
         colorBack       =   "Time_Reports.frx":1AA6
         colorIntern     =   "Time_Reports.frx":1AD0
         colorMO         =   "Time_Reports.frx":1AFA
         colorFocus      =   "Time_Reports.frx":1B24
         colorDisabled   =   "Time_Reports.frx":1B4E
         colorPressed    =   "Time_Reports.frx":1B78
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
      Begin BTNENHLib4.BtnEnh SSCommand9 
         Height          =   855
         Left            =   3720
         TabIndex        =   20
         Top             =   5160
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Time Records Report"
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
         textCaption     =   "Time_Reports.frx":1BA2
         textLT          =   "Time_Reports.frx":1C34
         textCT          =   "Time_Reports.frx":1C4C
         textRT          =   "Time_Reports.frx":1C64
         textLM          =   "Time_Reports.frx":1C7C
         textRM          =   "Time_Reports.frx":1C94
         textLB          =   "Time_Reports.frx":1CAC
         textCB          =   "Time_Reports.frx":1CC4
         textRB          =   "Time_Reports.frx":1CDC
         colorBack       =   "Time_Reports.frx":1CF4
         colorIntern     =   "Time_Reports.frx":1D1E
         colorMO         =   "Time_Reports.frx":1D48
         colorFocus      =   "Time_Reports.frx":1D72
         colorDisabled   =   "Time_Reports.frx":1D9C
         colorPressed    =   "Time_Reports.frx":1DC6
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
      Begin BTNENHLib4.BtnEnh SSCommand12 
         Height          =   855
         Left            =   240
         TabIndex        =   21
         Top             =   6120
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Individual Daily Time and Sales Report"
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
         textCaption     =   "Time_Reports.frx":1DF0
         textLT          =   "Time_Reports.frx":1E9C
         textCT          =   "Time_Reports.frx":1EB4
         textRT          =   "Time_Reports.frx":1ECC
         textLM          =   "Time_Reports.frx":1EE4
         textRM          =   "Time_Reports.frx":1EFC
         textLB          =   "Time_Reports.frx":1F14
         textCB          =   "Time_Reports.frx":1F2C
         textRB          =   "Time_Reports.frx":1F44
         colorBack       =   "Time_Reports.frx":1F5C
         colorIntern     =   "Time_Reports.frx":1F86
         colorMO         =   "Time_Reports.frx":1FB0
         colorFocus      =   "Time_Reports.frx":1FDA
         colorDisabled   =   "Time_Reports.frx":2004
         colorPressed    =   "Time_Reports.frx":202E
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
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Hours By Department"
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
         textCaption     =   "Time_Reports.frx":2058
         textLT          =   "Time_Reports.frx":20DE
         textCT          =   "Time_Reports.frx":20F6
         textRT          =   "Time_Reports.frx":210E
         textLM          =   "Time_Reports.frx":2126
         textRM          =   "Time_Reports.frx":213E
         textLB          =   "Time_Reports.frx":2156
         textCB          =   "Time_Reports.frx":216E
         textRB          =   "Time_Reports.frx":2186
         colorBack       =   "Time_Reports.frx":219E
         colorIntern     =   "Time_Reports.frx":21C8
         colorMO         =   "Time_Reports.frx":21F2
         colorFocus      =   "Time_Reports.frx":221C
         colorDisabled   =   "Time_Reports.frx":2246
         colorPressed    =   "Time_Reports.frx":2270
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
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   855
         Left            =   240
         TabIndex        =   23
         Top             =   5160
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Average Hourly Rate Exception Report for Tipped Employees"
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
         textCaption     =   "Time_Reports.frx":229A
         textLT          =   "Time_Reports.frx":236C
         textCT          =   "Time_Reports.frx":2384
         textRT          =   "Time_Reports.frx":239C
         textLM          =   "Time_Reports.frx":23B4
         textRM          =   "Time_Reports.frx":23CC
         textLB          =   "Time_Reports.frx":23E4
         textCB          =   "Time_Reports.frx":23FC
         textRB          =   "Time_Reports.frx":2414
         colorBack       =   "Time_Reports.frx":242C
         colorIntern     =   "Time_Reports.frx":2456
         colorMO         =   "Time_Reports.frx":2480
         colorFocus      =   "Time_Reports.frx":24AA
         colorDisabled   =   "Time_Reports.frx":24D4
         colorPressed    =   "Time_Reports.frx":24FE
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
         Height          =   855
         Left            =   1920
         TabIndex        =   24
         Top             =   7200
         Width           =   3495
         _Version        =   589828
         _ExtentX        =   6165
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
         textCaption     =   "Time_Reports.frx":2528
         textLT          =   "Time_Reports.frx":2590
         textCT          =   "Time_Reports.frx":25A8
         textRT          =   "Time_Reports.frx":25C0
         textLM          =   "Time_Reports.frx":25D8
         textRM          =   "Time_Reports.frx":25F0
         textLB          =   "Time_Reports.frx":2608
         textCB          =   "Time_Reports.frx":2620
         textRB          =   "Time_Reports.frx":2638
         colorBack       =   "Time_Reports.frx":2650
         colorIntern     =   "Time_Reports.frx":267A
         colorMO         =   "Time_Reports.frx":26A4
         colorFocus      =   "Time_Reports.frx":26CE
         colorDisabled   =   "Time_Reports.frx":26F8
         colorPressed    =   "Time_Reports.frx":2722
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
         Office2003ColorBorder=   12632256
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
      Begin BTNENHLib4.BtnEnh SSCommand2 
         Height          =   840
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   3480
         _Version        =   589828
         _ExtentX        =   6138
         _ExtentY        =   1482
         _StockProps     =   66
         Caption         =   "Pay Period Labor"
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
         PictureCustom   =   "C:\Users\ben\Downloads\flat_button_no_text.png"
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPictureCustom=   "Time_Reports.frx":274C
         textCaption     =   "Time_Reports.frx":27FA
         textLT          =   "Time_Reports.frx":287A
         textCT          =   "Time_Reports.frx":2892
         textRT          =   "Time_Reports.frx":28AA
         textLM          =   "Time_Reports.frx":28C2
         textRM          =   "Time_Reports.frx":28DA
         textLB          =   "Time_Reports.frx":28F2
         textCB          =   "Time_Reports.frx":290A
         textRB          =   "Time_Reports.frx":2922
         colorBack       =   "Time_Reports.frx":293A
         colorIntern     =   "Time_Reports.frx":2964
         colorMO         =   "Time_Reports.frx":298E
         colorFocus      =   "Time_Reports.frx":29B8
         colorDisabled   =   "Time_Reports.frx":29E2
         colorPressed    =   "Time_Reports.frx":2A0C
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
Attribute VB_Name = "Time_Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectReport As String
Private Type ALLEMPTime
    FirstName As String
    LastName As String
    MiddleName As String
    WholeName  As String
End Type

Private Enum PayrollClose
    ReprintAll
    PrintOne
    CloseOutAll
    Review
    PrintSelected
End Enum
Private Type EmpTimeRecords
    Name As String
    Job As String
    ShiftDate As String
    EndShiftDate As String
    TimeON As String
    TOnAlt As Integer
    BkOut  As String
    BOAlt  As Integer
    BkIn   As String
    BIAlt  As Integer
    TimeOut As String
    TOntAlt As Integer
    MGR     As String
    CashTip As Currency
    CCTip   As Currency
    Grat    As Currency
    TipShare As Currency
    TipOut   As Currency
    PayRate  As Currency
    RegHrs   As Single
    OTHrs    As Single
    DayOt    As Single
    GrossPay As Currency
    PaidBy   As Integer
    Sales    As Currency
    PayWeek  As Integer
    ShiftNo  As Long
    EMPID    As String
    EMPAP    As String
    OTPay    As Currency
    OTRate As Currency
    DayOTOn  As Integer
    RecordId As String
End Type
Private Type BreakReport
    TimeOut As String
    TimeIn  As String
    RecordId As String
    Name As String
End Type
Private Break               As AllHours
Private Display              As Integer
Private CHRS As AllHours
Private DoTimeReport        As Boolean
Private PayRollSummary      As AllHours
Private Total_Time          As Single
Private Total_depts         As Integer
Private LoadPRDates         As String
Private PayRollTo           As String

Public Function WriteHotSchedules(CPANYID As String, CCEPTID As String, LOCID As String, FileName As String) As String
Set CHRS = Nothing
Set CHRS = New AllHours
Set WeeksWorked = New AllHours
Set WeekS = New AllHours
Set PayRollSummary = New AllHours
Set EmpTotalHrs = New AllHours
Dim PayRollData As ServerPayRollData
Set DayOvrTime = New AllHours
Set ROT = New AllHours
Dim BeginDate As String
Dim EndDate   As String
Dim Db        As ADODB.Connection
Dim RS        As ADODB.Recordset
Dim ff        As Integer
Dim TimeRcd()           As EmpTimeRecords
Dim Employees()         As EmpTimeRecords
Dim EmPNames()          As ALLEMPTime
Dim EMPNme()            As String
Dim RMins               As Long
Dim OTMins              As Long
If ConnectADODB(SvrPath.TerminalPath & TimeRecordsDb, Db) = False Then
   WriteHotSchedules = "Error Connecting to Time Records"
   Exit Function
End If
Sql = "Select * from [Employee Records] "
Sql = Sql & " Where TRY_CAST([Shift Date] AS datetime) >= '" & Date - Val(Input_Option(25)) & "'"
Sql = Sql & " Order By TRY_CAST([Shift Date] AS datetime) asc"
If ConnectADORS(Sql, Db, RS) = False Then
   WriteHotSchedules = "Error Connecting to Time Records"
   Exit Function
End If
If Not RS.EOF Then
   If Not RS.BOF Then RS.MoveFirst
   BeginDate = Format(RS![Shift date], "MM/DD/YYYY")
   If Not RS.EOF Then RS.MoveLast
   EndDate = Format(RS![Shift date], "MM/DD/YYYY")
End If
DbClose Db, RS
If Not IsDate(BeginDate) Then Exit Function
If GetEmpLastName(BeginDate, EndDate, EmPNames) = False Then
   See_Msg "No Records Found", 5, ""
   Exit Function
End If
If RunPayroll_Cal(TimeRcd(), Employees(), BeginDate, EndDate, EmPNames, "Employee Records") = False Then Exit Function
ff = FreeFile
Open SvrPath.TerminalPath & "Hot Schedules\" & "timecard" & FileName & ".TXT" For Output As ff
On Error Resume Next
For CT = 0 To UBound(TimeRcd, 1) - 1
     With TimeRcd(CT)
          EMPNme = Split(.Name, " ")
          EMPID = GetEmpId(EMPNme(0), EMPNme(1))
          OTMins = 0
          If .RegHrs > 0 Then
            RMins = .RegHrs * 60
          End If
          If .OTHrs > 0 Then
            OTMins = .OTHrs * 60
          End If
          Print #ff, CPANYID & "|" & CCEPTID & "|" & LOCID & "|" & EMPID & "|" & ReturnJobCode(.Job) & "|" & _
          Format(.ShiftDate, "YYYY-MM-DD hh:mm:ss.000") & "|" & Format(.ShiftDate, "YYYY-MM-DD") & " " & Format(.TimeON, "HH:MM:SS.000") & "|" & _
          Format(.EndShiftDate, "YYYY-MM-DD") & " " & Format(.TimeOut, "HH:MM:SS.000") & "|" & _
          RMins & "|" & OTMins & "|" & Format(.RegHrs * .PayRate, "0.00") & "|" & Format(.OTHrs * (.PayRate * 1.5), "0.00") & "|0|" & Format(.CashTip + .CCTip, "0.00") & "|0|" & _
          Format(.PayRate, "0.00") & "|0.00|"
       End With
   Next CT
Close ff
Set CHRS = Nothing
Set WeeksWorked = Nothing
Set WeekS = Nothing
Set PayRollSummary = Nothing
Set EmpTotalHrs = Nothing
Set DayOvrTime = Nothing
Set ROT = Nothing
End Function
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
If Can_Continue(BtnEnh1.Caption) = False Then Exit Sub
Dim BeginWeek As String
Dim EndingDate As String
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
 Exit Sub
End If
BeginWeek = SearchDates.Item(1).StartDate
EndingDate = SearchDates.Item(1).EndDate
DepartmentIndReport BeginWeek, EndingDate
End Sub
Private Sub BtnEnh2_Click()
If Can_Continue(BtnEnh2.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim PrinterPort As String, Heading As String
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim FailureMsg As String
If Check_Option(96) = 0 Then
PrinterPort = GetSpoolPrinter(AllPrinter.Item("Receipt").PrinterName)
If RemotePrinterFailed(AllPrinter.Item("Receipt").PrinterName, PrinterPort, AllPrinter.Item("Receipt").Failed, FailureMsg) = True Then Exit Sub
CenterText = EC(8)
LeftText = EC(9)
EC(7) = vbCrLf
OneMoreTime:
End If
Select Case Val(Input_Option(30))
  Case 32
     ALine = 0
  Case Else
     ALine = 16
End Select
Dim PrinterData$
Dim Data As String
Heading = "Employee Phone List"
If Check_Option(96) = 0 Then
PrinterData$ = EC(1) & EC(4) & CenterText & Heading & EC(2) & EC(7)
PrinterData$ = PrinterData$ & EC(4) & LeftText & EC(7)
PrinterData$ = PrinterData$ & EC(7)
PrinterData$ = PrinterData$ & Format(Time, "HH:MM AmPm")
PrinterData$ = PrinterData$ & String(TabOver(Format(Time, "HH:MM ampm"), 20), 32) & "Date  " & Trim(Date) & EC(7)
PrinterData$ = PrinterData$ & String(Val(Input_Option(30)) + ALine, 96) & EC(7)
Else
SetPage ReportView.View, Heading
Data = ""
End If

OpenDb "MSaleData", Db
  Sql = "Select * from [General Data] Where ISNULL([Terminated], 0) = 0  order by [Last Name] asc"
ConnectADORS Sql, Db, Tb

Do While Not Tb.EOF
If Tb![home phone] & "" > "" Or Tb![Phone # 2] & "" > "" Then
If Check_Option(96) = 0 Then
  PrinterData$ = PrinterData$ & Trim(Tb(0)) & " " & Trim(Tb(2)) & String(TabOver(Trim(Tb(0)) & " " & Trim(Tb(2)), 25), 32) & Trim(Tb![home phone]) & " Cell " & Tb![Phone # 2] & "" & EC(7)
Else
    Data = Data & Tb(0) & " " & Tb(2) & "|" & Tb![home phone] & "|" & Tb![Phone # 2] & "" & ";"
End If
End If
Tb.MoveNext
Loop
DbClose Db, Tb
If Check_Option(96) = 0 Then
    PrinterData$ = PrinterData$ & String(Val(Input_Option(30)) + ALine, 96) & EC(7)
    PrinterData$ = PrinterData$ & EC(6) & EC(5)
    ReceiptChit PrinterData$, PrinterPort
Else
SetReport Data, SetReportFormat("<3500|>2500|>3000;"), SetReportHeader("Name|Home|Cell;"), tbRows
FinishJob 1
End If
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS8 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh3_Click()
If Can_Continue(BtnEnh3.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim PrinterPort As String, Heading As String
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim FailureMsg As String
If Check_Option(96) = 0 Then
PrinterPort = GetSpoolPrinter(AllPrinter.Item("Receipt").PrinterName)
If RemotePrinterFailed(AllPrinter.Item("Receipt").PrinterName, PrinterPort, AllPrinter.Item("Receipt").Failed, FailureMsg) = True Then Exit Sub
CenterText = EC(8)
LeftText = EC(9)
'RightText = Chr(27) + Chr(97) + Chr(2)
EC(7) = vbCrLf
OneMoreTime:
End If
Select Case Val(Input_Option(30))
  Case 32
     ALine = 0
  Case Else
     ALine = 16
End Select
Dim PrinterData$
Heading = "Employee Access Level List"
If Check_Option(96) = 0 Then
PrinterData$ = EC(1) & EC(4) & CenterText & Heading & EC(2) & EC(7)
PrinterData$ = PrinterData$ & EC(4) & LeftText & EC(7)
PrinterData$ = PrinterData$ & EC(7)
PrinterData$ = PrinterData$ & Format(Time, "HH:MM AmPm")
PrinterData$ = PrinterData$ & String(TabOver(Format(Time, "HH:MM ampm"), 20), 32) & "Date  " & Trim(Date) & EC(7)
PrinterData$ = PrinterData$ & String(Val(Input_Option(30)) + ALine, 96) & EC(7)
Else
SetPage ReportView.View, Heading
End If
OpenDb "MSaleData", Db
  Sql = "Select * from [General Data] Where ISNULL([Terminated], 0) = 0  order by [Last Name] asc"
ConnectADORS Sql, Db, Tb

Do While Not Tb.EOF
If Access_Required(Tb![Access Level]) < 3 Then
    If Check_Option(96) = 0 Then
        PrinterData$ = PrinterData$ & Trim(Tb(0)) & " " & Trim(Tb(2)) & String(TabOver(Trim(Tb(0)) & " " & Trim(Tb(2)), 25), 32) & Trim(Tb![Access Number]) & EC(7)
    Else
        PrintLine Trim(Tb(0)) & " " & Trim(Tb(2)), 8500
        PrintLine Trim(Tb![Access Number]), -1
    End If
End If
Tb.MoveNext
Loop
If Check_Option(96) = 0 Then
PrinterData$ = PrinterData$ & String(Val(Input_Option(30)) + ALine, 96) & EC(7)
PrinterData$ = PrinterData$ & EC(6) & EC(5)
ReceiptChit PrinterData$, PrinterPort
Else
DrawLine
ReportView.View.EndDoc
ReportView.Show 1
End If
DbClose Db, Tb
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Reports SS8 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub BtnEnh4_Click()
EList.Visible = False
EList.Left = -5000
End Sub
Private Sub BtnEnh5_Click()
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
Dim BeginWeek As String
Dim EndingDate As String
If SearchDates Is Nothing Then
     Exit Sub
Else
   BeginWeek = SearchDates.Item(1).StartDate
   EndingDate = SearchDates.Item(1).EndDate
    If Format(BeginWeek, "dddd") <> Trim(Input_Option(6)) Then
       If See_Msg("Payroll Start Date is Different than Begin Payroll Setup. Continue ?", 4, "Payroll Begins On " & Trim(Input_Option(6))) <> 6 Then Exit Sub
    End If
    If DateDiff("d", BeginWeek, EndingDate) < 6 Then
      EndingDate = CStr(CDate(SearchDates.Item(1).StartDate) + 6)
      See_Msg "Cannot print time record for less than 1 than a week", 0, "Ending Date adjusted to " & EndingDate
    End If
End If
PrintPayRoll False, BeginWeek, EndingDate, False, "Individual Time Records"
'Close_Time_Records SvrPath.TerminalPath & "Temp Time File.mdb", ReprintAll, BeginWeek, EndingDate, ""
End Sub

Private Sub BtnEnh6_Click()
End Sub
Private Sub BtnEnh7_Click()
AverageRateofPay
End Sub
Private Sub BtnEnh8_Click()
End Sub
Private Sub BtnEnh9_Click()
If Can_Continue(BtnEnh9.Caption) = False Then Exit Sub
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
Dim BeginWeek As String
Dim EndingDate As String
If SearchDates Is Nothing Then
     Exit Sub
Else
   BeginWeek = SearchDates.Item(1).StartDate
   EndingDate = SearchDates.Item(1).EndDate
End If
SyncPayrollData BeginWeek, EndingDate
PrintPayRollsummary BeginWeek, EndingDate
End Sub
Private Sub Command1_Click()
Dim L As Integer
For L = 0 To List2.ListCount - 1
 If List2.Selected(L) = False Then
  List2.Selected(L) = True
 Else
  List2.Selected(L) = False
 End If
Next L
End Sub
Private Sub DateList_Click()
If Trim(DateList.List(DateList.ListIndex)) = "" Then Exit Sub
If SelectReport = "Hourly Labor" Then
  HourlyLaborData DateList.List(DateList.ListIndex)
End If
DoTimeReport = False
End Sub
Public Sub HouseLabor(BeginDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim TRDB                As ADODB.Connection
Dim TRRS                As ADODB.Recordset
Dim CT                  As Long
Dim RCT                 As Long
Dim PD                  As Long
Dim Job_Code            As String
Dim Emp_Name            As String
Dim EMP_Rate            As Currency
Dim PageHeader          As String
Dim TimeRcd()           As EmpTimeRecords
Dim Employees()         As EmpTimeRecords
Dim ShiftsWorked        As Integer
Dim CurrentWeek         As Integer
Dim WeekS               As AllHours
Dim WeeksWorked         As AllHours
'Dim PayRollSummary      As AllHours
Dim EmpTotalHrs         As AllHours
Dim DayOvrTime          As AllHours
Dim ROT                 As AllHours

Dim PrinterPort         As String
Dim TimeEdits(0 To 3)   As String
Dim ShiftTips           As Currency
Dim TipSales            As Currency
Dim DayOtStr            As String
Dim EmpTotalRegHrs      As Single
Dim EmpTotalOtHrs       As Single
Dim ExportFileTo        As String
Dim NetLaborSales       As Currency
Dim TotalRegHrs         As Single
Dim TotalOtHrs          As Single
Dim Dept_Hrs            As Single
Dim Dept_Ot             As Single
Dim SplitShift          As String
Dim SplitShiftNo        As Long
Dim ShiftCredit         As Currency
Dim FindEmp             As Integer
Dim OTCal               As Single
Dim TotalDayOtHrs       As Single
Dim RegCal              As Single
Dim InOutIsSame         As Boolean
Dim EmPNames()          As ALLEMPTime
Dim Heading             As String
Dim RcdFound            As Boolean
Dim PntName             As Boolean
Dim SumOnly             As Integer
Heading = "Labor Information"
Set CHRS = Nothing
Set WeeksWorked = New AllHours
Set CHRS = New AllHours
Set WeekS = New AllHours
Set PayRollSummary = New AllHours
Set EmpTotalHrs = New AllHours
Dim PayRollData As ServerPayRollData
Set DayOvrTime = New AllHours
Set ROT = New AllHours
RcdFound = False
If GetEmpLastName(BeginDate, EndDate, EmPNames) = False Then
  See_Msg "No Records Found", 5, ""
  Exit Sub
End If
SyncPayrollData BeginDate, EndDate
SumOnly = See_Msg("Show Summary Only", 4, "")
If RunPayroll_Cal(TimeRcd(), Employees(), BeginDate, EndDate, EmPNames, "Employee Records") = False Then Exit Sub
   PageHeader = "Period " & BeginDate & " To " & EndDate
   NetLaborSales = GetNetSales(BeginDate, EndDate)
   If Check_Option(96) = 0 Then
     Printer_Drivers "Receipt"
     PrinterPort = "Close Time Records"
     FP = 0
     PF = FreeFile
     Open PrinterPort For Output Shared As #PF
     CenterText = EC(8)
     LeftText = EC(9)
     'RightText = Chr(27) + Chr(97) + Chr(2)
     Print #PF, EC(0); EC(4) 'Int printer
     Print #PF, EC(1); EC(3); CenterText; StoreName; EC(2); EC(7)
     Print #PF, LeftText; EC(7)
     Print #PF, EC(1); EC(3); CenterText; PageHeader; EC(2); EC(7)
     Print #PF, LeftText; EC(7)
     Print #PF, Time;
     Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
     Print #PF, EC(3); String(40, 45); EC(7)
    Else
     SetPrinter 8
     PrintCLine StoreName, -1, 0
     PrintCLine CStr(Heading), -1, 0
     PrintCLine CStr(PageHeader), -1, 0
     PrintLine "Report Printed " & Format(Date, "mm/dd/yyyy") & " Printed By " & LoggedOnAs(0).EMPName, -1
     DrawLine
    End If
     For RCT = 0 To UBound(Employees, 1) ' - 1
        If Trim$(Employees(RCT).Name) = "" Then GoTo NXTEMP
        RcdFound = False: PntName = True
        CurrentWeek = -1
        SplitShift = "": ShiftDept = "": SplitShiftNo = 0
        On Error Resume Next
        For CT = 0 To UBound(TimeRcd, 1) - 1
         If Err.Number <> 0 Then Exit Sub
         On Error GoTo ErrorHandler
         With TimeRcd(CT)
          If TimeRcd(CT).Name = Employees(RCT).Name Then
           If PntName Then
            If SumOnly = 7 Then
            If Check_Option(96) = 0 Then
              Print #PF, "Name : "; Employees(RCT).Name; EC(7)
            Else
              'PrintLine "", -1
              PrintLine Employees(RCT).Name, -1
              DrawLine
            End If
            End If
           End If
           PntName = False
          RcdFound = True
          If .TimeON = .TimeOut Then
             InOutIsSame = True
          Else
             InOutIsSame = False
          End If
          If SplitShift <> .ShiftDate Or .ShiftNo <> SplitShiftNo And InOutIsSame = False Then
           PayRollData.CashTips = PayRollData.CashTips + .CashTip: PayRollData.CreditCardTips = PayRollData.CreditCardTips + .CCTip: PayRollData.TipsPaidOut = PayRollData.TipsPaidOut + .TipOut
           PayRollData.TipShare = PayRollData.TipShare + .TipShare: PayRollData.GratutityCollected = GetServerGratuity(Employees(RCT).Name, BeginDate, EndDate, 0)
           If .Sales > 0 Then PayRollData.TipSales = PayRollData.TipSales + .Sales
          Else
            .Sales = 0:  .CCTip = 0:
            PayRollData.CashTips = PayRollData.CashTips + .CashTip
            PayRollData.TipShare = PayRollData.TipShare + .TipShare
          End If
          If InOutIsSame = False Then
           SplitShift = .ShiftDate
           SplitShiftNo = .ShiftNo
          End If
          If Check_Option(96) = 0 Then
          
          Else
           ShiftTips = ((.CashTip + .CCTip) + .TipShare) - .TipOut
           If Val(Input_Option(55)) > 0 Then DayOtStr = IIf(.DayOt > 0, " Day OT " & .DayOt, "")
          End If   '// Option (55) = day overtime start
          CurrentWeek = TimeRcd(CT).PayWeek
          End If
         End With
        Next CT
        If RcdFound Then
        RcdFound = False
         If SumOnly = 7 Then
           If Check_Option(96) = 0 Then
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Dept Name               Reg Hr     OT Hr"; EC(7)
            Print #PF, String(40, 45); EC(7)
           Else
            'DrawLine
            PrintLine "Department", 4500: PrintLine "Regular Hours", 6500: PrintLine "Over Time", 8500: PrintLine "Sales", -1
            DrawLine
           End If
         End If
         End If
         EmpTotalRegHrs = 0: EmpTotalOtHrs = 0: GrossPay = 0: TotalDayOtHrs = 0: TipSales = 0
         For CT = 1 To CHRS.count
           If CHRS.Item(CT).RegHrs > 0 Then
            If CHRS.Item(CT).EMPName = Employees(RCT).Name Then
               If SumOnly = 7 Then
                If Check_Option(96) = 0 Then
                 Print #PF, CHRS.Item(CT).Job; Tab(25); Align(5, CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs); Tab(34); Align(5, CHRS.Item(CT).OTHrs + CHRS.Item(CT).DayOt); EC(7)
                Else
                 PrintLine CHRS.Item(CT).Job, 4500: PrintLine Format(CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs, "0.00"), 6500: PrintLine Format(CHRS.Item(CT).OTHrs + CHRS.Item(CT).DayOt, "0.00"), 8500: PrintLine CHRS.Item(CT).ShiftSales, -1
                End If
              End If
              GrossPay = GrossPay + CHRS.Item(CT).Pay
              EmpTotalRegHrs = EmpTotalRegHrs + CHRS.Item(CT).RegHrs
              EmpTotalOtHrs = EmpTotalOtHrs + CHRS.Item(CT).OTHrs
              TotalDayOtHrs = TotalDayOtHrs + CHRS.Item(CT).DayOt
              TipSales = TipSales + CHRS.Item(CT).ShiftSales
            End If
           End If
         Next CT
         If EmpTotalRegHrs > 0 Then
         If SumOnly = 7 Then
          If Check_Option(96) = 0 Then
          Print #PF, String(40, 45); EC(7)
          Print #PF, "Total Hours"; Tab(25); Align(5, EmpTotalRegHrs - EmpTotalOtHrs); Tab(34); Align(5, EmpTotalOtHrs + TotalDayOtHrs); EC(7)
          'Print #PF, String(40, 45); Ec(7)
          Else
          'DrawLine
          PrintLine "Totals", 4500: PrintLine Format(EmpTotalRegHrs - EmpTotalOtHrs, "0.00"), 6500: PrintLine Format(EmpTotalOtHrs + TotalDayOtHrs, "0.00"), 8500: PrintLine Format(TipSales, "0.00"), -1
          'DrawLine
          End If
         End If
         End If
         PayRollData.NetTips = ((((PayRollData.CashTips + PayRollData.CreditCardTips) + PayRollData.TipShare) + PayRollData.GratutityCollected) - PayRollData.TipsPaidOut)
         If PayRollData.NetTips > 0 And PayRollData.TipSales > 0 Then
          PercOfTip = Format(((PayRollData.NetTips) / PayRollData.TipSales) * 100, "0.00")
         End If
         If (TotalTips + PayRollData.GratutityCollected) > 0 And PayRollData.TipSales > 0 Then
          PercOfTip = Format(((TotalTips + PayRollData.GratutityCollected) / PayRollData.TipSales) * 100, "0.00")
         End If
         If SumOnly = 7 Then
         If Check_Option(96) = 0 Then
          If PayRollData.NetTips <> 0 Then
          Print #PF, "Total Tips  " & Format(PayRollData.NetTips, "0.00"); Tab(23); "Tip %   " & Format(PercOfTip, "0.00"); EC(7)
          End If
          Print #PF, String(40, 45); EC(7)
         Else
          If PayRollData.NetTips <> 0 Then
           PrintLine "Summary Tips ", 2500: PrintLine "Declared " & Format(PayRollData.CashTips, "0.00"), 4500: PrintLine "Gratuity Charged " & Format(PayRollData.GratutityCollected, "0.00"), 6500: PrintLine "cc Tips " & Format(PayRollData.CreditCardTips, "0.00"), 8000: PrintLine "Tips Paid " & Format(PayRollData.TipsPaidOut, "0.00"), 9500: PrintLine "Net Total Tips " & Format(PayRollData.NetTips, "0.00"), -1
          End If
          DrawLine
         End If
         End If
         PayRollData.NetTips = 0: PayRollData.CashTips = 0: PayRollData.CreditCardSales = 0
         PayRollData.CreditCardTips = 0: PayRollData.DeclaredTips = 0: PayRollData.GratutityCollected = 0
         PayRollData.TipsPaidOut = 0: PayRollData.TipShare = 0: PayRollData.TipSales = 0
NXTEMP:
        Next RCT
        If Check_Option(96) = 0 Then
         Print #PF, "Totals                  Reg Hr     OT Hr"; EC(7)
         Print #PF, String(40, 45); EC(7)
        Else
         DrawLine
         PrintLine "Summary Totals", 3500: PrintLine "Regular Hours", 5500: PrintLine "Over Time", 7500: PrintLine "Dollars", 9500: PrintLine "Percentage", -1
         DrawLine
        End If
        TotalRegHrs = 0: TotalOtHrs = 0
        Dim TotalDollars As Currency
        TotalDollars = 0
        For RCT = 1 To PayRollSummary.count
          For PD = 1 To CHRS.count
            If PayRollSummary.Item(RCT).Job = CHRS.Item(PD).Job Then
              TotalRegHrs = TotalRegHrs + CHRS.Item(PD).RegHrs
              TotalOtHrs = TotalOtHrs + CHRS.Item(PD).OTHrs
              TotalDollars = TotalDollars + CHRS.Item(PD).Pay
              PayRollSummary.Item(RCT).RegHrs = PayRollSummary.Item(RCT).RegHrs + CHRS.Item(PD).RegHrs
              PayRollSummary.Item(RCT).OTHrs = PayRollSummary.Item(RCT).OTHrs + CHRS.Item(PD).OTHrs
              'If PayRollSummary.Item(RCT).OTHrs > 0 Then Stop
              PayRollSummary.Item(RCT).Pay = PayRollSummary.Item(RCT).Pay + CHRS.Item(PD).Pay
            End If
          Next PD
        Next RCT
        For PD = 1 To PayRollSummary.count
         If Check_Option(96) = 0 Then
          Print #PF, PayRollSummary.Item(PD).Job; Tab(25); Align(5, PayRollSummary.Item(PD).RegHrs - PayRollSummary.Item(PD).OTHrs); Tab(34); Align(5, Val(Check_Null(PayRollSummary.Item(PD).OTHrs))); EC(7)
          If NetLaborSales > 0 And PayRollSummary.Item(PD).Pay > 0 Then
           Print #PF, "Dollars Spent"; Tab(25); Align(5, PayRollSummary.Item(PD).Pay); Tab(34); Align(5, (PayRollSummary.Item(PD).Pay / NetLaborSales) * 100); EC(7)
          End If
         Else
          If NetLaborSales > 0 And PayRollSummary.Item(PD).Pay > 0 Then
           PrintLine PayRollSummary.Item(PD).Job, 3500: PrintLine Format(PayRollSummary.Item(PD).RegHrs - PayRollSummary.Item(PD).OTHrs, "0.00"), 5500: PrintLine Format(Val(PayRollSummary.Item(PD).OTHrs & ""), "0.00"), 7500: PrintLine Format(PayRollSummary.Item(PD).Pay, "0.00"), 9500: PrintLine Format((PayRollSummary.Item(PD).Pay / NetLaborSales) * 100, "0.00"), -1
          Else
           PrintLine PayRollSummary.Item(PD).Job, 3500: PrintLine Format(PayRollSummary.Item(PD).RegHrs - PayRollSummary.Item(PD).OTHrs, "0.00"), 5500: PrintLine Format(Val(PayRollSummary.Item(PD).OTHrs & ""), "0.00"), 7500: PrintLine Format(PayRollSummary.Item(PD).Pay, "0.00"), -1  ': PrintLine Format((PayRollSummary.Item(PD).Pay / NetLaborSales) * 100, "0.00"), -1
          End If
          End If
         
        Next PD
        If Check_Option(96) = 0 Then
         Print #PF, String(40, 45); EC(7)
         Print #PF, "Total Hours"; Tab(25); Align(5, TotalRegHrs - TotalOtHrs); Tab(34); Align(5, TotalOtHrs); EC(7)
         If NetLaborSales > 0 And TotalDollars > 0 Then
          Print #PF, "Total Dollars"; Tab(25); Align(5, TotalDollars); Tab(34); Align(5, (TotalDollars / NetLaborSales) * 100); EC(7)
         End If
        Else
           DrawLine
           PrintLine "Total Hours Used   ", 4500: PrintLine Format(TotalRegHrs, "0.00"), -1
           PrintLine "Total OT-Hours Used", 4500: PrintLine Format(TotalOtHrs, "0.00"), -1
           PrintLine "Total Dollars Spent", 4500: PrintLine Format(TotalDollars, "0.00"), -1
           If NetLaborSales > 0 And TotalDollars > 0 Then
             PrintLine "Labor Percentage", 4500: PrintLine Format((TotalDollars / NetLaborSales) * 100, "0.00"), -1
           End If
           PrintLine "Net Sales", 4500: PrintLine Format(NetLaborSales, "0.00"), -1
           DrawLine
        End If
        If Check_Option(96) = 0 Then
         Print #PF, String(40, 45); EC(7)
         Print #PF, EC(6) & EC(6); EC(5)
         Close #PF
         SendJobToPrinter PrinterPort, "Receipt", False
        Else
         DrawLine
         FinishJob 1
        End If
Set WeeksWorked = Nothing
Set CHRS = Nothing
Set PayRollSummary = Nothing
Call MemoryCheck(Time_Sheet)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Close Time Records ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub GetTimeReport(BeginDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim Ttl_L           As Double
Dim Heading         As String
Dim PrinterPort     As String
Dim Total_Time      As Single
Dim SearchName      As String
Dim JCS             As Integer
Dim TimeIn          As String
Dim TimeOut         As String
Dim Db              As ADODB.Connection
Dim Tb              As ADODB.Recordset
Dim EDb             As ADODB.Connection
Dim Id              As ADODB.Recordset
Dim PRTb            As ADODB.Recordset
Dim TotalWorked     As Long
Dim THours          As Single
Dim PRate           As Currency
Dim TTips           As Currency '
Dim TSales          As Currency
Dim TDeclared       As Currency
Dim RCount          As Integer
Dim PIF             As Integer
Dim DollarsUsed     As Currency
Dim EmpTwt          As Single
Dim WrkTime         As Single
Dim HowPaid         As Integer
Dim EmpPay          As Currency
Dim NetLaborSales   As Currency
Dim Ck              As Integer
Dim BeginWeek       As String
Dim LaborDate       As String
Dim getweek         As Integer
Load_Depts

Printer_Drivers "Receipt"
DollarsUsed = 0
Heading = "Time Keeping Report"
'Start Routine for Finding Sales Date for labor report
'Dim NtSales As Currency, DailySales As Currency
        LaborDate = BeginDate
        getweek = GetDateToNo(CDate(BeginDate))
        Heading = "Daily Labor Report"
        ForWhen = BeginDate & " To " & EndDate
        NetLaborSales = GetNetSales(BeginDate, EndDate)
        If CDate(BeginDate) = CDate(Date) Then
          Sales_Breakdown True, 0
          NetLaborSales = TotalNetSales
        End If
SetPage ReportView.View, Heading & " " & ForWhen
TWRKTIME = 0
'End On Clock Routine
   DrawLine
   PrintLine "NAME", 2500: PrintLine "Department", 4500: PrintLine "In Time", 5500: PrintLine "Out Time", 6500: PrintLine "HOURS", 7500: PrintLine "DOLLARS", 8500: PrintLine "Total Tips", 9500: PrintLine "Tip Sales", -1
   DrawLine
OpenDb "MSaleData", EDb
Sql = "Select * from [General Data] Where ISNULL([Terminated], 0) = 0  order by [Last Name] asc"
ConnectADORS Sql, EDb, Id
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db


Do While Not Id.EOF
    SearchName = Id(0) & " " & Id(2)
    Pen = 0: EmpTwt = 0
    Sql = "Select * From [Job Codes] Where [First Name] = '" & Id(0) & "'"
    Sql = Sql & " and [Last Name] = '" & Id(2) & "'"
  ConnectADORS Sql, EDb, PRTb
    If Not PRTb.EOF Then
      ReDim PayRate(RecordsCount(PRTb))
      JCS = 0
    End If
    Do While Not PRTb.EOF
         PayRate(JCS).JobCodes = PRTb![Job Code] & ""
         PayRate(JCS).JobRate = Val(PRTb![rate] & "")
         PayRate(JCS).JobHourly = Val(PRTb![Hourly] & "")
         PayRate(JCS).UseDayOT = Val(PRTb![Day OT] & "")
         PayRate(JCS).JobWage = 0
         JCS = JCS + 1
      PRTb.MoveNext
    Loop
  SetNothing PRTb
  Sql = "Select * from [Employee Records] "
  Sql = Sql & " Where TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BeginDate) & " " & GetStartTime & "'"
  Sql = Sql & " and '" & SetEndingDate(CDate(EndDate)) & " " & SetEndingTime & "'"
  Sql = Sql & " and [Employee Name] = '" & SearchName & "'"
  Sql = Sql & " and isdate([Time In]) = 1"
  Sql = Sql & " Order By TRY_CAST([Time In] AS datetime) asc"
ConnectADORS Sql, Db, Tb
'If Not Tb.BOF Then Tb.MoveFirst
If Not Tb.EOF Then
  RCount = RecordsCount(Tb)
End If
PIF = 0
Do While Not Tb.EOF
PIF = PIF + 1
'If CDate(LaborDate) = CDate(Tb(2)) Then
'  If Trim(SearchName) = Trim(Tb(0)) Then
     Hoursworked = CalculateTimeRecord(Tb(2) & "", Tb(7) & "", Tb(3) & "", Tb(4) & "", Tb(5) & "", Tb(6) & "", Total_Time)
     TimeIn = Tb![Time In] & ""
     TimeOut = Tb![Time Out] & ""
     HowPaid = 0
     HW = Hoursworked
     If Total_Time > 0 Then Total_Time = Format(Total_Time / 60, "0.00")
      For Ck = 0 To JCS - 1
          If Trim(Tb(1)) = Trim(PayRate(Ck).JobCodes) Then
               PRate = PayRate(Ck).JobRate
               HowPaid = Abs(PayRate(Ck).JobHourly)
               Exit For
          End If
      Next Ck
   WrkTime = 0
   Cal_Labor Tb(1) & "", PRate, WrkTime, HowPaid, 1, OvertimeHrs(EmpTwt, WrkTime), getweek
   PrintLine SearchName, 2500
   PrintLine Tb(1) & "", 4500
   PrintLine Format(Tb(3) & "", "HH:MM"), 5500
   PrintLine Format(Tb(6) & "", "HH:MM"), 6500
   PrintLine Format(Total_Time, "0.00"), 7500
   EmpPay = 0
   If HowPaid = 0 Then
      'If Val(Input_Option(25)) > 0 Then
      ' EmpPay = PRate / Val(Input_Option(25))
      'Else
       EmpPay = PRate / 7
      'End If
   Else '// Option (55) = day overtime start
     If Val(Input_Option(55)) > 0 And QualifiesForDayOt(Tb(1), SearchName, PRate) Then  'Check_Option(62) = 1
        If Total_Time > Val(Input_Option(55)) Then
          EmpPay = (Total_Time - Val(Input_Option(55))) * PRate * 1.5
          EmpPay = EmpPay + Val(Input_Option(55)) * PRate
        Else
          EmpPay = Total_Time * PRate
        End If
     Else
      EmpPay = Total_Time * PRate
     End If
   End If
   PrintLine Format(EmpPay, "0.00"), 8500
      For Ck = 0 To Total_depts
    If Dept_Pay(Ck, 0).DeptName = Tb(1) Then
       Dept_Pay(Ck, 0).DeptDollars = Dept_Pay(Ck, 0).DeptDollars + EmpPay
       Dept_Pay(Ck, 0).TotatHours = Dept_Pay(Ck, 0).TotatHours + Total_Time
       Exit For
    End If
   Next Ck
   
   Dim PayRollData As ServerPayRollData
   PayRollData.CreditCardTips = Val(Tb.fields("Credit Card Tips").Value & "")
   PayRollData.DeclaredTips = Val(Tb.fields("Tip Amount").Value & "")
   GetPayRollInfo PayRollData, SearchName, Tb(2), Tb(2), Val(Tb![Shift] & "")
   TDeclared = (Val(Tb![Tip Amount] & ""))
   PrintLine Format((((TDeclared + PayRollData.CreditCardTips) + PayRollData.GratutityCollected) - PayRollData.TipsPaidOut), "0.00"), 9500
   PrintLine Format(PayRollData.TipSales, "0.00"), -1
   
   TSales = TSales + PayRollData.TipSales
   TTips = TTips + (((Val(TDeclared & "") + PayRollData.CreditCardTips) + PayRollData.GratutityCollected) - PayRollData.TipsPaidOut)
   THours = THours + Total_Time
   DollarsUsed = DollarsUsed + EmpPay
   Total_Time = 0
   PayRollData.TipSales = 0: PayRollData.CashTips = 0: PayRollData.CreditCardSales = 0
   PayRollData.DeclaredTips = 0: PayRollData.TipsPaidOut = 0: PayRollData.CreditCardTips = 0
   PayRollData.GratutityCollected = 0
   Tb.MoveNext
 Loop
'End If
Id.MoveNext
Loop
DbClose EDb, Id
DbClose Db, Tb
DrawLine
   PrintLine "Total", 6500
   PrintLine Format(THours, "0.00"), 7500
   PrintLine Format(DollarsUsed, "0.00"), 8500
   PrintLine Format(TTips, "0.00"), 9500
   PrintLine Format(TSales, "0.00"), -1
   DrawLine
   PrintLine "Department Labor", 4500: PrintLine "Dollars", 6500: PrintLine "Percentage", 8500
   PrintLine "Hour Used", -1
   DrawLine
Dim LabPer As Single
Dim Ttl_P As Single
For Pl = 0 To Total_depts - 1
    If Dept_Pay(Pl, 0).DeptDollars > 0 And NetLaborSales > 0 Then LabPer = Dept_Pay(Pl, 0).DeptDollars / NetLaborSales * 100
    If Dept_Pay(Pl, 0).DeptDollars > 0 Then
     PrintLine Dept_Pay(Pl, 0).DeptName, 4500: PrintLine Format(Val(Dept_Pay(Pl, 0).DeptDollars), "$0.00"), 6500: PrintLine Format(Val(LabPer), "%0.00"), 8500
     PrintLine Format(Val(Dept_Pay(Pl, 0).TotatHours), "0.00"), -1
    End If
    LabPer = 0
Next Pl
DrawLine
PrintLine UCase("Total Labor"), 4500: PrintLine Format(DollarsUsed, "$0.00"), 6500
If DollarsUsed > 0 And NetLaborSales > 0 Then Ttl_P = DollarsUsed / NetLaborSales * 100
PrintLine Format(Ttl_P, "%0.00"), 8500
PrintLine Format(THours, "0.00"), -1
DrawLine
PrintLine UCase("Total Net Sales"), 4500: PrintLine Format(NetLaborSales, "$0.00"), -1
ReportView.View.EndDoc
ReportView.Show 1
DoEvents
Reportday.Left = -3840
Reportday.Visible = False

Exit Sub
ErrorHandler:
   ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Time Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub DepartmentIndReport(BeginDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim Heading As String
Dim PrinterPort As String, Total_Time As Single
Dim TimeIn As String, TimeOut As String
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim TotalWorked As Long
Dim THours As Single
Dim TDeclared As Currency
Dim TCharged As Currency
Dim DptName As String
Dim DptHours As Single
Dim AllTips  As Currency
Printer_Drivers "Receipt"
Heading = "Department Report"
'Start Routine for Finding Sales Date for labor report
Dim BeginWeek As String, LaborDate As String, getweek As Integer
    LaborDate = BeginDate
    getweek = GetDateToNo(CDate(BeginDate))
    Heading = "Daily Labor Report"
    ForWhen = BeginDate & " To " & EndDate
    SetPage ReportView.View, Heading & " " & ForWhen
TWRKTIME = 0
'End On Clock Routine
DrawLine
PrintLine "Department", 2500: PrintLine "Employee", 5000: PrintLine "Date", 6000: PrintLine "In Time", 7000: PrintLine "Out Time", 8000: PrintLine "HOURS", 9000:  PrintLine "Total Tips", -1
DrawLine
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
Sql = "Select * from [Employee Records] "
Sql = Sql & " Where TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BeginDate) & " " & GetStartTime & "'"
Sql = Sql & " and '" & SetEndingDate(CDate(EndDate)) & " " & SetEndingTime & "'"
Sql = Sql & "Order by [Department Name],TRY_CAST([Shift Date] AS datetime),[Employee Name] asc"
ConnectADORS Sql, Db, Tb
DptName = ""
DptHours = 0
Do While Not Tb.EOF
   If DptName = "" Then DptName = Tb![Department Name] & ""
   If DptName <> Tb![Department Name] & "" Then
       DrawLine
       PrintLine DptName, 8000: PrintLine Format(DptHours, "0.00"), -1
       DrawLine
       DptName = Tb![Department Name] & ""
       DptHours = 0
   End If
   Hoursworked = CalculateTimeRecord(Tb(2) & "", Tb(7) & "", Tb(3) & "", Tb(4) & "", Tb(5) & "", Tb(6) & "", Total_Time)
   TimeIn = Tb![Time In] & ""
   TimeOut = Tb![Time Out] & ""
   HW = Hoursworked
   If Total_Time > 0 Then Total_Time = Total_Time / 60
   DptHours = DptHours + Total_Time
   PrintLine Tb![Department Name] & "", 2500
   PrintLine Tb![employee name] & "", 5000
   PrintLine Format(Tb![Shift date], "MM/DD"), 6000
   PrintLine Format(Tb(3) & "", "HH:MM"), 7000
   PrintLine Format(Tb(6) & "", "HH:MM"), 8000
   PrintLine Format(Total_Time, "0.00"), 9000
   TCharged = Val(Tb![Credit Card Tips] & "")
   TDeclared = Val(Tb![Tip Amount] & "") - Val(Tb![Tips Paid Out] & "")
   PrintLine Format(TCharged + TDeclared, "0.00"), -1
   Tb.MoveNext
 Loop
DrawLine
PrintLine DptName, 8000: PrintLine Format(DptHours, "0.00"), -1
DrawLine
DbClose Db, Tb
DrawLine
ReportView.View.EndDoc
ReportView.Show 1
'DoEvents
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Department Ind Report ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub Exit_Click()
Call MemoryCheck(Time_Sheet)
Call MemoryCheck(Time_Reports)
End Sub
Private Sub PrintPayRollsummary(BeginDate As String, EndDate As String)
On Error GoTo ErrorHandler
Dim TRDB                As ADODB.Connection
Dim TRRS                As ADODB.Recordset
Dim CT                  As Long
Dim RCT                 As Long
Dim PD                  As Long
Dim Job_Code            As String
Dim EMP_Rate            As Currency
Dim Emp_Name            As String
Dim PageHeader          As String
Dim TimeRcd()           As EmpTimeRecords
Dim Employees()         As EmpTimeRecords
Dim Breaks()            As BreakReport
Dim ShiftsWorked        As Integer
Dim CurrentWeek         As Integer
Dim WeekS               As AllHours
Dim WeeksWorked         As AllHours
Dim PayRollSummary      As AllHours
Dim EmpTotalHrs         As AllHours
Dim DayOvrTime          As AllHours
Dim ROT                 As AllHours
Dim DayOtStr            As Single
Dim EmpTotalRegHrs      As Single
Dim EmpTotalOtHrs       As Single
Dim NetLaborSales       As Currency
Dim TotalRegHrs         As Single
Dim TotalOtHrs          As Single
Dim Dept_Hrs            As Single
Dim Dept_Ot             As Single
Dim SplitShift          As String
Dim SplitShiftNo        As Long
Dim ShiftCredit         As Currency
Dim FindEmp             As Integer
Dim OTCal               As Single
Dim TotalDayOtHrs       As Single
Dim RegCal              As Single
Dim InOutIsSame         As Boolean
Dim EmPNames()          As ALLEMPTime
Dim PayRollTo           As String
Dim F                   As String
Dim H                   As String
Dim FEMP                As Integer
Dim TotalDeclare        As Currency
Dim TotalCCTip          As Currency
Dim TotalGrat           As Currency
Dim TotalSales          As Currency
Dim AllRegHrs           As Single
Dim AllOTHrs            As Single
Dim AllPay              As Currency
Dim AllDecl             As Currency
Dim AllCCTip            As Currency
Dim AllGrat             As Currency
Dim OTPay               As Currency
Dim AllRWage            As Currency
Dim AllOTWage           As Currency
Dim PEMPName            As String
Dim TipsPaidOut         As Currency
Set CHRS = Nothing
Set WeeksWorked = New AllHours
Set CHRS = New AllHours
Set WeekS = New AllHours
Set PayRollSummary = New AllHours
Set EmpTotalHrs = New AllHours
Dim PayRollData As ServerPayRollData
Set DayOvrTime = New AllHours
Set ROT = New AllHours
Dim Data As String
Data = ""
    Dim MySqlStr As String
    Dim THistory As String
    Dim trecords As String
    'TimeFile = TimeRecordsDb
If GetEmpLastName(BeginDate, EndDate, EmPNames) = False Then
   See_Msg "No Records Found", 5, ""
   Exit Sub
End If
If RunPayroll_Cal(TimeRcd(), Employees(), BeginDate, EndDate, EmPNames, "Employee Records") = False Then Exit Sub
   ReportView.View.Orientation = orLandscape
   PageHeader = "Employee Payroll Summary " & vbCrLf
   PageHeader = PageHeader & "Period " & BeginDate & " To " & EndDate
   SetPage ReportView.View, PageHeader
   NetLaborSales = GetNetSales(BeginDate, EndDate)
   F = "+<2000|+<1200|+<1400|+>750|+>1000|+>1000|+>1000|+>1000|+>1000|+>1000|+>1000|+>1000|+>1000|+>1000;"
   H = "Emp Name|Emp Id|Job Code|Rate|Reg Hrs|OT Hrs|Total Hrs|Wages|Declared|Net Tips|Gratuity|Tip Share|Total|Tip Sales;"
   ReportView.View.MarginLeft = ".20in"
   ReportView.View.CurrentY = 2700
   ReportView.View.StartTable
   ReportView.View.TableBorder = tbBoxRows
   For RCT = 0 To UBound(Employees, 1) ' - 1
    If Employees(RCT).Name = "" Then GoTo NXTEMP
         Data = Data & Employees(RCT).Name & "|"
         Data = Data & Employees(RCT).EMPID & "|"
         FEMP = 0
         CurrentWeek = -1:  ShiftDept = ""
         EmpTotalRegHrs = 0: EmpTotalOtHrs = 0: GrossPay = 0: TotalDayOtHrs = 0: TipsPaidOut = 0
         TotalDeclare = 0: TotalGrat = 0: TotalSales = 0: TotalCCTip = 0: OTPay = 0
         For CT = 1 To CHRS.count
           If CHRS.Item(CT).EMPName = Employees(RCT).Name Then
                 'If InStr(Employees(RCT).Name, "Cory") > 0 Then Stop
                 FEMP = FEMP + 1
                 If FEMP > 1 Then
                  Data = Data & "|"
                  Data = Data & "|"
                 End If
                 Data = Data & CHRS.Item(CT).Job & "|"
                 If Access_Required("MASK RATES") > 1 Then
                  Data = Data & "**.**" & "|"
                 Else
                  Data = Data & Format(CHRS.Item(CT).PayRate, "0.00") & "|"
                 End If
                 If CHRS.Item(CT).DayOt > 0 Then
                  If ((CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs) - CHRS.Item(CT).DayOt) < 0 Then
                    Data = Data & "" & "|"
                    Data = Data & Format(CHRS.Item(CT).RegHrs, "0.00") & "|"
                    CHRS.Item(CT).Pay = CHRS.Item(CT).Pay + ((CHRS.Item(CT).PayRate * 0.5) * CHRS.Item(CT).RegHrs)
                    EmpTotalOtHrs = EmpTotalOtHrs + CHRS.Item(CT).RegHrs
                  Else
                    Data = Data & Format(((CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs) - CHRS.Item(CT).DayOt), "0.00") & "|"
                    Data = Data & Format(CHRS.Item(CT).DayOt + CHRS.Item(CT).OTHrs, "0.00") & "|"
                    EmpTotalRegHrs = EmpTotalRegHrs + CHRS.Item(CT).RegHrs
                    EmpTotalOtHrs = EmpTotalOtHrs + CHRS.Item(CT).OTHrs
                  End If
                  TotalDayOtHrs = TotalDayOtHrs + CHRS.Item(CT).DayOt
                 Else
                  If CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs < 0 Then
                   Data = Data & "" & "|"
                   Data = Data & Format(CHRS.Item(CT).RegHrs, "0.00") & "|"
                   EmpTotalOtHrs = EmpTotalOtHrs + CHRS.Item(CT).RegHrs
                  Else
                   Data = Data & Format((CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs), "0.00") & "|"
                   Data = Data & Format(CHRS.Item(CT).OTHrs, "0.00") & "|"
                   EmpTotalRegHrs = EmpTotalRegHrs + CHRS.Item(CT).RegHrs
                   EmpTotalOtHrs = EmpTotalOtHrs + CHRS.Item(CT).OTHrs
                  End If
                 End If
                 'If CHRS.Item(Ct).GratuityCollect > 0 Then Stop
                 Data = Data & Format(CHRS.Item(CT).RegHrs, "0.00") & "|"
                 Data = Data & Format(CHRS.Item(CT).Pay, "0.00") & "|"
                 Data = Data & Format(CHRS.Item(CT).CashTips, "0.00") & "|"
                 Data = Data & Format(CHRS.Item(CT).ShiftccTip - CHRS.Item(CT).TipPaidOut, "0.00") & "|"
                 Data = Data & Format(CHRS.Item(CT).GratuityCollect, "0.00") & "|"
                 Data = Data & Format(CHRS.Item(CT).TipShare, "0.00") & "|"
                 '
                 Data = Data & Format((((CHRS.Item(CT).ShiftccTip + CHRS.Item(CT).CashTips) + CHRS.Item(CT).GratuityCollect) + CHRS.Item(CT).TipShare) - CHRS.Item(CT).TipPaidOut, "0.00") & "|"
                 'Data = Data & Format((CHRS.Item(Ct).ShiftccTip + (CHRS.Item(Ct).CashTips + CHRS.Item(Ct).GratuityCollect)), "0.00") & "|"  '+ CHRS.Item(Ct).GratuityCollect
                 Data = Data & Format(CHRS.Item(CT).ShiftSales, "0.00") & ";"
                 GrossPay = GrossPay + CHRS.Item(CT).Pay
                 OTPay = OTPay + CHRS.Item(CT).OTPay
                 TotalCCTip = TotalCCTip + CHRS.Item(CT).ShiftccTip - CHRS.Item(CT).TipPaidOut
                 TotalDeclare = TotalDeclare + CHRS.Item(CT).CashTips
                 TotalGrat = TotalGrat + CHRS.Item(CT).GratuityCollect
                 TotalSales = TotalSales + CHRS.Item(CT).ShiftSales
                 TipsPaidOut = TipsPaidOut + CHRS.Item(CT).TipPaidOut
           End If
         Next CT
         Data = Data & "|"
         Data = Data & "|"
         Data = Data & "          ** Total|"
         Data = Data & "|"
         If EmpTotalRegHrs - EmpTotalOtHrs < 0 Then
           Data = Data & "" & "|"
           Data = Data & Format(EmpTotalOtHrs, "0.00") & "|"
           Data = Data & Format(EmpTotalOtHrs, "0.00") & "|"
           AllOTHrs = AllOTHrs + EmpTotalOtHrs
           AllRegHrs = AllRegHrs + 0
         Else
           Data = Data & Format(((EmpTotalRegHrs - EmpTotalOtHrs) - TotalDayOtHrs), "0.00") & "|"
           Data = Data & Format(EmpTotalOtHrs + TotalDayOtHrs, "0.00") & "|"
           Data = Data & Format(EmpTotalRegHrs, "0.00") & "|"
           AllRegHrs = AllRegHrs + Format(((EmpTotalRegHrs - EmpTotalOtHrs) - TotalDayOtHrs), "0.00")
           AllOTHrs = AllOTHrs + (EmpTotalOtHrs + TotalDayOtHrs)
         End If
         If LoggedOnAs(0).EmpAccessLevel < Access_Required("MASK RATES") Then
          Data = Data & "XXX.XX" & "|"
         Else
          Data = Data & Format(GrossPay, "0.00") & "|"
         End If
         Data = Data & Format(TotalDeclare, "0.00") & "|"
         Data = Data & Format(TotalCCTip, "0.00") & "|"
         Data = Data & Format(TotalGrat, "0.00") & "|"
         Data = Data & Format(TipsPaidOut, "0.00") & "|"
         Data = Data & Format(((TotalDeclare + TotalCCTip) + TotalGrat), "0.00") & "|"  '+ TotalGrat
         Data = Data & Format(TotalSales, "0.00") & ";"
         'Data = Data & ";"
         AllRWage = AllRWage + (GrossPay - OTPay)
         AllOTWage = AllOTWage + OTPay
         AllPay = AllPay + GrossPay
         AllDecl = AllDecl + TotalDeclare
         AllCCTip = AllCCTip + TotalCCTip
         AllGrat = AllGrat + TotalGrat
NXTEMP:
       Next RCT
       ReportView.View.addTable F, H, Data, vbYellow
       ReportView.View.EndTable
       ReportView.View.MarginLeft = ".20in"
       ReportView.View.CurrentY = ReportView.View.CurrentY + 200

       ReportView.View.StartTable
       ReportView.View.TableBorder = tbBoxColumns
       Data = ""
       F = "+>1000|+>1000|+>1000|+>1500|+>1500|+>1500|+>1500|+>1500|+>1000|+>1000|+>1000;"
       H = "Reg Hrs|OT Hrs|Total Hrs|Reg Wages|OT Wages|Total Wages|Net Sales|Labor %|Declare|CC Tips|Grat;"
       Data = Data & Format(AllRegHrs, "0.00") & "|"
       Data = Data & Format(AllOTHrs, "0.00") & "|"
       Data = Data & Format(AllRegHrs + AllOTHrs, "0.00") & "|"
       Data = Data & Format(AllRWage, "0.00") & "|"
       Data = Data & Format(AllOTWage, "0.00") & "|"
       Data = Data & Format(AllPay, "0.00") & "|"
       Data = Data & Format(NetLaborSales, "0.00") & "|"
       If NetLaborSales > 0 And AllPay > 0 Then
         Data = Data & Format((AllPay / NetLaborSales) * 100, "0.00") & "|"
       Else
         Data = Data & "" & "|"
       End If
       Data = Data & Format(AllDecl, "0.00") & "|"
       Data = Data & Format(AllCCTip, "0.00") & "|"
       Data = Data & Format(AllGrat, "0.00") & ";"
       
       ReportView.View.addTable F, H, Data, vbGreen
       ReportView.View.EndTable
       FinishJob 1
Set WeeksWorked = Nothing
Set CHRS = Nothing
Set PayRollSummary = Nothing
Set DayOvrTime = Nothing
Set ROT = Nothing
Call MemoryCheck(Time_Sheet)
Exit Sub
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Print Payroll summary ")
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
ResizeScreen1024 Me
CenterObject Me, Frame1
LoadPRDates = "Select Distinct TRY_CAST([Shift Date] AS datetime) as myDate from [Employee Records] order by TRY_CAST([Shift Date] AS datetime) desc"
Reportday.Left = -3480
Reportday.Top = 240 * TopScale1024
'Load_Depts
Display = -1
End Sub
Private Sub QuickSort(ByRef pvarArray() As ALLEMPTime)
    Dim i          As Long
    Dim iMin       As Long
    Dim iMax       As Long
    Dim varSwap    As Variant
    Dim blnSwapped As Boolean
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray) - 1
    Dim FName      As String
    Dim LName      As String
    Dim MName      As String
    Dim FullName   As String
    Do
        blnSwapped = False
        For i = iMin To iMax
            If pvarArray(i).LastName > pvarArray(i + 1).LastName Then
                LName = pvarArray(i).LastName
                FName = pvarArray(i).FirstName
                MName = pvarArray(i).MiddleName
                FullName = pvarArray(i).WholeName
                pvarArray(i).LastName = pvarArray(i + 1).LastName
                pvarArray(i).FirstName = pvarArray(i + 1).FirstName
                pvarArray(i).MiddleName = pvarArray(i + 1).MiddleName
                pvarArray(i).WholeName = pvarArray(i + 1).WholeName
                
                pvarArray(i + 1).LastName = LName
                pvarArray(i + 1).FirstName = FName
                pvarArray(i + 1).MiddleName = MName
                pvarArray(i + 1).WholeName = FullName
                blnSwapped = True
            End If
        Next
        iMax = iMax - 1
    Loop Until Not blnSwapped
End Sub
Private Function GetEmpLastName(BeginDate As String, EndDate As String, allemp() As ALLEMPTime) As Boolean
Dim TRDB                As ADODB.Connection
Dim TRRS                As ADODB.Recordset
Dim AllNames()          As String
If ConnectADODB(SvrPath.TerminalPath & TimeRecordsDb, TRDB) = False Then Exit Function
   Sql = "Select distinct [Employee Name] from [Employee Records] where "
   Sql = Sql & " TRY_CAST([Shift Date] AS datetime) Between '" & CDate(BeginDate) & "'"
   Sql = Sql & " and '" & CDate(EndDate) & "'"
   If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
    If TRRS.EOF Then
       GetEmpLastName = False
       DbClose TRDB, TRRS
       Exit Function
    Else
       GetEmpLastName = True
    End If
    If Not TRRS.EOF Then
     RCT = RecordsCount(TRRS)
     ReDim allemp(RCT)
     CT = -1
    End If
    Do While Not TRRS.EOF
      CT = CT + 1
      allemp(CT).WholeName = Trim$(TRRS![employee name] & "")
      AllNames = Split(Trim$(TRRS![employee name] & ""), " ")
      Select Case UBound(AllNames, 1)
      Case Is = 0
       allemp(CT).FirstName = AllNames(0)
       allemp(CT).LastName = AllNames(0)
       allemp(CT).MiddleName = ""
      Case Is = 1
       allemp(CT).FirstName = AllNames(0)
       allemp(CT).LastName = AllNames(1)
       allemp(CT).MiddleName = ""
      Case Is = 2
       allemp(CT).FirstName = AllNames(0)
       allemp(CT).LastName = AllNames(2)
       allemp(CT).MiddleName = AllNames(1)
      Case Is = 3
       allemp(CT).FirstName = AllNames(0)
       allemp(CT).LastName = AllNames(2) & " " & AllNames(3)
       allemp(CT).MiddleName = AllNames(1)
      End Select
     TRRS.MoveNext
    Loop
    QuickSort allemp
End Function
Private Function RunPayroll_Cal(TimeRcd() As EmpTimeRecords, Employees() As EmpTimeRecords, BeginDate As String, EndDate As String, EmPNames() As ALLEMPTime, TableName As String) As Boolean
On Error GoTo ErrorHandler
Dim TRDB                As ADODB.Connection
Dim TRRS                As ADODB.Recordset
Dim F                   As String 'Reporting var
Dim H                   As String 'Reporting Var
Dim Data                As String 'Reporting
Dim CT                  As Long   'counting var
Dim RCT                 As Long
Dim PD                  As Long
Dim Job_Code            As String
Dim Emp_Name            As String
Dim CurrentWeek         As Integer
Dim EMP_Rate            As Currency
Dim SplitShift          As String
Dim SplitShiftNo        As Long
Dim ShiftCredit         As Currency
Dim InOutIsSame         As Boolean
Dim DayInPayroll        As Integer
Dim DayOverTimeStarts   As Integer
Dim Export              As ExportPRData
Dim WeekS               As AllHours
Dim WeeksWorked         As AllHours
Dim EmpTotalHrs         As AllHours
Dim DayOverTime         As AllHours
Dim RegularOverTime     As AllHours
DayOverTimeStarts = Val(Input_Option(55))  'value from register options
'// Option (55) = day overtime start
Set WeekS = Nothing
Set WeekS = New AllHours

Set WeeksWorked = Nothing
Set WeeksWorked = New AllHours

Set RegularOverTime = Nothing
Set RegularOverTime = New AllHours

Set EmpTotalHrs = Nothing
Set EmpTotalHrs = New AllHours

Set PayRollSummary = Nothing
Set PayRollSummary = New AllHours

Set DayOverTime = Nothing
Set DayOverTime = New AllHours

Set Break = Nothing
Set Break = New AllHours


Dim PayRollData As ServerPayRollData

RunPayroll_Cal = True

If ConnectADODB(SvrPath.LocalTerminalPath & TimeRecordsDb, TRDB) = False Then Exit Function
   
   Sql = "DELETE FROM [Employee Record] " & _
      "WHERE [Employee Name] IS NULL " & _
      "OR LTRIM(RTRIM([Employee Name])) = ''"
   TRDB.Execute Sql
   
   Sql = "Set datefirst " & WeekStartsOn  'used for first day of the payroll week
   TRDB.Execute Sql
   
   
   Sql = "Select * from [Breaks]"
   Sql = Sql & " Where TRY_CAST([TimeOut] AS datetime)  Between '" & CDate(BeginDate) & " " & GetStartTime & "'"
   Sql = Sql & " and '" & SetEndingDate(CDate(EndDate)) & " " & SetEndingTime & "'"
   Sql = Sql & " Order by [Name] asc"
   If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
   Do While Not TRRS.EOF
    If TRRS.fields("RecordId").Value > "" Then
     If Break.Item(TRRS.fields("RecordId").Value) Is Nothing Then
        Break.add TRRS.fields("RecordId").Value
     End If
     With Break
        .Item(TRRS.fields("RecordId").Value).RecordId = TRRS.fields("RecordId").Value
        .Item(TRRS.fields("RecordId").Value).EMPName = TRRS.fields("Name").Value
        .Item(TRRS.fields("RecordId").Value).BreakOut = TRRS.fields("TimeOut").Value
        .Item(TRRS.fields("RecordId").Value).BreakIn = TRRS.fields("TimeIn").Value
        .Item(TRRS.fields("RecordId").Value).Job = TRRS.fields("BreakType").Value
     End With
    End If
   TRRS.MoveNext
   Loop
   TRRS.Close
   
   Sql = "Select * from [" & TableName & "]"
   Sql = Sql & " Where TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BeginDate) & " " & GetStartTime & "'"
   Sql = Sql & " and '" & SetEndingDate(CDate(EndDate)) & " " & SetEndingTime & "'"
   Sql = Sql & " and isdate([Shift Date] + [time out]) = 1"
   If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
   If Not TRRS.EOF Then
      RCT = RecordsCount(TRRS)
      See_Msg "Please Check Time Records", 0, CStr(RCT) & " Invalid Punch Time Were Found !!"
      SetPage ReportView.View, "Invalid Punch Times"
      H = "Employee Name|Date;"
      F = "+<2500|+>1000;"
      Data = ""
      Do While Not TRRS.EOF
        Data = Data & TRRS.fields("Employee Name").Value & "|"
        Data = Data & TRRS.fields("Shift Date").Value & ";"
        TRRS.MoveNext
      Loop
      ReportView.View.TableBorder = tbBoxColumns
      ReportView.View.addTable F, H, Data, vbBlack, vbWhite
      DbClose TRDB, TRRS
      FinishJob 0
      RunPayroll_Cal = False
      Exit Function
   End If
   DayInPayroll = DateDiff("d", CDate(BeginDate) & " " & GetStartTime, SetEndingDate(CDate(EndDate)) & " " & SetEndingTime, 3)
   SetNothing TRRS
    Sql = "Select Distinct[Employee Name] from [" & TableName & "]"
    Sql = Sql & " Where TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BeginDate) & " " & GetStartTime & "'"
    Sql = Sql & " and '" & SetEndingDate(CDate(EndDate)) & " " & SetEndingTime & "'"
    If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
    RCT = RecordsCount(TRRS)
    ReDim Employees(RCT)
    CT = -1
   For FindEmp = 0 To UBound(EmPNames, 1)
    If EmPNames(FindEmp).WholeName > "" Then
     Sql = "Select Distinct [Employee Name] from [" & TableName & "] Where "
     Sql = Sql & " TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BeginDate) & " " & GetStartTime & "'"
     Sql = Sql & " and '" & SetEndingDate(CDate(EndDate)) & " " & SetEndingTime & "'"
     Sql = Sql & " and [employee Name] = '" & EmPNames(FindEmp).WholeName & "'"
     If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
     Do While Not TRRS.EOF
       CT = CT + 1
       Employees(CT).Name = TRRS![employee name] & ""
       Employees(CT).EMPID = GetSSN(TRRS![employee name] & "")
       Employees(CT).Job = Employees(CT).EMPID
       TRRS.MoveNext
     Loop
     SetNothing TRRS
    End If
   Next FindEmp
   'MsgBox "Find Employee " & Ct
    Sql = "Select Distinct [shift Date] from [" & TableName & "] Where "
    Sql = Sql & " TRY_CAST([Shift Date] AS datetime) Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
    'Sql = Sql & " Order by [shift date] desc"
    If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
    Do While Not TRRS.EOF
     ' Debug.Print TRRS![Shift date]
      CurrentWeek = DatePart("ww", CDate(TRRS![Shift date]), WeekStartsOn, vbFirstFullWeek) 'vbFirstFullWeek
      If WeekS.Item(CStr(CurrentWeek)) Is Nothing Then
       WeekS.add (CStr(CurrentWeek))
       WeekS.Item(CStr(CurrentWeek)).ShiftWeek = CurrentWeek
       WeekS.Item(CStr(CurrentWeek)).ShiftYear = Year(CDate(TRRS![Shift date]))
      End If
     TRRS.MoveNext
    Loop
    SetNothing TRRS
    For CT = 1 To WeekS.count
     Sql = "Select * from [" & TableName & "] Where "
     Sql = Sql & " dateDiff(m,TRY_CAST([shift date] AS datetime),'" & CDate(BeginDate) & "') < 2"
     Sql = Sql & " order by [employee Name], TRY_CAST([shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) ,TRY_CAST([Punch Out Date] AS datetime) + TRY_CAST([time out] AS datetime) asc"
     If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
     Do While Not TRRS.EOF
      ' Debug.Print TRRS![Shift date]
       If DatePart("ww", CDate(TRRS![Shift date]), WeekStartsOn, vbFirstFullWeek) = WeekS.Item(CT).ShiftWeek Then
       Job_Code = Trim$(TRRS![Department Name] & ""): Emp_Name = Trim$(TRRS![employee name])
       EMP_Rate = Val(TRRS![Rate of Pay] & "")
        If EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)) Is Nothing Then
         EmpTotalHrs.add Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)
         EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).OTReach = False
         EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).OverHours = 0
        End If
        If WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)) Is Nothing Then
         WeeksWorked.add (Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek))
         WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).RegHrs = 0
         WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).DayOt = 0
        End If
        If WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)) Is Nothing Then
         WeeksWorked.add (Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek))
         WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TimeBeforeOT = 0
        End If
        EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).EMPName = Emp_Name
        CalculateTimeRecord TRRS![Shift date], TRRS![Punch out Date], TRRS![Time In], TRRS![Break Out] & "", TRRS![Break In] & "", TRRS![Time Out], Total_Time
        WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).EMPName = Emp_Name
        WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).RegHrs = WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).RegHrs + ConvertTimeToMinutes(Total_Time)
        EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours = EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours + ConvertTimeToMinutes(Total_Time)
        If DayOverTimeStarts > 0 Then
         If QualifiesForDayOt(Job_Code, Emp_Name, EMP_Rate) Then
          If WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).RegHrs > DayOverTimeStarts Then
           If EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours > Val(Input_Option(49)) And WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).CADayOt = 0 Then
            WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).CADayOt = (WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).RegHrs - DayOverTimeStarts)
           End If
           WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).DayOt = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).DayOt + (WeeksWorked.Item(Emp_Name & TRRS![Shift date] & CStr(WeekS.Item(CT).ShiftWeek)).RegHrs - DayOverTimeStarts)
          End If
         End If
        End If
        If EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).OTReach = False Then 'Or EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(Ct).ShiftWeek)).TotalHours > Val(Input_Option(49)) Then
         WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TimeBeforeOT = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TimeBeforeOT + ConvertTimeToMinutes(Total_Time)
         If WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TimeBeforeOT > Val(Input_Option(49)) Then
           WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TimeBeforeOT = Val(Input_Option(49))
         End If  ' catch hours over 40
        End If
        WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours + ConvertTimeToMinutes(Total_Time)
        'WeeksWorked.Item(Emp_Name & CStr(WeekS.Item(Ct).ShiftWeek)).TotalHours = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(Ct).ShiftWeek)).TotalHours + ConvertTimeToMinutes(Total_Time)
        If CDate(TRRS.fields("Shift Date").Value) > CDate(EndDate) Then
          WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OverHours = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OverHours + ConvertTimeToMinutes(Total_Time)
        End If
        If EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).OTReach = True Then
          WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OTHrs = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OTHrs + ConvertTimeToMinutes(Total_Time)
          WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OverHours = WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OTHrs
          'added this line 12.03.2018 caused calculation error
        End If
        If WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OverHours = 0 Then
        'dont change this line will mess up over time by day
         If EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours > Val(Input_Option(49)) Then
          If EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).OTReach Then
            If WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours > Val(Input_Option(49)) Then
              WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OTHrs = (EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours - Val(Input_Option(49)) - WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).DayOt)
            End If
          Else
            EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).OTReach = True
            WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).OTHrs = Format(((EmpTotalHrs.Item(Emp_Name & CStr(WeekS.Item(CT).ShiftWeek)).TotalHours + WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).CADayOt) - Val(Input_Option(49)) - WeeksWorked.Item(Emp_Name & Job_Code & CStr(WeekS.Item(CT).ShiftWeek)).DayOt), "0.00")
          End If
         End If
       End If
     
      Else
       'See_Msg "Problem", 0, ""
      End If
      TRRS.MoveNext
     Loop
     SetNothing TRRS
    Next CT
   Set WeekS = Nothing
   
   Sql = "Select * from [" & TableName & "] Where "
   Sql = Sql & " isdate([Punch Out Date]+[time Out]) = 0"
   Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) Between '" & CDate(BeginDate) & "'"
   Sql = Sql & " and '" & CDate(EndDate) & "'"
   Sql = Sql & " Order by [Employee Name], TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) asc" ',[Department Name]
   If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
    If TRRS.EOF Then
      DbClose TRDB, TRRS
      See_Msg "No Records Found", 5, "Loading Records"
      Exit Function
    End If
    If Not TRRS.EOF Then
     RCT = RecordsCount(TRRS)
     ReDim TimeRcd(RCT)
     CT = -1
    End If
    Sql = "Select * from [" & TableName & "] Where "
    Sql = Sql & " isdate([Punch Out Date]+[time Out]) = 0"
    Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) Between '" & CDate(BeginDate) & "'"
    Sql = Sql & " and '" & CDate(EndDate) & "'"
    Sql = Sql & " Order by [Employee Name],  TRY_CAST([Shift Date] AS datetime)+ TRY_CAST([Time In] AS datetime),TRY_CAST([Punch Out Date] AS datetime)+ TRY_CAST([time out] AS datetime) asc" ',[Department Name]
    If ConnectADORS(Sql, TRDB, TRRS) = False Then Exit Function
    Do While Not TRRS.EOF
      CT = CT + 1
       CurrentWeek = DatePart("ww", CDate(TRRS![Shift date]), WeekStartsOn, vbFirstFullWeek) 'vbFirstFullWeek
       Job_Code = Trim$(TRRS![Department Name] & ""): Emp_Name = Trim$(TRRS![employee name])
       If PayRollSummary.Item(Job_Code) Is Nothing Then
         PayRollSummary.add Job_Code
         PayRollSummary.Item(Job_Code).Job = Job_Code
       End If
       With TimeRcd(CT)
          .PayWeek = CurrentWeek
          .Name = Emp_Name
          .Job = Job_Code
          .ShiftDate = TRRS![Shift date] & ""
          .TimeON = TRRS![Time In] & ""
          .BkOut = TRRS![Break Out] & ""
          .BkIn = TRRS![Break In] & ""
          .TimeOut = TRRS![Time Out] & ""
          .EndShiftDate = TRRS![Punch out Date] & ""
          If TRRS![Altered By] & "" = "Approved" Or TRRS![Altered By] & "" = "Not Approved" Then
           .EMPAP = TRRS![Altered By]
          Else
           .MGR = TRRS![Altered By] & ""
          End If
          .TOnAlt = TRRS![Altered In] & ""
          .TOntAlt = TRRS![Altered Out] & ""
          .BIAlt = TRRS![Altered Brk In] & ""
          .BOAlt = TRRS![Altered Brk Out] & ""
          .PayRate = Format(Val(TRRS![Rate of Pay] & ""), "0.00")
          .OTRate = ReturnOTRate(Format(Val(TRRS.fields("Rate of Pay").Value), "0.00"))
          .OTPay = ReturnOTRate(Format(Val(TRRS.fields("Rate of Pay").Value), "0.00"))
          .PaidBy = Val(TRRS![Paid By] & "")
          .Sales = Format(Val(TRRS![Sales] & ""), "0.00")
          .CashTip = Format(Val(TRRS![Tip Amount] & ""), "0.00")
          .CCTip = Format(Val(TRRS![Credit Card Tips] & ""), "0.00")
          .TipOut = Format(Val(TRRS![Tips Paid Out] & ""), "0.00")
          .TipShare = Format(Val(TRRS![Tip Share] & ""), "0.00")
          .ShiftNo = Val(TRRS![Shift] & "")
          .EMPID = TRRS![SSN] & ""
          .EMPAP = TRRS![Altered By] & ""
          .RecordId = TRRS.fields("RecordId").Value & ""
          Call CalculateTimeRecord(.ShiftDate, .EndShiftDate, .TimeON, .BkOut, .BkIn, .TimeOut, Total_Time)
          .RegHrs = ConvertTimeToMinutes(Total_Time)
          If DayOverTimeStarts > 0 And QualifiesForDayOt(Job_Code, Emp_Name, TimeRcd(CT).PayRate) Then
           If .RegHrs > DayOverTimeStarts Then
            .DayOt = Format(.RegHrs - DayOverTimeStarts, "0.00")
           Else
            .DayOt = 0
           End If
          End If
          .OTHrs = 0
          If .TimeON = .TimeOut Then
             InOutIsSame = True
          Else
             InOutIsSame = False
          End If
          If SplitShift <> .ShiftDate Or .ShiftNo <> SplitShiftNo And InOutIsSame = False Then
          Else
              .Sales = 0:  .CCTip = 0
          End If
          If InOutIsSame = False Then
             SplitShift = .ShiftDate
             SplitShiftNo = .ShiftNo
          End If
          If Check_Option(61) = 1 And PayRollTo = "QUICKBOOKS" Then
             Export.Id = .EMPID: Export.DeptName = .Job: Export.DayOt = .DayOt
             Export.EMPName = .Name: Export.BeginPeriod = .ShiftDate + " " + .TimeON
             Export.EndPeriod = .EndShiftDate + " " + .TimeOut
             Export.RegTime = .RegHrs
             ExportPayroll Export
          End If
          'If SiteLic.EnterPrize = 1 Then
            ExportTimeRecords .Name, Format(.ShiftDate, "mm/dd/yyyy"), Format(.TimeON, "HH:MM"), Format(.BkOut, "HH:MM"), Format(.BkIn, "HH:MM"), Format(.TimeOut, "HH:MM"), Format(.RegHrs, "0.00"), Format(.DayOt, "0.00"), Format(.RegHrs + .DayOt, "0.00"), .Job, .EMPID, .EMPAP
          'End If
       End With
        If RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)) Is Nothing Then RegularOverTime.add (Emp_Name & CStr(CurrentWeek))
        If DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")) Is Nothing Then DayOverTime.add Emp_Name & CStr(TRRS![Shift date] & "")
        If CHRS.Item(Emp_Name & Job_Code) Is Nothing Then CHRS.add (Emp_Name & Job_Code)
        CHRS.Item(Emp_Name & Job_Code).SSN = Trim$(TRRS![SSN] & "")
        CHRS.Item(Emp_Name & Job_Code).ShiftsWorked = CHRS.Item(Emp_Name & Job_Code).ShiftsWorked + 1
        CHRS.Item(Emp_Name & Job_Code).EMPName = Trim$(TRRS![employee name] & "")
        CHRS.Item(Emp_Name & Job_Code).Job = Trim$(TRRS![Department Name] & "")
        CHRS.Item(Emp_Name & Job_Code).PayRate = TimeRcd(CT).PayRate
        CHRS.Item(Emp_Name & Job_Code).OTRate = TimeRcd(CT).OTRate
        If Val(TRRS![Paid By] & "") = 0 Then 'paid by 0 for salary and 1 for hourly
          If TimeRcd(CT).PayRate > 0 Then
           CHRS.Item(Emp_Name & Job_Code).Pay = (TimeRcd(CT).PayRate / 7) * DayInPayroll
          Else
           CHRS.Item(Emp_Name & Job_Code).Pay = TimeRcd(CT).PayRate
          End If
        Else
          CHRS.Item(Emp_Name & Job_Code).Pay = CHRS.Item(Emp_Name & Job_Code).Pay + Format((TimeRcd(CT).RegHrs * TimeRcd(CT).PayRate), "0.00")
        End If
        CHRS.Item(Emp_Name & Job_Code).HourlyPos = Val(TRRS![Paid By] & "")
        CHRS.Item(Emp_Name & Job_Code).RegHrs = CHRS.Item(Emp_Name & Job_Code).RegHrs + TimeRcd(CT).RegHrs
        RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs = RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs + TimeRcd(CT).RegHrs
        If RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs > Val(Input_Option(49)) Then
         TimeRcd(CT).OTHrs = RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs - Val(Input_Option(49))
         RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs = RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs - (RegularOverTime.Item(Emp_Name & CStr(CurrentWeek)).RegHrs - Val(Input_Option(49)))
        End If
        If DayOverTimeStarts > 0 And QualifiesForDayOt(Job_Code, Emp_Name, TimeRcd(CT).PayRate) Then
          DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOt = DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOt + TimeRcd(CT).RegHrs
          If DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOt > DayOverTimeStarts Then 'And TimeRcd(Ct).DayOt = 0 Then
            If TimeRcd(CT).OTHrs > 0 Then
            
            Else
             TimeRcd(CT).DayOt = (Format(DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOt - DayOverTimeStarts, "0.00") - DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOTReached)
             CHRS.Item(Emp_Name & Job_Code).DayOt = CHRS.Item(Emp_Name & Job_Code).DayOt + TimeRcd(CT).DayOt
             DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOTReached = DayOverTime.Item(Emp_Name & CStr(TRRS![Shift date] & "")).DayOTReached + TimeRcd(CT).DayOt
             CHRS.Item(Emp_Name & Job_Code).Pay = CHRS.Item(Emp_Name & Job_Code).Pay + (TimeRcd(CT).DayOt * (TimeRcd(CT).PayRate * 0.5))
            End If
          Else
            CHRS.Item(Emp_Name & Job_Code).DayOt = CHRS.Item(Emp_Name & Job_Code).DayOt + TimeRcd(CT).DayOt
          End If
        Else
          CHRS.Item(Emp_Name & Job_Code).DayOt = 0
        End If
        
        If Not WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)) Is Nothing Then
         If WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).EMPName = Emp_Name Then
          If WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).DayOt > 0 Then
           WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).DayOt = 0
          End If
          If WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).OTHrs > 0 Then
            CHRS.Item(Emp_Name & Job_Code).OTHrs = CHRS.Item(Emp_Name & Job_Code).OTHrs + WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).OTHrs
            If Val(TRRS![Paid By] & "") = 0 Then
            
            Else
             CHRS.Item(Emp_Name & Job_Code).Pay = CHRS.Item(Emp_Name & Job_Code).Pay + ((WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).OTHrs) * (TimeRcd(CT).OTPay - TimeRcd(CT).PayRate))
             CHRS.Item(Emp_Name & Job_Code).OTPay = CHRS.Item(Emp_Name & Job_Code).OTPay + ((WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).OTHrs) * ReturnOTRate(TimeRcd(CT).PayRate))
            'Fixed Over Time Calculation for Split weeks
            End If
            WeeksWorked.Item(Emp_Name & Job_Code & CStr(CurrentWeek)).OTHrs = 0
          End If
         End If
        End If
        CHRS.Item(Emp_Name & Job_Code).CashTips = CHRS.Item(Emp_Name & Job_Code).CashTips + TimeRcd(CT).CashTip
        CHRS.Item(Emp_Name & Job_Code).ShiftSales = CHRS.Item(Emp_Name & Job_Code).ShiftSales + TimeRcd(CT).Sales
        CHRS.Item(Emp_Name & Job_Code).ShiftccTip = CHRS.Item(Emp_Name & Job_Code).ShiftccTip + TimeRcd(CT).CCTip
        CHRS.Item(Emp_Name & Job_Code).TipPaidOut = CHRS.Item(Emp_Name & Job_Code).TipPaidOut + TimeRcd(CT).TipOut
        CHRS.Item(Emp_Name & Job_Code).TipShare = CHRS.Item(Emp_Name & Job_Code).TipShare + TimeRcd(CT).TipShare
        CHRS.Item(Emp_Name & Job_Code).GratuityCollect = GetServerGratuity(Emp_Name, BeginDate, EndDate, 0)  'TimeRcd(Ct).sales)
    TRRS.MoveNext
    Loop
   DbClose TRDB, TRRS
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Calculate time records ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Sub PrintPayRoll(DailyExport As Boolean, BeginDate As String, EndDate As String, PreviewRecords As Boolean, Header As String)
On Error GoTo ErrorHandler
Dim TRDB                As ADODB.Connection
Dim TRRS                As ADODB.Recordset
Dim CT                  As Long
Dim RCT                 As Long
Dim PD                  As Long
Dim Job_Code            As String
Dim Emp_Name            As String
Dim PageHeader          As String
Dim TimeRcd()           As EmpTimeRecords
Dim Employees()         As EmpTimeRecords
Dim ShiftsWorked        As Integer
Dim CurrentWeek         As Integer
Dim WeekS               As AllHours
Dim WeeksWorked         As AllHours
'Dim PayRollSummary      As AllHours
Dim EmpTotalHrs         As AllHours
Dim DayOvrTime          As AllHours
Dim ROT                 As AllHours
Dim PrinterPort         As String
Dim TimeEdits(0 To 3)   As String
Dim ShiftTips           As Currency
Dim DayOtStr            As Single
Dim EmpTotalRegHrs      As Single
Dim EmpTotalOtHrs       As Single
Dim ExportFileTo        As String
Dim NetLaborSales       As Currency
Dim TotalRegHrs         As Single
Dim TotalOtHrs          As Single
Dim Dept_Hrs            As Single
Dim Dept_Ot             As Single
Dim SplitShift          As String
Dim SplitShiftNo        As Long
Dim ShiftCredit         As Currency
Dim FindEmp             As Integer
Dim OTCal               As Single
Dim TotalDayOtHrs       As Single
Dim RegCal              As Single
Dim InOutIsSame         As Boolean
Dim EmPNames()          As ALLEMPTime
Dim EMP_Rate            As Currency
Dim Exports             As ExportPRData
Dim EmployeeTable       As String
Dim BKMINS              As Single
If Check_Option(61) = 1 Then
  PayRollTo = UCase(GetConfigInfo("Exports", "Export Format", "DEFAULT", "PayRollExport.INI"))
End If
Set CHRS = Nothing
Set CHRS = New AllHours

Set WeeksWorked = Nothing
Set WeeksWorked = New AllHours

Set WeekS = Nothing
Set WeekS = New AllHours

Set PayRollSummary = Nothing
Set PayRollSummary = New AllHours

Set EmpTotalHrs = Nothing
Set EmpTotalHrs = New AllHours

Set DayOvrTime = Nothing
Set DayOvrTime = New AllHours

Set ROT = Nothing
Set ROT = New AllHours



Dim PayRollData As ServerPayRollData
EmployeeTable = "Employee Records"

DeleteFile SvrPath.TerminalPath & "TimeRecords.CSV"
DeleteFile SvrPath.TerminalPath & "PayRollExport.CSV"

If Check_Option(61) = 1 Then
  If CustOption.PayRollFileName > "" Then
     ExportFileTo = IIf(GetConfigInfo("Exports", "Export Location", "", "PayRollExport.INI") = "", SvrPath.TerminalPath, GetConfigInfo("Exports", "Export Location", "", "PayRollExport.INI"))
     If Dir(ExportFileTo & CustOption.PayRollFileName) > "" Then
      DeleteFile ExportFileTo & CustOption.PayRollFileName
     End If
  End If
End If
If GetEmpLastName(BeginDate, EndDate, EmPNames) = False Then
   See_Msg "No Records Found", 5, ""
   Exit Sub
End If
If Header = "Individual Time Records" Then
  Call ConnectADODB(SvrPath.LocalTerminalPath & TimeRecordsDb, TRDB)
  Sql = "IF EXISTS(Select * from information_schema.tables Where table_Name = 'EmpRecords' and table_schema = 'dbo') drop table dbo.[EmpRecords]"
  TRDB.Execute Sql
  Sql = "Select * INTO [EmpRecords] from [Employee Records]"
  TRDB.Execute Sql
  For FindEmp = 0 To List2.ListCount - 1
   If List2.Selected(FindEmp) = False Then
     Sql = "Delete from [EmpRecords] Where [Employee Name] = '" & List2.List(FindEmp) & "'"
     TRDB.Execute Sql
   End If
 Next FindEmp
 EmployeeTable = "EmpRecords"
 SetNothing TRDB
End If

'// make sure that setting are in the MSaledata database
See_Msg "Calculating Payroll", 5, ""
If RunPayroll_Cal(TimeRcd(), Employees(), BeginDate, EndDate, EmPNames, EmployeeTable) = False Then Exit Sub
For PD = 1 To CHRS.count
  If Check_Option(61) = 1 Then
     If Check_Option(61) = 1 And PayRollTo <> "QUICKBOOKS" Then
       With Exports
             CHRS.Item(PD).SSN = GetSSN(CHRS.Item(PD).EMPName)
            .Id = CHRS.Item(PD).SSN
            .SSNumber = CHRS.Item(PD).SSN
            .DeptName = CHRS.Item(PD).Job
            .RegTime = CHRS.Item(PD).RegHrs - (CHRS.Item(PD).OTHrs + CHRS.Item(PD).DayOt)
            .OT = CHRS.Item(PD).OTHrs + CHRS.Item(PD).DayOt '
            .DayOt = CHRS.Item(PD).DayOt
            .TipAmt = Format(((((CHRS.Item(PD).CashTips + CHRS.Item(PD).ShiftccTip) + CHRS.Item(PD).GratuityCollect)) - CHRS.Item(PD).TipPaidOut), "0.00")
            .RateofPay = CHRS.Item(PD).PayRate
            .OTPayRate = CHRS.Item(PD).OTRate
            .EMPName = CHRS.Item(PD).EMPName
            .TipSales = CHRS.Item(PD).ShiftSales
            .ShiftsWorked = CHRS.Item(PD).ShiftsWorked
            .BeginPeriod = BeginDate
            .EndPeriod = EndDate
            .GrossPayAmt = CHRS.Item(PD).Pay
            .TipShare = CHRS.Item(PD).TipShare
            .CreditCardTips = CHRS.Item(PD).ShiftccTip
       End With
       ExportPayroll Exports
     End If
  End If
  EnterPrizePayExport CHRS.Item(PD).SSN, BeginDate, EndDate, CHRS.Item(PD).Job, CHRS.Item(PD).EMPName, CHRS.Item(PD).RegHrs - (CHRS.Item(PD).OTHrs + CHRS.Item(PD).DayOt), CHRS.Item(PD).OTHrs + CHRS.Item(PD).DayOt, CHRS.Item(PD).Pay
Next PD
If DailyExport Then Exit Sub
PageHeader = Header & vbCrLf
PageHeader = PageHeader & "Period " & BeginDate & " To " & EndDate
   NetLaborSales = GetNetSales(BeginDate, EndDate)
   If Check_Option(96) = 0 And PreviewRecords = False Then
     Printer_Drivers "Receipt"
     PrinterPort = "Close Time Records"
     FP = 0
     PF = FreeFile
     Open PrinterPort For Output Shared As #PF
     CenterText = EC(8)
     LeftText = EC(9)
     'RightText = Chr(27) + Chr(97) + Chr(2)
     Print #PF, EC(0); EC(4) 'Int printer
     Print #PF, EC(1); EC(3); CenterText; StoreName; EC(2); EC(7)
     Print #PF, LeftText; EC(7)
     Print #PF, EC(1); EC(3); CenterText; PageHeader; EC(2); EC(7)
     Print #PF, LeftText; EC(7)
     Print #PF, Time;
     Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
     Print #PF, EC(3); String(40, 45); EC(7)
    Else
     SetPrinter 8
     PrintCLine StoreName, -1, 0
     'PrintCLine CStr(Heading), -1
     PrintCLine CStr(PageHeader), -1, 0
     PrintLine "Report Printed " & Format(Date, "mm/dd/yyyy") & " Printed By " & LoggedOnAs(0).EMPName, -1
     DrawLine
    End If
     For RCT = 0 To UBound(Employees, 1) ' - 1
       If Employees(RCT).Name = "" Then GoTo NXTEMP
       If Check_Option(96) = 0 And PreviewRecords = False Then
        Print #PF, "Name : "; Employees(RCT).Name; EC(7)
        Print #PF, "SSN  : "; Employees(RCT).Job; EC(7)
        Print #PF, String(40, 45); EC(7)
        Print #PF, "Work   Time    Brk    Brk   Time   Total"; EC(7)
        Print #PF, "Date    In     Out    In    Out    Hours"; EC(7)
        Print #PF, String(40, 45); EC(7)
       Else
        PrintLine "", -1
        PrintLine "Name  " & Employees(RCT).Name, 3500: PrintLine "SSN " & Employees(RCT).Job, -1
        If Val(Input_Option(55)) > 0 Then '// Option (55) = day overtime start
         PrintLine "Work Date", 1500: PrintLine "Time", 2500: PrintLine "Break Out", 3500: PrintLine "Break In", 4500: PrintLine "Out", 5300: PrintLine "Reg", 6000: PrintLine "Day OT", 6800: PrintLine "Total Hours", 8000: PrintLine "Tip Sales", 9000: PrintLine "T-Tips", 10000: PrintLine "Job", -1
        Else
         PrintLine "Work Date", 1500: PrintLine "Time", 2500: PrintLine "Break Out", 3500: PrintLine "Break In", 4500: PrintLine "Out", 5300: PrintLine "Reg", 6000: PrintLine "OT", 6800: PrintLine "Total Hours", 8000: PrintLine "Tip Sales", 9000: PrintLine "T-Tips", 10000: PrintLine "Job", -1
        End If
        DrawLine
       End If
        CurrentWeek = -1
        SplitShift = "": ShiftDept = "": SplitShiftNo = 0
        'moved here
         On Error Resume Next
         For CT = 0 To UBound(TimeRcd, 1) - 1
          ' If Err.Number > 0 Then Exit Sub
           On Error GoTo ErrorHandler
           With TimeRcd(CT)
            If TimeRcd(CT).Name = Employees(RCT).Name Then
            'If InStr(Employees(RCT).Name, "Andrew") > 0 Then Stop
            If .TimeON = .TimeOut Then
               InOutIsSame = True
            Else
               InOutIsSame = False
            End If
            If SplitShift <> .ShiftDate Or .ShiftNo <> SplitShiftNo And InOutIsSame = False Then
             PayRollData.CashTips = PayRollData.CashTips + .CashTip: PayRollData.CreditCardTips = PayRollData.CreditCardTips + .CCTip: PayRollData.TipsPaidOut = PayRollData.TipsPaidOut + .TipOut
             PayRollData.TipShare = PayRollData.TipShare + .TipShare: PayRollData.GratutityCollected = GetServerGratuity(Employees(RCT).Name, BeginDate, EndDate, 0)
             PayRollData.TipSales = PayRollData.TipSales + .Sales
            Else
               .Sales = 0:  .CCTip = 0
              If .ShiftNo = 0 Then
                PayRollData.TipShare = PayRollData.TipShare + .TipShare
              End If
              PayRollData.CashTips = PayRollData.CashTips + .CashTip
              'PayRollData.TipsPaidOut = PayRollData.TipsPaidOut + .TipOut
            End If
            If InOutIsSame = False Then
             SplitShift = .ShiftDate
             SplitShiftNo = .ShiftNo
            End If
            TimeEdits(0) = "": TimeEdits(1) = "": TimeEdits(2) = "": TimeEdits(3) = ""
            If .MGR > "" Then
             If .TOnAlt = 1 Then TimeEdits(0) = "*" Else TimeEdits(0) = ""
             If .BOAlt = 1 Then TimeEdits(1) = "*" Else TimeEdits(1) = ""
             If .BIAlt = 1 Then TimeEdits(2) = "*" Else TimeEdits(2) = ""
             If .TOntAlt = 1 Then TimeEdits(3) = "*" Else TimeEdits(3) = ""
            End If
            If Check_Option(96) = 0 And PreviewRecords = False Then
              If CurrentWeek <> TimeRcd(CT).PayWeek And CurrentWeek <> -1 Then
               Print #PF, String(40, 45); EC(7)
              End If
              If IsDate(.BkOut) Then
               Print #PF, Format(.ShiftDate, "mm/dd"); Tab(7); Format(.TimeON, "H:MM"); Tab(14); Format(.BkOut, "H:MM"); Tab(21); Format(.BkIn, "H:MM"); Tab(28); Format(.TimeOut, "H:MM"); Tab(34); Align(6, .RegHrs); EC(7)
              Else
               Print #PF, Format(.ShiftDate, "mm/dd"); Tab(7); Format(.TimeON, "H:MM"); Tab(14); ""; Tab(21); ""; Tab(28); Format(.TimeOut, "H:MM"); Tab(34); Align(6, .RegHrs); EC(7)
              End If
              If Not Break.Item(.RecordId) Is Nothing Then
               BKMINS = CalculateBreakTime(Break.Item(.RecordId).BreakOut, Break.Item(.RecordId).BreakIn, BreakMins)
               Print #PF, Break.Item(.RecordId).Job; Tab(7); ""; Tab(14); Format(Break.Item(.RecordId).BreakOut, "H:MM"); Tab(21); Format(Break.Item(.RecordId).BreakIn, "H:MM"); Tab(28); ""; Tab(34); Align(6, .RegHrs); EC(7)
              End If
              If .MGR > "" Then
               Print #PF, "Alt"; Tab(9); TimeEdits(0); Tab(14); TimeEdits(1); Tab(21); TimeEdits(2); Tab(30); TimeEdits(3); Tab(34); .MGR; EC(7)
              End If
              If Val(Input_Option(55)) > 0 And .DayOt > 0 Then Print #PF, .Job; "Day OT " & .DayOt; EC(7); ""
              If .EMPAP = "Not Approved" Then '// Option (55) = day overtime start
               Print #PF, "This Shift Was Not Approved "; EC(7)
              End If
            Else
             ShiftTips = ((.CashTip + .CCTip) + .TipShare) - .TipOut
             If Val(Input_Option(55)) > 0 Then '// Option (55) = day overtime start
               If .DayOt > 0 Then
                 If .RegHrs - .DayOt < 0 Then
                   DayOtStr = .RegHrs - .DayOt
                   .DayOt = Val(DayOtStr) + .DayOt
                   DayOtStr = 0
                 Else
                   DayOtStr = .RegHrs - .DayOt
                 End If
               Else
                  DayOtStr = .RegHrs
               End If
               If IsDate(.BkOut) Then
                 PrintLine Format(.ShiftDate, "mm/dd"), 1500: PrintLine Format(.TimeON, "H:MM") & TimeEdits(0), 2500: PrintLine Format(.BkOut, "H:MM") & TimeEdits(1), 3500: PrintLine Format(.BkIn, "H:MM") & TimeEdits(2), 4500: PrintLine Format(.TimeOut, "H:MM") & TimeEdits(3), 5300:  PrintLine Format(DayOtStr, "0.00"), 6000: PrintLine Format(.DayOt, "0.00"), 7000: PrintLine Format(.RegHrs, "0.00"), 8000: PrintLine Format(.Sales, "0.00"), 9000: PrintLine Format(ShiftTips, "0.00"), 10000: PrintLine .Job, -1
               Else
                 PrintLine Format(.ShiftDate, "mm/dd"), 1500: PrintLine Format(.TimeON, "H:MM") & TimeEdits(0), 2500: PrintLine "" & TimeEdits(1), 3500: PrintLine "" & TimeEdits(2), 4500: PrintLine Format(.TimeOut, "H:MM") & TimeEdits(3), 5300:   PrintLine Format(DayOtStr, "0.00"), 6000: PrintLine Format(.DayOt, "0.00"), 7000: PrintLine Format(.RegHrs, "0.00"), 8000: PrintLine Format(.Sales, "0.00"), 9000: PrintLine Format(ShiftTips, "0.00"), 10000: PrintLine .Job, -1
               End If
             Else
               If IsDate(.BkOut) Then
                PrintLine Format(.ShiftDate, "mm/dd"), 1500: PrintLine Format(.TimeON, "H:MM") & TimeEdits(0), 2500: PrintLine Format(.BkOut, "H:MM") & TimeEdits(1), 3500: PrintLine Format(.BkIn, "H:MM") & TimeEdits(2), 4500: PrintLine Format(.TimeOut, "H:MM") & TimeEdits(3), 5300:  PrintLine Format(.RegHrs - .OTHrs, "0.00"), 6000: PrintLine Format(.OTHrs, "0.00"), 7000: PrintLine Format(.RegHrs, "0.00"), 8000: PrintLine Format(.Sales, "0.00"), 9000: PrintLine Format(ShiftTips, "0.00"), 10000: PrintLine .Job, -1
               Else
                PrintLine Format(.ShiftDate, "mm/dd"), 1500: PrintLine Format(.TimeON, "H:MM") & TimeEdits(0), 2500: PrintLine "" & TimeEdits(1), 3500: PrintLine "" & TimeEdits(2), 4500: PrintLine Format(.TimeOut, "H:MM") & TimeEdits(3), 5300:   PrintLine Format(.RegHrs - .OTHrs, "0.00"), 6000: PrintLine Format(.OTHrs, "0.00"), 7000: PrintLine Format(.RegHrs, "0.00"), 8000: PrintLine Format(.Sales, "0.00"), 9000: PrintLine Format(ShiftTips, "0.00"), 10000: PrintLine .Job, -1
              End If
             End If
             If Not Break.Item(.RecordId) Is Nothing Then
               BKMINS = CalculateBreakTime(Break.Item(.RecordId).BreakOut, Break.Item(.RecordId).BreakIn, BreakMins)
               PrintLine Break.Item(.RecordId).Job, 1500: PrintLine "", 2500: PrintLine Format(Break.Item(.RecordId).BreakOut, "H:MM"), 3500: PrintLine Format(Break.Item(.RecordId).BreakIn, "H:MM"), 4500: PrintLine "", 5300: PrintLine Format(BKMINS, "0.00"), -1
             End If
             If .MGR > "" Then
               PrintLine "Shift Edit By " & .MGR, -1
             End If
             If .EMPAP = "Not Approved" Then
               PrintLine "Shift was not approved ", -1
             End If
            End If
            CurrentWeek = TimeRcd(CT).PayWeek
            End If
           End With
         Next CT
         If Check_Option(96) = 0 And PreviewRecords = False Then
            Print #PF, String(40, 45); EC(7)
            Print #PF, "Dept Name               Reg Hr     OT Hr"; EC(7)
            Print #PF, String(40, 45); EC(7)
         Else
            DrawLine
            PrintLine "Department Name", 4500: PrintLine "Regular Hours", 6500: PrintLine "Over Time", -1
            DrawLine
         End If
         EmpTotalRegHrs = 0: EmpTotalOtHrs = 0: GrossPay = 0: TotalDayOtHrs = 0
         For CT = 1 To CHRS.count
           If Employees(RCT).Name = CHRS.Item(CT).EMPName Then
              If Check_Option(96) = 0 And PreviewRecords = False Then
                 If CHRS.Item(CT).DayOt > 0 Then
                  Print #PF, CHRS.Item(CT).Job; Tab(25); Align(5, (CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs) - CHRS.Item(CT).DayOt); Tab(34); Align(5, CHRS.Item(CT).DayOt + CHRS.Item(CT).OTHrs); EC(7)
                 Else
                  Print #PF, CHRS.Item(CT).Job; Tab(25); Align(5, CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs); Tab(34); Align(5, CHRS.Item(CT).OTHrs); EC(7)
                 End If
              Else
                 If CHRS.Item(CT).DayOt > 0 Then
                   PrintLine CHRS.Item(CT).Job, 4500: PrintLine Format(((CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs) - CHRS.Item(CT).DayOt), "0.00"), 6500: PrintLine Format(CHRS.Item(CT).OTHrs + CHRS.Item(CT).DayOt, "0.00"), -1
                 Else
                  PrintLine CHRS.Item(CT).Job, 4500: PrintLine Format(CHRS.Item(CT).RegHrs - CHRS.Item(CT).OTHrs, "0.00"), 6500: PrintLine Format(CHRS.Item(CT).OTHrs, "0.00"), -1
                 End If
              End If
              GrossPay = GrossPay + CHRS.Item(CT).Pay
              EmpTotalRegHrs = EmpTotalRegHrs + CHRS.Item(CT).RegHrs
              EmpTotalOtHrs = EmpTotalOtHrs + CHRS.Item(CT).OTHrs
              TotalDayOtHrs = TotalDayOtHrs + CHRS.Item(CT).DayOt
           End If
         Next CT
         If Check_Option(96) = 0 And PreviewRecords = False Then
          Print #PF, String(40, 45); EC(7)
          Print #PF, "Total Hours"; Tab(25); Align(5, ((EmpTotalRegHrs - EmpTotalOtHrs) - TotalDayOtHrs)); Tab(34); Align(5, EmpTotalOtHrs + TotalDayOtHrs); EC(7)
         Else
          DrawLine
          PrintLine "Total Hours", 4500: PrintLine Format(((EmpTotalRegHrs - EmpTotalOtHrs) - TotalDayOtHrs), "0.00"), 6500: PrintLine Format(EmpTotalOtHrs + TotalDayOtHrs, "0.00"), -1
         End If
         

         PayRollData.NetTips = ((((PayRollData.CashTips + PayRollData.CreditCardTips) + PayRollData.TipShare) + PayRollData.GratutityCollected) - PayRollData.TipsPaidOut)
         PercOfTip = 0
         If PayRollData.NetTips > 0 And PayRollData.TipSales > 0 Then
          PercOfTip = Format((PayRollData.NetTips / PayRollData.TipSales) * 100, "0.00")
         End If
         If Check_Option(96) = 0 And PreviewRecords = False Then
          Print #PF, String(40, 45); EC(7)
          Print #PF, ; EC(7)
          Print #PF, "Sign :_________________________________"; EC(7)
          Print #PF, "I here by acknowledge that this is a "; EC(7)
          Print #PF, "true and complete record of my time. "; EC(7)
          Print #PF, ; EC(7)
          Print #PF, String(40, 45); EC(7)
         Else
          DrawLine
          PrintLine "", -1
          PrintLine "Sign :_________________________________ I here by acknowledge that this is a true and complete record of my time. ", -1
          PrintLine "", -1
          PrintLine "", -1
          DrawLine
         End If
         'If (TotalTips + PayRollData.GratutityCollected) > 0 And PayRollData.TipSales > 0 Then
         ' PercOfTip = Format(((TotalTips + PayRollData.GratutityCollected) / PayRollData.TipSales) * 100, "0.00")
         'End If
         If Check_Option(96) = 0 And PreviewRecords = False Then
          Print #PF, "Summary "; EC(7)
          Print #PF, String(40, 45); EC(7)
          If LoggedOnAs(0).EmpAccessLevel < Access_Required("MASK RATES") Then
           Print #PF, "Sales " & Format(PayRollData.TipSales, "0.00"); Tab(25); "Gross Pay " & "XXX.XX"; EC(7)
          Else
           Print #PF, "Sales " & Format(PayRollData.TipSales, "0.00"); Tab(25); "Gross Pay " & Format(GrossPay, "0.00"); EC(7)
          End If
          If PayRollData.CashTips > 0 Then
           Print #PF, "Cash Tips Declared " & Format(PayRollData.CashTips, "0.00"); EC(7)
          End If
          If PayRollData.CreditCardTips > 0 Then
           Print #PF, "Credit Card Tips   " & Format(PayRollData.CreditCardTips, "0.00"); EC(7)
          End If
          If PayRollData.GratutityCollected > 0 Then
           Print #PF, "Gratuity Charged   " & Format(PayRollData.GratutityCollected, "0.00"); EC(7)
          End If
          If PayRollData.TipShare > 0 Then
           Print #PF, "Tip Share Received " & Format(PayRollData.TipSales, "0.00"); EC(7)
          End If
          If PayRollData.TipsPaidOut > 0 Then
           Print #PF, "Tips Paid Out      " & Format(PayRollData.TipsPaidOut, "0.00"); EC(7)
           Print #PF, "(deducted from tip total)"; EC(7)
          End If
          If PayRollData.NetTips > 0 Or PercOfTip > 0 Then
           Print #PF, "Total Tips  " & Format(PayRollData.NetTips, "0.00"); Tab(23); "Tip %   " & Format(PercOfTip, "0.00"); EC(7)
          End If
          Print #PF, String(40, 45); EC(7)
         Else
          DrawLine
          If LoggedOnAs(0).EmpAccessLevel < Access_Required("MASK RATES") Then
           PrintLine "Summary ", 2500: PrintLine "Gross Pay " & "XXX.XX", 4500: PrintLine "Tip Sales  " & Format(PayRollData.TipSales, "0.00"), 7500: PrintLine "% Sales " & Format(PercOfTip, "0.00"), -1
          Else
           PrintLine "Summary ", 2500: PrintLine "Gross Pay " & Format(GrossPay, "0.00"), 4500: PrintLine "Tip Sales  " & Format(PayRollData.TipSales, "0.00"), 7500: PrintLine "% Sales " & Format(PercOfTip, "0.00"), -1
          End If
          If PayRollData.NetTips <> 0 Then
           PrintLine "Summary Tips ", 2500: PrintLine "Declared " & Format(PayRollData.CashTips, "0.00"), 4500: PrintLine "Gratuity Charged " & Format(PayRollData.GratutityCollected, "0.00"), 6500: PrintLine "cc Tips " & Format(PayRollData.CreditCardTips, "0.00"), 8000: PrintLine "Tips Paid " & Format(PayRollData.TipsPaidOut, "0.00"), 9500: PrintLine "Net Total Tips " & Format(PayRollData.NetTips, "0.00"), -1
          End If
          If GetConfigInfo("PayRoll", "Separate Pages for Time Records", "No", "Functions.INI") = "ENABLED" Then ReportView.View.NewPage Else DrawLine
         End If
         PercOfTip = 0: GrossPay = 0
         PayRollData.NetTips = 0: PayRollData.CashTips = 0: PayRollData.CreditCardSales = 0
         PayRollData.CreditCardTips = 0: PayRollData.DeclaredTips = 0: PayRollData.GratutityCollected = 0
         PayRollData.TipsPaidOut = 0: PayRollData.TipShare = 0: PayRollData.TipSales = 0: GrossPay = 0
NXTEMP:
        Next RCT
        If Check_Option(96) = 0 And PreviewRecords = False Then
         Print #PF, "Totals                  Reg Hr     OT Hr"; EC(7)
         Print #PF, String(40, 45); EC(7)
        Else
         DrawLine
         PrintLine "Totals", 3500: PrintLine "Regular Hours", 5500: PrintLine "Over Time", 7500: PrintLine "Dollars", 9500: PrintLine "Percentage", -1
         DrawLine
        End If
        TotalRegHrs = 0: TotalOtHrs = 0: TotalDayOtHrs = 0
        Dim TotalDollars As Currency
        TotalDollars = 0
        For RCT = 1 To PayRollSummary.count
          For PD = 1 To CHRS.count
            If PayRollSummary.Item(RCT).Job = CHRS.Item(PD).Job Then
              TotalRegHrs = TotalRegHrs + CHRS.Item(PD).RegHrs
              TotalOtHrs = TotalOtHrs + (CHRS.Item(PD).OTHrs + CHRS.Item(PD).DayOt)
              TotalDayOtHrs = TotalDayOtHrs + CHRS.Item(PD).DayOt
              TotalDollars = TotalDollars + CHRS.Item(PD).Pay
              PayRollSummary.Item(RCT).RegHrs = PayRollSummary.Item(RCT).RegHrs + CHRS.Item(PD).RegHrs
              PayRollSummary.Item(RCT).OTHrs = PayRollSummary.Item(RCT).OTHrs + CHRS.Item(PD).OTHrs
              PayRollSummary.Item(RCT).DayOt = PayRollSummary.Item(RCT).DayOt + CHRS.Item(PD).DayOt
              PayRollSummary.Item(RCT).Pay = PayRollSummary.Item(RCT).Pay + CHRS.Item(PD).Pay
            End If
          Next PD
        Next RCT
        For PD = 1 To PayRollSummary.count
         If Check_Option(96) = 0 And PreviewRecords = False Then
           Print #PF, PayRollSummary.Item(PD).Job; Tab(25); Align(5, (PayRollSummary.Item(PD).RegHrs - PayRollSummary.Item(PD).OTHrs) - PayRollSummary.Item(PD).DayOt); Tab(34); Align(5, Val(Check_Null(PayRollSummary.Item(PD).OTHrs)) + Val(Check_Null(PayRollSummary.Item(PD).DayOt))); EC(7)
           
           If PayRollSummary.Item(PD).DayOt > 0 Then
            Print #PF, "Regular Over Time"; Tab(34); Align(5, PayRollSummary.Item(PD).OTHrs); EC(7)
            Print #PF, "Day Over Time Used"; Tab(34); Align(5, PayRollSummary.Item(PD).DayOt); EC(7)
           End If
           If NetLaborSales > 0 And PayRollSummary.Item(PD).Pay > 0 Then
            Print #PF, "Dollars Spent"; Tab(25); Align(5, PayRollSummary.Item(PD).Pay); Tab(34); Align(5, (PayRollSummary.Item(PD).Pay / NetLaborSales) * 100); EC(7)
           End If
         Else
          If NetLaborSales > 0 And PayRollSummary.Item(PD).Pay > 0 Then
           PrintLine PayRollSummary.Item(PD).Job, 3500: PrintLine Format((PayRollSummary.Item(PD).RegHrs - PayRollSummary.Item(PD).OTHrs) - PayRollSummary.Item(PD).DayOt, "0.00"), 5500: PrintLine Format(Val(PayRollSummary.Item(PD).OTHrs & "") + Val(PayRollSummary.Item(PD).DayOt & ""), "0.00"), 7500: PrintLine Format(PayRollSummary.Item(PD).Pay, "0.00"), 9500: PrintLine Format((PayRollSummary.Item(PD).Pay / NetLaborSales) * 100, "0.00"), -1
          
          Else
           PrintLine PayRollSummary.Item(PD).Job, 3500: PrintLine Format((PayRollSummary.Item(PD).RegHrs - PayRollSummary.Item(PD).OTHrs) - PayRollSummary.Item(PD).DayOt, "0.00"), 5500: PrintLine Format(Val(PayRollSummary.Item(PD).OTHrs & "") + Val(PayRollSummary.Item(PD).DayOt & ""), "0.00"), 7500: PrintLine Format(PayRollSummary.Item(PD).Pay, "0.00"), -1 ': PrintLine Format((PayRollSummary.Item(PD).Pay / NetLaborSales) * 100, "0.00"), -1
          End If
          If PayRollSummary.Item(PD).DayOt > 0 Then
           PrintLine "Regular Over Time Used", 5500:  PrintLine Format(PayRollSummary.Item(PD).OTHrs, "0.00"), -1
           PrintLine "Day Over Time Used", 5500:  PrintLine Format(PayRollSummary.Item(PD).DayOt, "0.00"), -1
          End If
         End If
        Next PD
        If Check_Option(96) = 0 And PreviewRecords = False Then
        
         Print #PF, "Total Hours"; Tab(25); Align(5, TotalRegHrs - TotalOtHrs); Tab(34); Align(5, TotalOtHrs); EC(7)
         If TotalDayOtHrs > 0 Then
          Print #PF, "Total Day Over Time"; Tab(34); Align(5, TotalDayOtHrs); EC(7)
         End If
         If NetLaborSales > 0 And TotalDollars > 0 Then
          Print #PF, "Total Dollars"; Tab(25); Align(5, TotalDollars); Tab(34); Align(5, (TotalDollars / NetLaborSales) * 100); EC(7)
         End If
        Else
          If NetLaborSales > 0 And TotalDollars > 0 Then
           DrawLine
           PrintLine "Total Hours Used   ", 4500: PrintLine Format(TotalRegHrs, "0.00"), -1
           PrintLine "Total Regular Over Time Used   ", 4500: PrintLine Format(TotalOtHrs, "0.00"), -1
           If TotalDayOtHrs > 0 Then
           PrintLine "Total Day Over Time Used   ", 4500: PrintLine Format(TotalDayOtHrs, "0.00"), -1
           End If
           PrintLine "Total Dollars Spent", 4500: PrintLine Format(TotalDollars, "0.00"), -1
           PrintLine "Labor Percentage", 4500: PrintLine Format((TotalDollars / NetLaborSales) * 100, "0.00"), -1
           PrintLine "Net Sales", 4500: PrintLine Format(NetLaborSales, "0.00"), -1
           DrawLine
          End If
        End If
        If Check_Option(96) = 0 And PreviewRecords = False Then
         Print #PF, String(40, 45); EC(7)
         Print #PF, EC(6) & EC(6); EC(5)
         Close #PF
         SendJobToPrinter PrinterPort, "Receipt", False
        Else
         DrawLine
         ReportView.View.EndDoc
         'If SendEmails(10).Send Then
         '  E_MailReport "Payroll Report"
         '  If Header = "EMail Time Records" Then Exit Sub
         'End If
         ReportView.Show 1
        End If
Set WeeksWorked = Nothing
Set CHRS = Nothing
Set PayRollSummary = Nothing
Set DayOvrTime = Nothing
Set ROT = Nothing
If Check_Option(61) = 1 Then
  'SaveSetting "Micro$ale", "Register Setup", "Export File Name", ExportFile.ExportFileName
  If GetSetting("Micro$ale", "Register Setup", "Export File Name") > "" Then  'NEED THIS FOR ADP export file name changes
    CustOption.PayRollFileName = GetSetting("Micro$ale", "Register Setup", "Export File Name")
    SaveSetting "Micro$ale", "Register Setup", "Export File Name", ""
  Else
    CustOption.PayRollFileName = GetConfigInfo("Exports", "Export File Name", "", "PayRollExport.INI")
  End If
  If CustOption.PayRollFileName > "" Then
    EmailFiles GetConfigInfo("Exports", "Export Location", "", "PayRollExport.INI"), CustOption.PayRollFileName, GetConfigInfo("Exports", "EMail File To", "", "PayRollExport.INI"), "PayRollExport", "Pay Roll Exports"
  End If
  If Dir(ExportFileTo & CustOption.PayRollFileName) > "" Then
     DeleteFile ExportFileTo & CustOption.PayRollFileName
  End If
End If
Call MemoryCheck(Time_Sheet)
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Print Payroll ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub HistoryOn_Click()

End Sub

Private Sub List2_DblClick()
On Error Resume Next
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim wkStartDate As String
Dim wkEndDate   As String
ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
Sql = "Select * from [Employee Records] "
Sql = Sql & "order by TRY_CAST([Shift Date] AS datetime) asc"
ConnectADORS Sql, Db, RS
If Not RS.EOF Then
    If Not RS.EOF Then RS.MoveLast
    wkEndDate = RS![Shift date] & ""
    If Not RS.BOF Then RS.MoveFirst
    wkStartDate = RS![Shift date] & ""
    DbClose Db, RS
    See_Msg "Period Start " & wkStartDate & " Period End " & wkEndDate, 0, "Employee Work Records"
Else
    DbClose Db, RS
    See_Msg "No Dates Found", 5, "No Records Found"
End If
End Sub
Private Sub LoadBox()
Dim TRDB As ADODB.Connection
Dim TRRS As ADODB.Recordset
OpenDb SvrPath.TerminalPath & TimeRecordsDb, TRDB
 Sql = "Select Distinct TRY_CAST([Shift Date] AS date) as MyDate from [Employee Records]" 'order by TRY_CAST([Shift Date] AS datetime) asc"
 Sql = Sql & " Group By TRY_CAST([Shift Date] AS date) Order by TRY_CAST([Shift Date] AS date) desc"
ConnectADORS Sql, TRDB, TRRS
If TRRS.EOF Then
   See_Msg "No Records Found", 5, ""
   DbClose TRDB, TRRS
   Exit Sub
End If
DateList.Clear
  Do While Not TRRS.EOF
    DateList.AddItem Format(TRRS.fields("MyDate").Value, "mm/dd/yyyy")
    TRRS.MoveNext
  Loop
DbClose TRDB, TRRS
End Sub
Private Sub SSCommand1_Click()
End Sub
Private Sub SSCommand11_Click()
If Can_Continue(SSCommand11.Caption) = False Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
On Error GoTo ErrorHandler
Dim BRKMINS As Double
Dim PRate As Currency
Dim PrinterPort As String
Dim Heading As String
Dim PF As Integer
Dim getweek As Integer
Dim breaktotal As Integer
Dim BeginWeek As String
Dim EndingDate As String
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
 Exit Sub
Else
 BeginWeek = SearchDates.Item(1).StartDate
 EndingDate = SearchDates.Item(1).EndDate
End If
Heading = "Break Report"
BRKMINS = 0
SetPage ReportView.View, Heading
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
Sql = "Select * From [Employee Records] Where TRY_CAST([Shift Date] AS date) between '" & BeginWeek & "' and '" & EndingDate & "'"
Sql = Sql & " order By TRY_CAST([Shift Date] AS date), [Employee Name] asc"

ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
      PrintLine "Name", 3500
      PrintLine "Work Date", 5500
      PrintLine "Break Out", 7500
      PrintLine "Break In", 9500
      PrintLine "Total Mins", -1
  End If
  Do While Not Tb.EOF
   If Trim(Tb(4)) > "" And Tb(4) <> "00:00" Then
   If SearchName <> Tb![employee name] Then
      DrawLine
      SearchName = Tb![employee name]
      PrintLine SearchName, -1
      DrawLine
   End If
   If Trim(Tb(5)) > "" And Tb(5) <> "00:00" And Trim(Tb(4)) > "" And Tb(4) <> "00:00" Then
   breaktotal = DateDiff("n", CDate(Tb(4)), CDate(Tb(5)), WeekStartsOn, 1)
   End If
      PrintLine "", 3500
      PrintLine Format(Tb(2), "mm/dd"), 5500
      PrintLine Format(Tb(4), "hh:mm ampm"), 7500
      PrintLine Format(Tb(5), "HH:MM ampm"), 9500
      PrintLine Format(breaktotal, "####"), -1
      BRKMINS = BRKMINS + Val(breaktotal)
  End If
Tb.MoveNext
Loop
DrawLine
PrintLine "Total Break Mins", 9500: PrintLine Format(BRKMINS, "####"), -1
FinishJob 1
If Check_Option(96) = 1 Then Exit Sub
If See_Msg("Print Report", 4, "") <> 6 Then Exit Sub
Printer_Drivers "Receipt"
PF = FreeFile
PrinterPort = "Break Report"
OneMoreTime:
Open PrinterPort For Output Shared As #PF
CenterText = EC(8)
LeftText = EC(9)
'RightText = Chr(27) + Chr(97) + Chr(2)
FP = 0
Print #PF, EC(0); EC(4) 'Int printer
Print #PF, EC(1); EC(3); CenterText; StoreName; EC(2); EC(7)
Print #PF, LeftText; EC(7)
Print #PF, EC(1); EC(3); CenterText; Heading; EC(2); EC(7)
Print #PF, ; EC(9); EC(7)
Print #PF, CenterText; ForWhen; LeftText; EC(2); EC(7)
Print #PF, LeftText; EC(7)
Print #PF, Time;
Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
Print #PF, EC(4); String(40, 45); EC(7)
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
 Sql = "Select * From [Employee Records]  order By TRY_CAST([Shift Date] AS datetime), [Employee Name] asc"
ConnectADORS Sql, Db, Tb
 Do While Not Tb.EOF
   If Trim(Tb(4)) > "" And Tb(4) <> "00:00" Then
   If SearchName <> Tb![employee name] Then
   SearchName = Tb![employee name]
   Print #PF,
   Print #PF, String(40, 45); EC(7)
   Print #PF, "Name : "; SearchName; EC(7)
   Print #PF, String(40, 45); EC(7)
   Print #PF, "Date      Brk Out    Brk In      Break"; EC(7)
   Print #PF, String(40, 45); EC(7)
   End If
   If Trim(Tb(5)) > "" And Tb(5) <> "00:00" And Trim(Tb(4)) > "" And Tb(4) <> "00:00" Then
   breaktotal = DateDiff("n", CDate(Tb(4)), CDate(Tb(5)), WeekStartsOn, 1)
   End If
   Print #PF, Format(Tb(2), "mm/dd"); Tab(10); Format(Tb(4), "HH:MM ampm"); Tab(21); Format(Tb(5), "HH:MM ampm"); Tab(34); Format(breaktotal, "0.00"); EC(7)
End If
NXTEMP:
Tb.MoveNext
Loop
Print #PF, String(40, 45); EC(7)
Print #PF, EC(6); EC(5)
Close
SendJobToPrinter PrinterPort, "Receipt", False
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Report SS1 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub

Private Sub SSCommand12_Click()

If Can_Continue(SSCommand12.Caption) = False Then Exit Sub
DoTimeReport = True
Dim BeginWeek As String
Dim EndingDate As String
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
 Exit Sub
Else
BeginWeek = SearchDates.Item(1).StartDate
EndingDate = SearchDates.Item(1).EndDate
End If
GetTimeReport BeginWeek, EndingDate
End Sub
Private Sub SSCommand13_Click()
If Can_Continue(SSCommand13.Caption) = False Then Exit Sub
On Error Resume Next
LoadBox
SelectReport = "Hourly Labor"
CenterObject Time_Reports, Reportday
Reportday.Visible = True
Reportday.ZOrder 0
End Sub
Private Sub HourlyLaborData(SDate As String)
If IsDate(SDate) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim PrinterPort As String
Dim CountHrs As Integer
Dim EachHour As String
Dim Db As ADODB.Connection
Dim PRTb As ADODB.Recordset
Dim EndTime As String
Dim HrsUse  As Double
Dim DidPrint As Boolean
Dim ProductionHrs As Single
Dim HrlyPerc As Single
Dim LbPec As Single
Dim LbDol As Single
Dim MHPro As Single
Dim TSale As Currency
Dim LSpt  As Currency
Dim DN As Integer
Dim DOW As String
If CDate(SDate) = CDate(Date) Then
  DN = 0
Else
  DN = GetDateToNo(SDate)
End If
Erase Hourly, MealPeriod
ReDim Hourly(1, 24, 3)
ReDim MealPeriod(1, 5, 4)
''''''''''''''''''''''''''''''
Dim Data$
''''''''''''''''''''''''''''''
If DN = 0 Then
 HourlySales Hourly(), MealPeriod(), MyRegisters(0).regName, DN
 DOW = 0
Else
  ReturnHourlySaleHistory SDate, Hourly()
DOW = 1
End If
ReportView.View.Orientation = orPortrait
SetPage ReportView.View, "Hourly Labor " & "For " & Format(SDate, "MM/DD/YY")
Dim F$
F = "+2500|+>2000|+>2000|+>1500|+>1500;"
Dim H$
H = "Hours Used|Spent|Sales|Labor %|Production|;"
Dim GetHourTime As String
Dim MinsWorked As Long
Dim EH As String
Dim BH As String
Dim CH As String
Dim BT As String
Data$ = ""
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
  Sql = "Select * From [Employee Records] Where "
  Sql = Sql & " TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(SDate) & " " & GetStartTime & "'"
  Sql = Sql & " and '" & SetEndingDate(CDate(SDate)) & " " & SetEndingTime & "'"
  Sql = Sql & " and [Paid By] <> 0"
  Sql = Sql & " order by TRY_CAST([time In] AS datetime) asc"
ConnectADORS Sql, Db, PRTb
ReDim PayRate(24)
For CountHrs = 0 To 23
  GetHourTime = Format(CStr(CountHrs) & ":00", "HH:MM ampm")
  If Not PRTb.BOF Then PRTb.MoveFirst
    Do While Not PRTb.EOF
       If Not IsDate(PRTb![Time Out]) Then
        EndTime = Time
       Else
        EndTime = PRTb![Time Out]
       End If
       BH = Hour(CDate(PRTb![Shift date]) & " " & CDate(PRTb![Time In])) '// begin hour
       EH = Hour(CDate(PRTb![Punch out Date]) & " " & CDate(EndTime))    '// end hour
       CH = Hour(CDate(PRTb![Punch out Date]) & " " & CDate(GetHourTime)) '// current hour
       If Val(EH) = Val(CH) And Val(BH) = Val(CH) Then
        MinsWorked = Abs((DateDiff("n", CDate(PRTb![Punch out Date] & " " & CDate(EndTime)), CDate(PRTb![Shift date]) & " " & CDate(PRTb![Time In]))))
        HrsUse = MinsWorked / 60
        If HrsUse < 0 Then HrsUse = 0 '// if end hour and begin hour = current hour
       ElseIf Val(EH) = Val(CH) Then
        MinsWorked = Abs((DateDiff("n", CDate(PRTb![Punch out Date] & " " & CDate(EndTime)), CDate(PRTb![Punch out Date]) & " " & CDate(GetHourTime))))
        HrsUse = MinsWorked / 60
        If HrsUse < 0 Then HrsUse = 0
       ElseIf Val(BH) = Val(CH) And Val(EH) >= Val(CH) Then
         MinsWorked = Abs((DateDiff("n", CDate(PRTb![Shift date] & " " & CDate(PRTb![Time In])), CDate(PRTb![Punch out Date]) & " " & CDate(Format(CStr(CountHrs + 1) & ":00", "HH:MM ampm")))))
         HrsUse = MinsWorked / 60
         If HrsUse < 1 Then
         
         Else
          HrsUse = 1
         End If
         If HrsUse < 0 Then HrsUse = 0
       ElseIf Val(BH) < Val(CH) And Val(EH) >= Val(CH) Then
        HrsUse = 1
       Else
        HrsUse = 0
       End If
       If HrsUse > 0 Then
        PayRate(CountHrs + DOW).JobCodes = Format(CStr(Hourly(0, CountHrs + DOW, 1)), "0.00")
        PayRate(CountHrs + DOW).JobRate = PayRate(CountHrs + DOW).JobRate + (Val(PRTb![Rate of Pay] & "") * HrsUse)
        PayRate(CountHrs + DOW).HoursUsed = PayRate(CountHrs + DOW).HoursUsed + HrsUse
       End If
       HrsUse = 0
      PRTb.MoveNext
    Loop
    'Debug.Print PayRate(CountHrs + DOW).HoursUsed
Next CountHrs
DbClose Db, PRTb
DidPrint = False
TSale = 0: HrlyPerc = 0: LSpt = 0: MHPro = 0
For CountHrs = 0 To 23
  EachHour = Format(CountHrs & ":00", "HH:MM ampm")
     If PayRate(CountHrs).HoursUsed > 0 Then
        DidPrint = True
        TSale = TSale + Val(PayRate(CountHrs).JobCodes)
        If Val(PayRate(CountHrs).HoursUsed) > 0 And Val(PayRate(CountHrs).JobCodes) > 0 Then
          ProductionHrs = Val(PayRate(CountHrs).JobCodes) / Val(PayRate(CountHrs).HoursUsed)
        Else
          ProductionHrs = 0
        End If
        If Val(PayRate(CountHrs).HoursUsed) > 0 And Val(PayRate(CountHrs).JobCodes) > 0 Then
          HrlyPerc = Val(PayRate(CountHrs).JobRate) / Val(PayRate(CountHrs).JobCodes)
        Else
          HrlyPerc = 0
        End If
      LbPec = LbPec + HrlyPerc
      MHPro = MHPro + Val(PayRate(CountHrs).HoursUsed)
      LSpt = LSpt + Val(PayRate(CountHrs).JobRate)
      'PrintLine , 4500: PrintLine , 6500: PrintLine , 8500: PrintLine , 10500: PrintLine , -1
      Data$ = Data$ & EachHour & " " & " Hours Used      " & Format(PayRate(CountHrs).HoursUsed, "0.00") & "|"
      Data$ = Data$ & Format(PayRate(CountHrs).JobRate, "0.00") & "|"
      Data$ = Data$ & Format(Val(PayRate(CountHrs).JobCodes & ""), "0.00") & "|"
      Data$ = Data$ & Format(HrlyPerc, "0.00%") & "|"
      Data$ = Data$ & Format(ProductionHrs, "0.00") & "|;"
    End If
Next CountHrs

If DidPrint Then
    'DrawLine
    Data$ = Data$ & "|"
    Data$ = Data$ & "|"
    Data$ = Data$ & "|"
    Data$ = Data$ & "|"
    Data$ = Data$ & "|;"
    If LSpt > 0 And TSale > 0 Then
     Data$ = Data$ & "Total Hours Used             " & Format(MHPro, "0.00") & "|"
     Data$ = Data$ & Format(LSpt, "0.00") & "|"
     Data$ = Data$ & Format(TSale, "0.00") & "|"
     Data$ = Data$ & Format((LSpt / TSale), "0.00%") & "|"
     Data$ = Data$ & Format((TSale / MHPro), "0.00") & "|;"
   End If
    ReportView.View.TableBorder = tbTopBottom
    ReportView.View.addTable F, H, Data$, vbBlack, vbWhite
    FinishJob 1
Else
    See_Msg "No Data Found", 5, ""
End If
If Check_Option(96) = 0 Then
    If See_Msg("Print Report ?", 4, "") <> 6 Then Exit Sub
    Printer_Drivers "Receipt"
    Dim PF As Integer
    PF = FreeFile
    PrinterPort = "Hourly Labor"
OneMoreTime:
    Open PrinterPort For Output Shared As #PF
    CenterText = EC(8)
    LeftText = EC(9)
    'RightText = Chr(27) + Chr(97) + Chr(2)
    TSale = 0: HrlyPerc = 0: LSpt = 0: MHPro = 0
    Print #PF, CenterText & "Hourly Labor" & LeftText
    Print #PF, EC(9); String(33, 45); EC(7)
    Print #PF, "Time        Hrs      $        %"; EC(7)
    Print #PF, String(33, 45); EC(7)
    For CountHrs = 0 To 23
     EachHour = Format(CountHrs & ":00", "HH:MM ampm")
     If PayRate(CountHrs).HoursUsed > 0 Then
        TSale = TSale + Val(PayRate(CountHrs).JobCodes)
        If Val(PayRate(CountHrs).HoursUsed) > 0 And Val(PayRate(CountHrs).JobCodes) > 0 Then
          ProductionHrs = Val(PayRate(CountHrs).JobCodes) / Val(PayRate(CountHrs).HoursUsed)
        Else
          ProductionHrs = 0
        End If
        If Val(PayRate(CountHrs).HoursUsed) > 0 And Val(PayRate(CountHrs).JobCodes) > 0 Then
          HrlyPerc = Val(PayRate(CountHrs).JobRate) / Val(PayRate(CountHrs).JobCodes)
        Else
          HrlyPerc = 0
        End If
      LbPec = LbPec + HrlyPerc
      MHPro = MHPro + Val(PayRate(CountHrs).HoursUsed)
      LSpt = LSpt + Val(PayRate(CountHrs).JobRate)
      Print #PF, EachHour; Tab(13); Format(PayRate(CountHrs).HoursUsed, "0.00"); Tab(20); Format(PayRate(CountHrs).JobCodes, "0.00"); Tab(30); Format(HrlyPerc, "0.00%"); EC(7)
      'Format(ProductionHrs, "0.00")
    End If
   Next CountHrs
    If LSpt > 0 And TSale > 0 Then
      HrlyPerc = LSpt / TSale
    End If
    Print #PF, String(33, 45); EC(7)
    Print #PF, "Used "; Tab(10); Format(LSpt, "$0.00"); Tab(20); Format(TSale, "0.00"); Tab(30); Format((HrlyPerc), "0.00%"); EC(7)
    Print #PF, EC(6) & EC(5); EC(7)
   End If
   Close #PF
   SendJobToPrinter PrinterPort, "Receipt", False

Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Hourly Labor Data ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub ReturnHourlySaleHistory(HistoryDate As String, Hourly() As Variant)
On Error GoTo ErrorHandler
Dim Db      As ADODB.Connection
Dim OP      As ADODB.Recordset
Dim NetSale As Currency
ConnectADODB ProductHistory, Db
 Sql = "select * from [Detail] where TRY_CAST([Date] AS datetime) Between '" & Format(CDate(CDate(HistoryDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(HistoryDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'"
 Sql = Sql & " Order by [date] asc"
ConnectADORS ReturnDateString(Sql), Db, OP
         Do While Not OP.EOF
           If OP![Department Name] <> "Gift Certificate" _
             And OP![Department Name] <> "Gratuity'" _
             And OP![Department Name] <> "Service Charge" _
             And OP![Department Name] <> "Paid In" _
             And OP![Department Name] <> "Gift EPay New Issue" _
             And OP![Department Name] <> "Gift EPay ReLoad" _
             And OP![Department Name] <> "Account Payment" Then

                If Check_Option(91) = 1 Then
                  NetSale = Abs(Val(OP![Item Price] & ""))
                Else
                  NetSale = Val(OP![Item Price] & "")
                End If
                HR = Hour(OP![Time]) + 1
                If HR = 24 Then HR = 0  'Set 24 = to Midnight
                Hourly(0, HR, 0) = HR & ":00"
                Hourly(0, HR, 1) = Hourly(0, HR, 1) + Val(NetSale) 'Net Sales
                Hourly(0, HR, 2) = Val(Hourly(0, HR, 2)) + 1 'Transaction count
                Hourly(0, HR, 3) = Val(Hourly(0, HR, 3)) + 1 'Transaction count
           End If
     OP.MoveNext
   Loop
DbClose Db, OP
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Hourly Labor Data ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub SSCommand2_Click()
If Can_Continue(SSCommand2.Caption) = False Then Exit Sub
Dim BeginWeek As String
Dim EndingDate As String
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
 Exit Sub
Else
BeginWeek = SearchDates.Item(1).StartDate
EndingDate = SearchDates.Item(1).EndDate
End If
Sleep 0
HouseLabor BeginWeek, EndingDate
Sleep 0
End Sub
Private Sub SSCommand3_Click()
End Sub
Public Function GetPunchRecordsDates(EndDate As String) As String
Dim Db                As ADODB.Connection
Dim Tb                As ADODB.Recordset
If ConnectADODB(SvrPath.TerminalPath & TimeRecordsDb, Db) = False Then Exit Function
   Sql = "Select * from [Employee Records] "
   Sql = Sql & " order by TRY_CAST([Shift Date] AS datetime) asc"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
GetPunchRecordsDates = ""
EndDate = ""
If Not Tb.EOF Then
  If Not Tb.BOF Then Tb.MoveFirst
     GetPunchRecordsDates = Format(Tb.fields("Shift Date").Value, "mm/dd/yyyy")
  If Not Tb.EOF Then Tb.MoveLast
  EndDate = Format(Tb.fields("Shift Date").Value, "mm/dd/yyyy")
End If
DbClose Db, Tb
Exit Function
End Function
Public Sub SSCommand4_Click()
If Can_Continue(SSCommand4.Caption) = False Then Exit Sub
PunchReport TimeRecordsDb
End Sub
Private Sub PunchReport(DatabaseName As String)
On Error GoTo ErrorHandler
Dim Hoursworked As String
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim RS          As ADODB.Recordset
Dim BRKMINS     As Single
Dim PRate       As Currency
Dim PrinterPort As String, Heading As String
Dim PF          As Integer
Dim getweek     As Integer
Dim Total_Time  As Single
Dim TTime       As Single
Dim shiftTotal  As Single
Dim PayRate     As Currency
Dim DollarSpent As Currency
Dim IndHours    As Single
Dim OTPay       As Currency
Dim breaktotal  As Single
Dim EmpHours    As Single
Dim EmPNames()  As ALLEMPTime
Dim FindEmp     As Integer
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
  Exit Sub
End If
If GetEmpLastName(SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, EmPNames) = False Then
   See_Msg "No Records Found", 5, ""
   Exit Sub
End If
Heading = "Punch Report": BRKMINS = 0: DollarSpent = 0
SetPage ReportView.View, Heading
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
   PrintLine "Name", 2000
   PrintLine "Work Date", 3000
   PrintLine "Time In", 4000
   PrintLine "Break Out", 5000
   PrintLine "Break In", 6000
   PrintLine "Time Out", 7000
   PrintLine "Total Time", 8000
   PrintLine "Dollar", 9000
   PrintLine "Sales", 10000
   PrintLine "Tips", -1
   For FindEmp = 0 To UBound(EmPNames, 1)
    If EmPNames(FindEmp).WholeName > "" Then
        Sql = "Select * from [Employee Records] Where [employee Name] = '" & EmPNames(FindEmp).WholeName & "'"
        Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(SearchDates.Item(1).StartDate) & " " & GetStartTime & "'"
        Sql = Sql & " and '" & SetEndingDate(CDate(SearchDates.Item(1).EndDate)) & " " & SetEndingTime & "'"
        Sql = Sql & " Order By TRY_CAST([Shift Date] AS datetime) asc"
        DrawLine
        PrintLine Trim$(EmPNames(FindEmp).WholeName), -1
        DrawLine
        IndHours = 0
    ConnectADORS Sql, Db, RS
     Do While Not RS.EOF
       If Trim(RS(5)) > "" And RS(5) <> "00:00" And Trim(RS(4)) > "" And RS(4) <> "00:00" Then
        breaktotal = DateDiff("n", CDate(RS(4)), CDate(RS(5)), WeekStartsOn, 1)
        If BreakMins > 0 And breaktotal <= BreakMins Then
          breaktotal = 0   '//  Option (4) = paid break
        Else
          breaktotal = breaktotal - BreakMins
        End If
       End If
       If Check_Null(RS![Time In]) > "" And Check_Null(RS![Time In]) <> "00:00" And Check_Null(RS![Time Out]) > "" And Check_Null(RS![Time Out]) <> "00:00" Then
        shiftTotal = DateDiff("n", CDate(RS![Time In]), CDate(RS![Time Out]), WeekStartsOn, 1)
       End If
       If shiftTotal > 0 Then
         TTime = (shiftTotal - breaktotal) / 60
       Else
         TTime = 0
       End If
       Hoursworked = CalculateTimeRecord(RS(2) & "", RS(7) & "", RS(3) & "", RS(4) & "", RS(5) & "", RS(6) & "", Total_Time)
       PayRate = Val(RS![Rate of Pay] & "")
       TTime = ConvertTimeToMinutes(Total_Time)
       OTPay = 0
       IndHours = IndHours + TTime
       EmpHours = EmpHours + TTime
       PrintLine "", 2000
       PrintLine Format(RS(2), "mm/dd"), 3000
       PrintLine Format(RS![Time In], "hh:mm ampm"), 4000
       PrintLine Format(RS(4), "hh:mm ampm"), 5000
       PrintLine Format(RS(5), "HH:MM ampm"), 6000
       PrintLine Format(RS![Time Out], "hh:mm ampm"), 7000
       PrintLine Format(TTime, "0.00"), 8000
       If Val(RS![Paid By] & "") = 0 Then
        PrintLine "0.00", 9000
       Else
        PrintLine Format((PayRate * TTime) + OTPay, "0.00"), 9000
       End If
       PrintLine Format(Val(RS![Sales] & ""), "0.00"), 10000
       If breaktotal > 0 Then
        PrintLine Format(Val(RS![Tip Amount] & "") + Val(RS![Credit Card Tips] & ""), "0.00"), 10500
        If BreakMins > 0 Then  '//  Option (4) = paid break
         PrintLine "Not Paid " & breaktotal, -1
        Else
         PrintLine "Break " & breaktotal, -1
        End If
       Else
        PrintLine Format(Val(RS![Tip Amount] & "") + Val(RS![Credit Card Tips] & ""), "0.00"), -1
       End If
       breaktotal = 0
       BRKMINS = BRKMINS + TTime
       If Val(RS![Paid By] & "") = 0 Then
      
       Else
        DollarSpent = DollarSpent + ((TTime * PayRate) + OTPay)
       End If
      RS.MoveNext
     Loop
     SetNothing RS
     End If
     
    Next FindEmp
    DbClose Db, RS
    DrawLine
    PrintLine "", 2000
    PrintLine "", 3000
    PrintLine "", 4000
    PrintLine "", 5000
    PrintLine "", 6000
    PrintLine "", 7000
    PrintLine Format(EmpHours, "0.00"), 8000
    PrintLine "", 9000
    PrintLine "", 10000
    PrintLine "", -1
    DrawLine
    EmpHours = 0
DrawLine
PrintLine "Total (Time by 100's) No overtime calculated", 7000: PrintLine Format(BRKMINS, "0.00"), 8000: PrintLine Format(DollarSpent, "0.00"), -1
If BreakMins > 0 Then '//  Option (4) = paid break
 PrintLine "All Breaks are recorded less paid minutes " & BreakMins, -1
End If
FinishJob 1
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Reports SS4 ")
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
Reportday.Visible = False
Reportday.Left = -3440
Sleep 0
End Sub
Private Sub Cal_Labor(Dept_Worked As String, Dept_PayRate As Currency, TimeWorked As Single, PaidHourly As Integer, DaysWorked As Integer, OverTime As Single, WeekNumber As Integer)
    On Error GoTo ERH
    Dim Cld        As Integer
    Dim SalaryPaid As Currency
    Dim DayRate    As Currency
    For Cld = 0 To Total_depts
       If Trim$(Dept_Worked) = Trim$(Dept_Pay(Cld, WeekNumber).DeptName) Then
        If PaidHourly = 1 Then
            Dept_Pay(Cld, WeekNumber).DeptDollars = Dept_Pay(Cld, WeekNumber).DeptDollars + ((Dept_PayRate * TimeWorked) + CalculateOT(Val(OverTime), Dept_PayRate))
            Dept_Pay(Cld, WeekNumber).TotatHours = (Val(Dept_Pay(Cld, WeekNumber).TotatHours) + (TimeWorked + CalculateOT(Val(OverTime), CCur(Dept_PayRate))))
        Else
            If Dept_PayRate > 0 And DaysWorked > 0 Then
                DayRate = Dept_PayRate / 7
                SalaryPaid = DayRate * DaysWorked
                Dept_Pay(Cld, WeekNumber).DeptDollars = Dept_Pay(Cld, WeekNumber).DeptDollars + SalaryPaid
            End If
        End If
        Exit For
       End If
    Next Cld
ERH:
End Sub
Private Sub SSCommand6_Click()
On Error GoTo ErrorHandler
If Can_Continue(SSCommand6.Caption) = False Then Exit Sub
Dim EDb          As ADODB.Connection
Dim Id           As ADODB.Recordset
Dim Db           As ADODB.Connection
Dim Tb           As ADODB.Recordset
Dim PRTb         As ADODB.Recordset
Dim JCS          As Integer
Dim BeginWeek    As String
Dim EndingDate   As String
Dim getweek      As Integer
Dim WrkTime      As Single
Dim EmpTwt()     As Single
Dim SearchName() As String
Dim NamesInOT    As Integer
Dim RN           As Integer
Dim PaidHourly   As Integer  '// 0 = Salary 1 = paid hourly
Dim Ck           As Integer
Dim DaysWorked   As Integer
Dim PRate        As Currency
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
  Exit Sub
Else
  BeginWeek = SearchDates.Item(1).StartDate
  EndingDate = SearchDates.Item(1).EndDate
End If
SSCommand6.Caption = "Working......"
Sleep 0
NamesInOT = 0
getweek = GetDateToNo(CDate(BeginWeek))
OpenDb "MSaleData", EDb
Sql = "Select * from [General Data] Where ISNULL([Terminated], 0) = 0  order by [last Name] asc"
ConnectADORS Sql, EDb, Id
 TotalRec = RecordsCount(Id) + 1
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
 Sql = "Select * from [Employee Records]  "
 Sql = Sql & "where TRY_CAST([Shift Date] AS date) between '" & CDate(BeginWeek) & "' and '" & CDate(EndingDate) & "'"
ConnectADORS Sql, Db, Tb
ReDim SearchName(TotalRec)
ReDim EmpTwt(TotalRec)
RN = -1
Do While Not Id.EOF
    RN = RN + 1
    SearchName(RN) = Id.fields("First Name").Value & " " & Id.fields("Last Name").Value
    Sql = "Select * from [Job Codes] where [First Name] = '" & Id.fields("First Name").Value & "' and  [Last Name] = '" & Id.fields("Last Name").Value & "'"
    ConnectADORS Sql, EDb, PRTb
    If Not PRTb.EOF Then
     ReDim PayRate(RecordsCount(PRTb))
    End If
    JCS = 0
    Do While Not PRTb.EOF
         PayRate(JCS).JobCodes = PRTb![Job Code] & ""
         PayRate(JCS).JobRate = Val(PRTb![rate] & "")
         PayRate(JCS).JobHourly = Val(PRTb![Hourly] & "")
         PayRate(JCS).UseDayOT = Val(PRTb![Day OT] & "")
         JCS = JCS + 1
      PRTb.MoveNext
    Loop
  SetNothing PRTb
  If Not Tb.BOF Then Tb.MoveFirst
  PaidHourly = 0
  Do While Not Tb.EOF
   If Trim$(SearchName(RN)) = Trim$(Tb.fields("Employee Name").Value) Then
      Call CalculateTimeRecord(Tb.fields("Shift Date").Value & "", Tb.fields("Punch Out Date").Value & "", Tb.fields("Time In").Value & "", Tb![Break Out] & "", Tb![Break In] & "", Tb.fields("Time Out").Value & "", Total_Time)
      For Ck = 0 To JCS - 1
        If Trim$(Tb.fields("Department Name").Value) = Trim$(PayRate(0).JobCodes) Then
           PRate = Val(Tb![Rate of Pay] & "")
           PaidHourly = Abs(PayRate(Ck).JobHourly)
           Exit For
        End If
      Next Ck
     WrkTime = ConvertTimeToMinutes(Total_Time)
     DaysWorked = GetDateToNo(CDate(Date))
     Cal_Labor Tb.fields("Department Name").Value, PRate, WrkTime, PaidHourly, DaysWorked, OvertimeHrs(EmpTwt(RN), WrkTime), getweek
     EmpTwt(RN) = EmpTwt(RN) + WrkTime
     'If EmpTwt(RN) >= Input_Option(5) Then NamesInOT = 1
   End If
   Tb.MoveNext
 Loop
Id.MoveNext
Loop
DbClose Db, Tb
DbClose Db, Id
If NamesInOT = 0 Then
   See_Msg "Hours are OK", 5, ""
   SSCommand6.Caption = "Over Time Watch"
End If
If NamesInOT = 0 Then Exit Sub
OneMoreTime:
SetPage ReportView.View, "Over Time Watch"
Dim Data As String
Data = ""
For RN = 0 To TotalRec
    'If EmpTwt(RN) >= Val(Input_Option(5)) Then   '//Input_Option(5) overtime hours
    '  Data = Data & SearchName(RN) & "|" & Format(EmpTwt(RN), "0.00") & ";"
    'End If
Next RN
SetReport Data, SetReportFormat("+<2000|>+1000;"), SetReportHeader("Name|Hours Worked;"), tbBox
FinishJob 1
SSCommand6.Caption = "Over Time Watch"
DoEvents
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Reports SS6 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Public Sub AverageRateofPay()
Dim EDb          As ADODB.Connection
Dim Id           As ADODB.Recordset
Dim Db           As ADODB.Connection
Dim Tb           As ADODB.Recordset
Dim BeginWeek    As String
Dim EndingDate   As String
Dim Hoursworked  As Single
Dim Data         As String
Dim Income       As Currency
Dim HourAvg      As Currency
Dim ShiftNo      As String
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
If SearchDates Is Nothing Then
  Exit Sub
Else
  BeginWeek = SearchDates.Item(1).StartDate
  EndingDate = SearchDates.Item(1).EndDate
End If
See_Msg "Reports Tipped Employees ", 0, "Average hourly rate under " & Format(Val(Input_Option(61)), "0.00" & " per hour")
OpenDb SvrPath.LocalTerminalPath & TimeRecordsDb, EDb
Sql = "Select Distinct [Employee Name] from [Employee Records] order by [Employee Name] asc"
Data = ""
ConnectADORS Sql, EDb, Id
Do While Not Id.EOF
    Hoursworked = 0: Income = 0
    Sql = "Select * from [Employee Records] Where [Employee Name] = '" & Id.fields("Employee Name").Value & "'"
    Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) between '" & BeginWeek & "' and '" & EndingDate & "'"
    ConnectADORS Sql, EDb, Tb
    ShiftNo = ""
    Do While Not Tb.EOF
     
     If Val(Tb.fields("Tip Amount").Value & "") > 0 Or Val(Tb.fields("Credit Card Tips").Value & "") > 0 Then
       Income = Income + Val(Tb.fields("Tip Amount").Value & "")
       If ShiftNo <> Tb.fields("Shift").Value & "" Then
        Income = Income + Val(Tb.fields("Credit Card Tips").Value & "")
        Income = Income + Val(Tb.fields("Tip Share").Value & "")
        Income = Income - Val(Tb.fields("Tips Paid Out").Value & "")
       End If
       ShiftNo = Tb.fields("Shift").Value & ""
       Call CalculateTimeRecord(Tb(2) & "", Tb(7) & "", Tb(3) & "", Tb![Break Out] & "", Tb![Break In] & "", Tb(6) & "", Total_Time)
       If Val(Tb.fields("Rate Of Pay").Value & "") > 0 Then
         Income = Income + Val(Tb.fields("Rate Of Pay").Value & "") * ConvertTimeToMinutes(Total_Time)
       Else
         Income = Income + Val(Input_Option(62)) * ConvertTimeToMinutes(Total_Time)
       End If
       Hoursworked = Hoursworked + Total_Time
      End If
      Tb.MoveNext
    Loop
    If Hoursworked > 0 And Income > 0 Then
     HourAvg = Income / ConvertTimeToMinutes(Hoursworked)
     If HourAvg < Val(Input_Option(61)) Then
      Data = Data & Id.fields("Employee Name").Value & "|" & Format(ConvertTimeToMinutes(Hoursworked), "0.00") & "|" & Format(Income, "0.00") & "|" & Format(HourAvg, "0.00") & ";"
     End If
    End If
    SetNothing Tb
Id.MoveNext
Loop
DbClose Db, Id
If Data = "" Then
   See_Msg "No Exceptions..", 5, ""
   Exit Sub
End If
SetPage ReportView.View, "Average Rate of Pay" & vbCrLf & " Period " & BeginWeek & " To " & EndingDate

SetReport Data, SetReportFormat("+<2500|>+1500|>+1500|>+1500;"), SetReportHeader("Name|Hours Worked|Income|Average;"), tbBox
FinishJob 1
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Reports SS6 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub SSCommand7_Click()

If Can_Continue(SSCommand7.Caption) = False Then Exit Sub
'SSCommand7.Caption = "Working....."
Sleep 0
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
Dim PrinterPort As String
Dim Heading As String
Heading = "Who's Punched In"
TWRKTIME = 0: found1 = 0
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
 Sql = "Select Distinct [employee Name],[Time In] from [Employee records] where [time in] > '" & "" & "' and isdate([time out]) = 0 order by [employee Name] asc"
ConnectADORS Sql, Db, Tb
 If Tb.EOF Then
   Call See_Msg("No employees are on the clock.", 5, "")
   DbClose Db, Tb
   Exit Sub
 End If
 SetPage ReportView.View, Heading
 PrintLine "On the Clock...", 6500: PrintLine "Time On", -1
 DrawLine
 Do While Not Tb.EOF
  If Trim$(Tb![employee name]) > "" Then
   found1 = 1
   PrintLine Tb![employee name], 6500: PrintLine Format(Tb![Time In], "hh:mm ampm"), -1
  End If
 Tb.MoveNext
Loop
DbClose Db, Tb
'SSCommand7.Caption = "Who's on the Clock"
DrawLine
FinishJob 1
If Check_Option(96) = 1 Then Exit Sub
If See_Msg("Print Report", 4, "") <> 6 Then Exit Sub
Printer_Drivers "Receipt"
PF = FreeFile
PrinterPort = "Who on the clock"
FP = 0
OneMoreTime:
Open PrinterPort For Output Shared As #PF
CenterText = EC(8)
LeftText = EC(9)
'RightText = Chr(27) + Chr(97) + Chr(2)
Print #PF, EC(0) 'Int printer
Print #PF, EC(1); EC(3); CenterText; StoreName; EC(2); EC(7)
Print #PF, LeftText; EC(7)
Print #PF, EC(1); EC(3); CenterText; Heading; EC(2); EC(7)
Print #PF, ; EC(9); EC(7)
Print #PF, CenterText; EC(1); EC(3); ForWhen; LeftText; EC(2); EC(7)
Print #PF, LeftText; EC(7)
Print #PF, Time;
Print #PF, Tab(20); "Date  "; Trim(Date); EC(7)
Print #PF, String(33, 45); EC(7)
TWRKTIME = 0: found1 = 0
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
 Sql = "Select Distinct [employee Name],[Time In] from [Employee records] where [time in] > '" & "" & "' and [time out] = '" & "" & "' order by [employee Name] asc"
ConnectADORS Sql, Db, Tb
 If Tb.RecordCount > 0 Then
   Print #PF, "On The Clock            Time On"; EC(7)
   Print #PF, String(33, 45); EC(7)
   Do While Not Tb.EOF
   If Trim$(Tb![employee name]) > "" Then
   found1 = 1
   Print #PF, Tb![employee name]; Tab(24); Format(Tb![Time In], "HH:MM ampm"); EC(7)
   End If
   Tb.MoveNext
   Loop
 End If
DbClose Db, Tb
If found1 = 0 Then Print #PF, "All are punched out"; EC(7)
Print #PF, String(33, 45); EC(7)
Print #PF, EC(6) & EC(5); EC(7)
Close #PF
SendJobToPrinter PrinterPort, "Receipt", False
SSCommand7.Caption = "Who's on the Clock"
Sleep 0
End Sub
Private Sub SSCommand8_Click()
Dim BeginWeek As String
Dim EndingDate As String
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Sql = "Select distinct [Employee Name] from [Employee Records] "
Sql = Sql & " order by [Employee Name] asc"
ConnectADODB SvrPath.LocalTerminalPath & TimeRecordsDb, Db
ConnectADORS Sql, Db, RS
List2.Clear
Do While Not RS.EOF
  'List2.AddItem RS![First Name] & "" & " " & RS![Last Name] & ""
  List2.AddItem RS![employee name] & ""
  RS.MoveNext
Loop
DbClose Db, RS
CenterObject Time_Reports, EList
EList.Visible = True
EList.ZOrder 0
End Sub
Private Sub SSCommand9_Click()
If Can_Continue(SSCommand9.Caption) = False Then Exit Sub
Dim BeginWeek As String
Dim EndingDate As String
LoadDates LoadPRDates, TimeRecordsDb, "Shift Date"
 If SearchDates Is Nothing Then
  Exit Sub
 Else
  BeginWeek = SearchDates.Item(1).StartDate
  EndingDate = SearchDates.Item(1).EndDate
End If
SyncPayrollData BeginWeek, EndingDate
PrintPayRoll False, BeginWeek, EndingDate, False, "PrePrint Closed Time Records"
Trans_Log "Payroll Print out for " & BeginWeek & " " & EndingDate
End Sub
Public Sub Load_Depts()
On Error Resume Next
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim lw As Integer
Dim DN As Integer
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
Sql = "Select * from [Job Titles] order by [Job Title] asc"
ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then Tb.MoveLast
  Total_depts = Tb.RecordCount
  If Not Tb.BOF Then Tb.MoveFirst
  ReDim Dept_Pay(Total_depts, 53)
   DN = -1
   Do While Not Tb.EOF
      DN = DN + 1
      For lw = 0 To 52
      Dept_Pay(DN, lw).DeptName = Tb(0) & ""
      Dept_Pay(DN, lw).DeptDollars = 0
      Dept_Pay(DN, lw).TipPostion = Val(Tb(1) & "")    'Tip Position
      Dept_Pay(DN, lw).TipsOut = Val(Tb(2) & "") 'Tip Pool Perc
      Dept_Pay(DN, lw).TotatHours = 0
      Next lw
      Tb.MoveNext
   Loop
DbClose Db, Tb
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        