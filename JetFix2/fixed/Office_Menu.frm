VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Object = "{84E5CF37-E467-4AC2-89C4-C6002FFB5055}#25.1#0"; "CHARTV~1.OCX"
Begin VB.Form Office_Menu 
   Appearance      =   0  'Flat
   BackColor       =   &H001F1B18&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POS Office Menu"
   ClientHeight    =   10380
   ClientLeft      =   990
   ClientTop       =   1020
   ClientWidth     =   15210
   FillColor       =   &H001F1B18&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H001F1B18&
   Icon            =   "Office_Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh VoidWatch 
      Height          =   495
      Left            =   2160
      TabIndex        =   77
      Top             =   3720
      Width           =   1455
      _Version        =   589828
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Shape           =   1
      CornerFactor    =   10
      Surface         =   2
      BackColorContainer=   16119544
      ButtonRaiseFactor=   0
      LogPixels       =   96
      SpecialEffectFactor=   2
      ForeColorDisabled=   16119544
      UserData        =   0.1
      textCaption     =   "Office_Menu.frx":1CCA
      textLT          =   "Office_Menu.frx":1CE2
      textCT          =   "Office_Menu.frx":1CFA
      textRT          =   "Office_Menu.frx":1D12
      textLM          =   "Office_Menu.frx":1D2A
      textRM          =   "Office_Menu.frx":1D42
      textLB          =   "Office_Menu.frx":1D5A
      textCB          =   "Office_Menu.frx":1D72
      textRB          =   "Office_Menu.frx":1D8A
      colorBack       =   "Office_Menu.frx":1DA2
      colorIntern     =   "Office_Menu.frx":1DCC
      colorMO         =   "Office_Menu.frx":1DF6
      colorFocus      =   "Office_Menu.frx":1E20
      colorDisabled   =   "Office_Menu.frx":1E4A
      colorPressed    =   "Office_Menu.frx":1E74
      Style           =   8
      ProgressBarColor=   16777215
   End
   Begin VB.Frame MenuList 
      Appearance      =   0  'Flat
      BackColor       =   &H001F1B18&
      Caption         =   "Menu List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   11160
      TabIndex        =   51
      Top             =   4560
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox MenuName 
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
         Height          =   495
         Left            =   3600
         TabIndex        =   60
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox CountDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   58
         Text            =   "0"
         Top             =   3140
         Width           =   1250
      End
      Begin VB.TextBox Inventory 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   57
         Text            =   "0"
         Top             =   2050
         Width           =   1250
      End
      Begin VB.ListBox ListMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   4080
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   3600
         MaskColor       =   &H00808080&
         TabIndex        =   53
         Top             =   3840
         Width           =   2685
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4680
         TabIndex        =   52
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Item"
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
         Height          =   375
         Left            =   3600
         TabIndex        =   59
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number Available"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   56
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Inventory"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         TabIndex        =   55
         Top             =   1680
         Width           =   2655
      End
   End
   Begin VB.PictureBox Progression 
      Appearance      =   0  'Flat
      BackColor       =   &H001B1816&
      FillColor       =   &H001B1816&
      ForeColor       =   &H00404040&
      Height          =   5850
      Left            =   3840
      ScaleHeight     =   5820
      ScaleWidth      =   9165
      TabIndex        =   61
      Top             =   2280
      Visible         =   0   'False
      Width           =   9195
      Begin BTNENHLib4.BtnEnh Progress 
         Height          =   855
         Left            =   1320
         TabIndex        =   62
         Top             =   3000
         Width           =   6495
         _Version        =   589828
         _ExtentX        =   11456
         _ExtentY        =   1508
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
         Shape           =   1
         CornerFactor    =   10
         BackColorContainer=   1775638
         ButtonRaiseFactor=   0
         LogPixels       =   96
         SpecialEffectFactor=   2
         UserData        =   0.1
         textCaption     =   "Office_Menu.frx":1E9E
         textLT          =   "Office_Menu.frx":1EB6
         textCT          =   "Office_Menu.frx":1ECE
         textRT          =   "Office_Menu.frx":1EE6
         textLM          =   "Office_Menu.frx":1EFE
         textRM          =   "Office_Menu.frx":1F16
         textLB          =   "Office_Menu.frx":1F2E
         textCB          =   "Office_Menu.frx":1F46
         textRB          =   "Office_Menu.frx":1F5E
         colorBack       =   "Office_Menu.frx":1F76
         colorIntern     =   "Office_Menu.frx":1FA0
         colorMO         =   "Office_Menu.frx":1FCA
         colorFocus      =   "Office_Menu.frx":1FF4
         colorDisabled   =   "Office_Menu.frx":201E
         colorPressed    =   "Office_Menu.frx":2048
         Style           =   8
         LightDirection  =   8
         SpotlightColor  =   2104855
         SpotlightType   =   1
         ProgressBarValue=   1
         ProgressBarColor=   8421504
      End
      Begin VB.Label UPdateWhere 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Importing Progress"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   63
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H001B1816&
         BorderColor     =   &H00FF6600&
         BorderWidth     =   3
         FillColor       =   &H001B1816&
         FillStyle       =   0  'Solid
         Height          =   5535
         Left            =   120
         Top             =   120
         Width           =   8895
      End
   End
   Begin VB.PictureBox BKDROP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3135
      Left            =   4080
      ScaleHeight     =   3105
      ScaleWidth      =   3705
      TabIndex        =   36
      Top             =   240
      Width           =   3735
      Begin VB.TextBox SafeAdjust 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox SafeCash 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox BKDP 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   3
         Left            =   2040
         TabIndex        =   46
         Text            =   "0.00"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox BKDP 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   2
         Left            =   2040
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox BKDP 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   1
         Left            =   2040
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox BKDP 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Save 
         BackColor       =   &H0080FF80&
         Caption         =   "Save Bank Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Safe Adjustment"
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
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash In Safe"
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
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Bank Deposits"
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
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Deposit 4"
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
         TabIndex        =   41
         Top             =   1440
         Width           =   1700
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Deposit 3"
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
         TabIndex        =   40
         Top             =   1080
         Width           =   1700
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Deposit 2"
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
         TabIndex        =   39
         Top             =   720
         Width           =   1700
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Deposit 1"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1700
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   1530
         Left            =   120
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   2520
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   720
      Width           =   1095
   End
   Begin VSFlex8DAOCtl.VSFlexGrid LaborData 
      Height          =   2895
      Left            =   720
      TabIndex        =   27
      Top             =   7200
      Width           =   6495
      _cx             =   1987456192
      _cy             =   1987449842
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorBkg    =   2038552
      BackColorAlternate=   -2147483643
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
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Index           =   5
      Left            =   2520
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Index           =   4
      Left            =   2520
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Index           =   3
      Left            =   2520
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7320
      Top             =   9360
   End
   Begin VB.TextBox ABT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   18
      Top             =   6435
      Width           =   1695
   End
   Begin VB.TextBox ABT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   17
      Top             =   6435
      Width           =   1695
   End
   Begin CDChartViewer.ChartViewer LD 
      Height          =   5055
      Left            =   8160
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8916
      Appearance      =   1
   End
   Begin VB.TextBox APR 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Top             =   5955
      Width           =   1695
   End
   Begin VB.TextBox APR 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   13
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox AW 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   10
      Top             =   5475
      Width           =   1695
   End
   Begin VB.TextBox AW 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000080&
      Height          =   330
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00404000&
      Height          =   330
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox OfficeData 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   1080
      Width           =   1095
   End
   Begin CDChartViewer.ChartViewer LaborBreakdown 
      Height          =   4695
      Left            =   8160
      Top             =   5520
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      Appearance      =   1
   End
   Begin VB.Frame SiteEditor 
      BackColor       =   &H00404040&
      Caption         =   "Add/Edit Sites"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3135
      Left            =   8640
      TabIndex        =   64
      Top             =   2280
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set Active"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         TabIndex        =   76
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox PW 
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
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   75
         Top             =   1800
         Width           =   3495
      End
      Begin VB.ComboBox User 
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
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   74
         Top             =   1320
         Width           =   3495
      End
      Begin VB.ComboBox IpAddress 
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
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   73
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2040
         TabIndex        =   71
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4200
         TabIndex        =   70
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3120
         TabIndex        =   69
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox SiteName 
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
         Left            =   1680
         TabIndex        =   65
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ip Address"
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
         Left            =   360
         TabIndex        =   72
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   68
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   360
         TabIndex        =   67
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Site Name"
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
         Left            =   360
         TabIndex        =   66
         Top             =   360
         Width           =   1215
      End
   End
   Begin BTNENHLib4.BtnEnh DiscWatch 
      Height          =   495
      Left            =   2160
      TabIndex        =   78
      Top             =   4320
      Width           =   1455
      _Version        =   589828
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Surface         =   2
      LogPixels       =   96
      UserData        =   0.1
      textCaption     =   "Office_Menu.frx":2072
      textLT          =   "Office_Menu.frx":208A
      textCT          =   "Office_Menu.frx":20A2
      textRT          =   "Office_Menu.frx":20BA
      textLM          =   "Office_Menu.frx":20D2
      textRM          =   "Office_Menu.frx":20EA
      textLB          =   "Office_Menu.frx":2102
      textCB          =   "Office_Menu.frx":211A
      textRB          =   "Office_Menu.frx":2132
      colorBack       =   "Office_Menu.frx":214A
      colorIntern     =   "Office_Menu.frx":2174
      colorMO         =   "Office_Menu.frx":219E
      colorFocus      =   "Office_Menu.frx":21C8
      colorDisabled   =   "Office_Menu.frx":21F2
      colorPressed    =   "Office_Menu.frx":221C
      Style           =   8
      ProgressBarColor=   12632256
   End
   Begin BTNENHLib4.BtnEnh ServeTimeWatch 
      Height          =   495
      Left            =   5520
      TabIndex        =   79
      Top             =   3720
      Width           =   1815
      _Version        =   589828
      _ExtentX        =   3201
      _ExtentY        =   873
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
      Surface         =   2
      LogPixels       =   96
      UserData        =   0.1
      textCaption     =   "Office_Menu.frx":2246
      textLT          =   "Office_Menu.frx":225E
      textCT          =   "Office_Menu.frx":2276
      textRT          =   "Office_Menu.frx":228E
      textLM          =   "Office_Menu.frx":22A6
      textRM          =   "Office_Menu.frx":22BE
      textLB          =   "Office_Menu.frx":22D6
      textCB          =   "Office_Menu.frx":22EE
      textRB          =   "Office_Menu.frx":2306
      colorBack       =   "Office_Menu.frx":231E
      colorIntern     =   "Office_Menu.frx":2348
      colorMO         =   "Office_Menu.frx":2372
      colorFocus      =   "Office_Menu.frx":239C
      colorDisabled   =   "Office_Menu.frx":23C6
      colorPressed    =   "Office_Menu.frx":23F0
      Style           =   8
      ProgressBarColor=   12632256
   End
   Begin BTNENHLib4.BtnEnh LaborWatch 
      Height          =   495
      Left            =   5520
      TabIndex        =   80
      Top             =   4320
      Width           =   1815
      _Version        =   589828
      _ExtentX        =   3201
      _ExtentY        =   873
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
      Surface         =   2
      LogPixels       =   96
      UserData        =   0.1
      textCaption     =   "Office_Menu.frx":241A
      textLT          =   "Office_Menu.frx":2432
      textCT          =   "Office_Menu.frx":244A
      textRT          =   "Office_Menu.frx":2462
      textLM          =   "Office_Menu.frx":247A
      textRM          =   "Office_Menu.frx":2492
      textLB          =   "Office_Menu.frx":24AA
      textCB          =   "Office_Menu.frx":24C2
      textRB          =   "Office_Menu.frx":24DA
      colorBack       =   "Office_Menu.frx":24F2
      colorIntern     =   "Office_Menu.frx":251C
      colorMO         =   "Office_Menu.frx":2546
      colorFocus      =   "Office_Menu.frx":2570
      colorDisabled   =   "Office_Menu.frx":259A
      colorPressed    =   "Office_Menu.frx":25C4
      Style           =   8
      ProgressBarColor=   12632256
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serving Time Watch "
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
      Height          =   495
      Index           =   21
      Left            =   4000
      TabIndex        =   35
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Watch"
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
      Index           =   20
      Left            =   3960
      TabIndex        =   34
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Watch"
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
      Index           =   19
      Left            =   360
      TabIndex        =   33
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Void Watch"
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
      Index           =   18
      Left            =   360
      TabIndex        =   32
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Business Day Is"
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
      Index           =   17
      Left            =   360
      TabIndex        =   30
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Information"
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
      Height          =   315
      Index           =   16
      Left            =   275
      TabIndex        =   29
      Top             =   260
      Width           =   3680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close Out Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   28
      Top             =   5160
      Width           =   7575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H001F1B18&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   240
      Top             =   7080
      Width           =   7575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Average"
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
      Index           =   14
      Left            =   480
      TabIndex        =   26
      Top             =   2520
      Width           =   2000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Count"
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
      Index           =   13
      Left            =   480
      TabIndex        =   23
      Top             =   2160
      Width           =   2000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "WTD Net Sales"
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
      Index           =   12
      Left            =   480
      TabIndex        =   22
      Top             =   2880
      Width           =   2000
   End
   Begin VB.Label BPL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
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
      Index           =   1
      Left            =   2040
      TabIndex        =   20
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label BPL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "End"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   19
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "End"
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
      Index           =   9
      Left            =   4920
      TabIndex        =   16
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
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
      Index           =   8
      Left            =   2040
      TabIndex        =   15
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "End"
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
      Index           =   7
      Left            =   4920
      TabIndex        =   12
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
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
      Index           =   6
      Left            =   2400
      TabIndex        =   11
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label BPL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Batch Period"
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
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Payroll Period"
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
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Week Period"
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
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Voided Sales"
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
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   2000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discounts"
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
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Net Sales"
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
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2000
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   240
      Top             =   240
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   240
      Top             =   5160
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0027201C&
      BorderColor     =   &H004E4844&
      FillColor       =   &H0027201C&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   240
      Top             =   3480
      Width           =   7575
   End
   Begin VB.Menu Quit 
      Caption         =   "Exit"
   End
   Begin VB.Menu TRE 
      Caption         =   "Maintenance"
      Begin VB.Menu MMain 
         Caption         =   "Menu Maintenance"
      End
      Begin VB.Menu EMain 
         Caption         =   "Employee Maintenance"
      End
      Begin VB.Menu TRMain 
         Caption         =   "Time Records"
      End
      Begin VB.Menu RSVATION 
         Caption         =   "Reservations"
      End
      Begin VB.Menu TBLLAY 
         Caption         =   "Dining Room Editor"
      End
      Begin VB.Menu RSETWL 
         Caption         =   "Reset Waiting List"
      End
      Begin VB.Menu ImportMenuFile 
         Caption         =   "Import Menu"
      End
      Begin VB.Menu ADDEDIT 
         Caption         =   "Add/Edit Remote Sites"
      End
   End
   Begin VB.Menu MRGMenu 
      Caption         =   "Manager Functions"
      Begin VB.Menu MGRMENU 
         Caption         =   "Manager Menu"
      End
      Begin VB.Menu Set86 
         Caption         =   "Set Inventory/86 Item"
      End
      Begin VB.Menu UPHS 
         Caption         =   "Update Hot Schedules"
      End
   End
   Begin VB.Menu DayClose 
      Caption         =   "Daily Close"
   End
   Begin VB.Menu WeekClose 
      Caption         =   "Weekly Close"
   End
   Begin VB.Menu FReports 
      Caption         =   "Financial Reports"
      Begin VB.Menu ShowFinancial 
         Caption         =   "Sales Report"
      End
      Begin VB.Menu CashReport 
         Caption         =   "Cash Report"
         Visible         =   0   'False
      End
      Begin VB.Menu DEPTTEND 
         Caption         =   "Department/Tender Report"
      End
   End
   Begin VB.Menu PSheet 
      Caption         =   "Prep Sheet"
   End
   Begin VB.Menu RPM 
      Caption         =   "Food Cost"
   End
   Begin VB.Menu Sch 
      Caption         =   "Employee Schedular"
   End
   Begin VB.Menu SVRRPTS 
      Caption         =   "Server Reports"
   End
   Begin VB.Menu SafeAdj 
      Caption         =   "Safe Adjustment"
      Begin VB.Menu ATS 
         Caption         =   "Add to safe"
      End
      Begin VB.Menu dfs 
         Caption         =   "Deduct from safe"
      End
   End
End
Attribute VB_Name = "Office_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TerminalFailToRead As Boolean
Private PollingTime As String
Private Type RestLabor
        JobTitle As String
        JobRate  As Currency
        JobHours As Single
End Type
Private Jobs()                  As RestLabor
Private Total_depts             As Integer
Private Timing                  As Long
Private Cd As ChartDirector.API
'Dim WithEvents HSFTP As SFTPLib.SFTP
Private Sub ADDEDIT_Click()
Dim TXTFileInfo() As Variant
Dim Wl            As Integer
SiteName.Clear: IpAddress.Clear: User.Clear: PW.Clear
If ParseData(ReturnAllFileData(SvrPath.TerminalPath & "DAT Files\" & "RemoteSites.dat"), TXTFileInfo()) Then
    For Wl = 0 To UBound(TXTFileInfo, 1) Step 4
       SiteName.AddItem TXTFileInfo(Wl)
       IpAddress.AddItem TXTFileInfo(Wl + 1), SiteName.NewIndex
       User.AddItem TXTFileInfo(Wl + 2), SiteName.NewIndex
       PW.AddItem TXTFileInfo(Wl + 3), SiteName.NewIndex
    Next Wl
End If
End Sub
Private Sub Command1_Click()
MenuList.Visible = False
MenuList.Left = -5000
End Sub
Private Sub ATS_Click()
Dim AddSafeCash As Currency
AddSafeCash = Val(ShowNumberPad("Enter Cash Amount", False, 0, 0))
If AddSafeCash = 0 Then Exit Sub
End Sub
Private Sub Command2_Click()
If MenuName.Text = "" Then Exit Sub
Dim Menu     As ADODB.Connection
Dim MRecords As ADODB.Recordset
ConnectADODB SvrPath.TerminalPath & GetCountDownDB, Menu
  Sql = "Update from [Menu Items] set [Quick Button] = " & Nz(Inventory.Text) & " and [Inventory Amount] = " & Nz(CountDown.Text) & " Where [menu item name] = '" & MenuName.Text & "'"
 Menu.Execute Sql
DbClose Menu, Nothing
Inventory.Text = 0
CountDown.Text = 9999
MenuName.Text = ""
See_Msg "Saved....", 5, "Inventory Set"
End Sub
Private Sub Command3_Click()
   If SiteName.Text = "" Then
     See_Msg "Nothing To Set To Active", 5, ""
     Exit Sub
   End If
End Sub
Private Sub Command4_Click()
If SiteName.Text = "" Then Exit Sub
Dim Wl            As Integer
Dim Dstr          As String
Wl = Finditem(SiteName, SiteName.Text)
If Wl = -1 Then
 SiteName.AddItem SiteName.Text
 IpAddress.AddItem IpAddress.Text, SiteName.NewIndex
 User.AddItem User.Text, SiteName.NewIndex
 PW.AddItem PW.Text, SiteName.NewIndex
Else
 SiteName.AddItem SiteName.Text
 IpAddress.AddItem IpAddress.Text, SiteName.NewIndex
 User.AddItem User.Text, SiteName.NewIndex
 PW.AddItem PW.Text, SiteName.NewIndex
 SiteName.RemoveItem Wl
 IpAddress.RemoveItem Wl
 User.RemoveItem Wl
 PW.RemoveItem Wl
End If
DeleteFile SvrPath.TerminalPath & "DAT Files\" & "RemoteSites.dat"
For Wl = 0 To SiteName.ListCount - 1
Dstr = SiteName.List(Wl) & "," & IpAddress.List(Wl) & "," & User.List(Wl) & "," & PW.List(wi)
WriteTxtFile SvrPath.TerminalPath & "DAT Files\" & "RemoteSites.dat", Dstr, AppendFile
Next Wl
SiteName.Text = "": IpAddress.Text = "": User.Text = "": PW.Text = ""
See_Msg "Saved.....", 5, ""
End Sub
Private Sub Command5_Click()
SiteEditor.Visible = False
SiteEditor.Left = -5000
End Sub
Private Sub Command6_Click()
If SiteName.Text = "" Then Exit Sub
Dim Wl            As Integer
Dim Dstr          As String
Wl = Finditem(SiteName, SiteName.Text)
If Wl = -1 Then
  Exit Sub
Else
 SiteName.RemoveItem Wl
 IpAddress.RemoveItem Wl
 User.RemoveItem Wl
 PW.RemoveItem Wl
End If
DeleteFile SvrPath.TerminalPath & "DAT Files\" & "RemoteSites.dat"
For Wl = 0 To SiteName.ListCount - 1
Dstr = SiteName.List(Wl) & "," & IpAddress.List(Wl) & "," & User.List(Wl) & "," & PW.List(wi)
WriteTxtFile SvrPath.TerminalPath & "DAT Files\" & "RemoteSites.dat", Dstr, AppendFile
Next Wl
SiteName.Text = "": IpAddress.Text = "": User.Text = "": PW.Text = ""
See_Msg "Done.....", 5, ""
End Sub

Private Sub DayClose_Click()
If Check_Option(199) = 1 Then
   See_Msg "Cannot Close From This Station", 5, "Close Out Disabled."
   Exit Sub
End If
If Can_Continue(DayClose.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim SRC As Integer
Dim Db  As ADODB.Connection
Dim RS  As ADODB.Recordset
Dim AuditsOpen As Boolean
SETvariables
IsCommonPresent
'FindAllSalesData
FindNoPayments
'If Check_Option(65) = 0 Then
   If UCase(GetConfigInfo("Daily Close", "Close Day With Open Phone Orders", "NO", "FUNCTIONS.INI")) = "ENABLED" Then
      
   Else
      ConnectADODB CallInOrderDb, Db
        Sql = "Select * from [Open Checks] where [Presented] = 1"
      ConnectADORS Sql, Db, RS
        If Not RS.EOF Then
          See_Msg "Please Close All Presented Phone Orders", 5, "Open Phone Orders"
          DbClose Db, RS
          Exit Sub
        End If
        DbClose Db, RS
   End If
   ConnectADODB CheckStatDb, Db
   ConnectADORS "Select * from [Open Checks] where [Close Out Day] = 0", Db, RS
       If Not RS.EOF Then
          See_Msg "There is " & CStr(RecordsCount(RS)) & " Check(s) Open", 5, "Can Not Close Out"
          DbClose Db, RS
          Exit Sub
       End If
       DbClose Db, RS
      AuditsOpen = False
      ConnectADODB FinancialDb, Db
      ConnectADORS "Select * from [Audits]", Db, RS
        If Not RS.EOF Then
          See_Msg "Cashier Audit Open On " & MyRegisters(SRC).regName, 5, "Cannot Close System"
          AuditsOpen = True
        End If
        DbClose Db, RS
        ConnectADODB MyRegisters(SRC).Svr & "MSaleData", Db
          Sql = "Select * from [Check Numbers] Where [Audit Status] = 0"
        ConnectADORS Sql, Db, RS
         If Not RS.EOF Then
          If See_Msg("Server Audits Open On " & MyRegisters(SRC).SvrName, 4, "Close All Server Audits ? ") <> 6 Then
           AuditsOpen = True
          End If
         End If
         DbClose Db, RS
         If AuditsOpen Then Exit Sub
         CloseAllServerAudits
If ALLClockedOut = False Then Exit Sub
Timer1.Enabled = False
If UCase(GetConfigInfo("Daily Close", "Restrict Daily Close", "NO", "FUNCTIONS.INI")) = "ENABLED" Then
  Dim StartTime As String
  Dim EndTime As String
  StartTime = GetConfigInfo("Daily Close", "Start Time", "08:00:00", "FUNCTIONS.INI")
  EndTime = GetConfigInfo("Daily Close", "End Time", "22:00:00", "FUNCTIONS.INI")
  If IsBetween(CStr(Now), CStr(Now), Format(Date, "MM/DD/YYYY ") + StartTime, Format(Date, "MM/DD/YYYY ") + EndTime) Then
     See_Msg "Daily Close Restricted", 5, "Between " & Format(StartTime, "HH:MM AMPM") & " And " & Format(EndTime, "HH:MM AMPM")
  End If
End If
Closing_Report.Show
Closing_Report.ZOrder 0
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Office Menu SS1")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub GetActiveWeek()
On Error GoTo ErrorHandler
Dim SRC As Integer
Dim Db As ADODB.Connection
Dim OP As ADODB.Recordset
Dim Tb As ADODB.Recordset
          ConnectADODB CheckStatDb, Db
             Sql = "Select * from [Check Detail] Order By TRY_CAST([Check Date] AS datetime) asc"
            ConnectADORS Sql, Db, Tb
            If Not Tb.EOF Then
              If Not Tb.BOF Then Tb.MoveFirst
                If Not IsDate(AW(0).Text) Then
                   AW(0).Text = Format(Tb![Check Date], "MM/DD/YYYY")
                Else
                   If CDate(Tb![Check Date]) < CDate(AW(0).Text) Then
                     AW(0).Text = Format(Tb![Check Date], "MM/DD/YYYY")
                   End If
                End If
                SetNothing Tb
                 Sql = "Select * from [Check Detail] Order By TRY_CAST([Check Date] AS datetime) desc"
                ConnectADORS Sql, Db, Tb
                If Not IsDate(AW(1).Text) Then
                   AW(1).Text = Format(Tb![Check Date], "MM/DD/YYYY")
                Else
                   If CDate(Tb![Check Date]) < CDate(AW(1).Text) Then
                     AW(1).Text = Format(Tb![Check Date], "MM/DD/YYYY")
                   End If
                End If
          Else
            AW(0).Text = "Closed"
            AW(1).Text = ""
          End If
       DbClose Db, Tb
If IsDate(AW(0).Text) And IsDate(AW(1).Text) Then
  If DateDiff("d", AW(0).Text, AW(1).Text) > 7 Then
     AW(0).ForeColor = vbRed
     AW(1).ForeColor = vbRed
     
  Else
     AW(0).ForeColor = vbBlack
     AW(1).ForeColor = vbBlack
  End If
End If
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Active Week")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub GetPayrollPeriods()
On Error GoTo ErrorHandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If

Dim SRC As Integer
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
          ConnectADODB TimeRecordsDb, Db
             Sql = "Select * from [Employee Records] Order By TRY_CAST([Shift Date] AS datetime) asc"
            ConnectADORS Sql, Db, Tb
            If Not Tb.EOF Then
              If Not Tb.BOF Then Tb.MoveFirst
                If Not IsDate(APR(0).Text) Then
                   APR(0).Text = Format(Tb![Shift date], "MM/DD/YYYY")
                Else
                   If CDate(Tb![Shift date]) < CDate(APR(0).Text) Then
                     APR(0).Text = Format(Tb![Shift date], "MM/DD/YYYY")
                   End If
                End If
               If Not Tb.EOF Then Tb.MoveLast
                If Not IsDate(APR(1).Text) Then
                   APR(1).Text = Format(Tb![Shift date], "MM/DD/YYYY")
                Else
                   If CDate(Tb![Shift date]) < CDate(APR(1).Text) Then
                     APR(1).Text = Format(Tb![Shift date], "MM/DD/YYYY")
                   End If
                End If
         Else
          APR(0).Text = "Closed"
          APR(1).Text = ""
         End If
       DbClose Db, Tb
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Active Payroll")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub GetActiveBatch()
On Error GoTo ErrorHandler
Dim SRC As Integer
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
'If EMVDeviceIsPresent Then
     ConnectADODB DebitDb, Db
        Sql = "Select * from [Debit] Order By [Date] asc"
     ConnectADORS Sql, Db, Tb
      If Not Tb.EOF Then
        If Not Tb.BOF Then Tb.MoveFirst
          If Not IsDate(ABT(0).Text) Then
             ABT(0).Text = Format(Tb![Date], "MM/DD/YYYY")
          Else
             If CDate(Tb![Date]) < CDate(ABT(0).Text) Then
               ABT(0).Text = Format(Tb![Date], "MM/DD/YYYY")
             End If
          End If
         If Not Tb.EOF Then Tb.MoveLast
          If Not IsDate(ABT(1).Text) Then
             ABT(1).Text = Format(Tb![Date], "MM/DD/YYYY")
          Else
             If CDate(Tb![Date]) < CDate(ABT(1).Text) Then
               ABT(1).Text = Format(Tb![Date], "MM/DD/YYYY")
             End If
          End If
      Else
        ABT(0).Text = "Closed"
        ABT(1).Text = ""
      End If
'End If
DbClose Db, Tb
If IsDate(ABT(0).Text) And IsDate(ABT(1).Text) Then
  If DateDiff("d", ABT(0).Text, ABT(1).Text) > 1 Then
     ABT(0).ForeColor = vbRed
     ABT(1).ForeColor = vbRed
     
  Else
     ABT(0).ForeColor = vbBlack
     ABT(1).ForeColor = vbBlack
  End If
End If
Exit Sub
ErrorHandler:
 ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Active Batch")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub DEPTTEND_Click()
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then

Else
    DepartmentTender SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
    Set SearchDates = Nothing
End If
End Sub
Private Sub dfs_Click()
Dim DeductSafeCash As Currency
DeductSafeCash = Val(ShowNumberPad("Enter Cash Amount", False, 0, 0))
If DeductSafeCash = 0 Then Exit Sub
End Sub
Private Sub EMain_Click()
If Can_Continue(EMain.Caption) = False Then Exit Sub
    Dim AMM As MaintenanceMenus
    Set AMM = New MaintenanceMenus
    AMM.EmployeeMaintenance LoggedOnAs(0).EMPName, LoggedOnAs(0).EmpAccessLevel, -1
    Set AMM = Nothing
End Sub
Public Sub CreateBarChart(viewer As Object, Data() As Variant, Labels() As Variant, Colors() As Variant, ChartTitle As String)
    Dim C As XYChart
    Set C = Cd.XYChart(440 * WidScale1024, 320 * TopScale1024, Cd.goldColor(), -1, 2)

    'Add a title box using 10 point Arial Bold font. Set the background color to
    'metallic blue (9999FF) Use a 1 pixel 3D border.
    Call C.addTitle(ChartTitle, "arialbd.ttf", 10).setBackground( _
        Cd.metalColor(&H9999FF), -1, 1)
    'Set the plotarea at (40, 40) and of 200 x 150 pixels in size
    Call C.setPlotArea(40, 40, 280 * WidScale1024, 250 * TopScale1024)
    'Add a multi-color bar chart layer using the given data and colors. Use a 1 pixel
    '3D border for the bars.
    Call C.addBarLayer3(Data, Colors).setBorderColor(-1, 1)
    'Set the labels on the x axis.
    Call C.xAxis().setLabels(Labels)
    'output the chart
    Set viewer.Picture = C.makePicture()
End Sub
Public Sub createChart(viewer As Object, Data() As Variant, Labels() As Variant, Colors() As Variant, ChartTitle As String)
    Dim C As PieChart
    Set C = Cd.PieChart(450 * WidScale1024, 300 * TopScale1024)
    Call C.setPieSize(225 * WidScale1024, 150 * TopScale1024, 75 * WidScale1024)
    Call C.set3D
    Call C.setColors(Cd.transparentPalette)
    Call C.setBackground(C.gradientColor(Cd.silverGradient), 0, 1)
    Call C.setLabelLayout(Cd.SideLayout)
    Call C.addTitle(ChartTitle, "timesbi.ttf", 7).setBackground( _
        Cd.metalColor(&HFF9999))
    Call C.setLabelStyle().setBackground(Cd.SameAsMainColor, Cd.Transparent, 1)
    Call C.setLineColor(Cd.SameAsMainColor, &H0)
    'Set the pie data and the pie labels
    Call C.SetData(Data, Labels)
    'Explode the 1st sector (index = 0)
    Call C.setExplode(0)
    'output the chart
    Set viewer.Picture = C.makePicture()
    Exit Sub
ERH:
Resume
End Sub
Private Sub CalDayLabor(WorkDept As String, Dept_Pay As Currency, WorkTime As Single, HowPaid As Integer, DaysWorked As Integer, OverTime As Single)
    On Error GoTo ERH
    Dim Cld As Integer
    For Cld = 0 To Total_depts
       If Trim(WorkDept) = Trim(Jobs(Cld).JobTitle) Then
        If HowPaid = 1 Then
            Jobs(Cld).JobRate = Jobs(Cld).JobRate + (Dept_Pay * WorkTime) + CalculateOT(Val(OverTime), CCur(Dept_Pay))
            Jobs(Cld).JobHours = Jobs(Cld).JobHours + WorkTime + CalculateOT(Val(OverTime), CCur(Dept_Pay))
        End If
        Exit For
       End If
    Next Cld
ERH:
End Sub
Private Sub Form_Activate()
On Error GoTo ERH
Set Cd = New ChartDirector.API
Cd.setLicenseCode "RDST-247P-CD3B-ZSV7-96F8-4DDA"
Dim DB1 As ADODB.Connection
Dim RS1 As ADODB.Recordset
Dim AllDeposits As Currency
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
ConnectADODB SvrPath.LocalTerminalPath & SystemOptionsDB, Db
ConnectADORS "Select * from [Job Titles] order by [Job Title] asc", Db, Tb
  Total_depts = RecordsCount(Tb)
  ReDim Jobs(Total_depts)
   DN = -1
   Do While Not Tb.EOF
      DN = DN + 1
      Jobs(DN).JobTitle = Check_Null(Tb(0))   ' Dept Name
      Tb.MoveNext
   Loop
DbClose Db, Tb
OfficeData(6).Text = Format(GetBusinessDay, "MM/DD/YYYY")
ConnectADODB FinancialDb, DB1
 Sql = "Select * from [Bank Deposits] "
ConnectADORS Sql, DB1, RS1
If Not RS1.EOF Then
   BKDP(0).Text = Format(RS1.fields("Bank 1").Value & "", "0.00")
   AllDeposits = Format(RS1.fields("Bank 1").Value & "", "0.00")
   
   BKDP(1).Text = Format(RS1.fields("Bank 2").Value & "", "0.00")
   AllDeposits = AllDeposits + Format(RS1.fields("Bank 2").Value & "", "0.00")
   
   BKDP(2).Text = Format(RS1.fields("Bank 3").Value & "", "0.00")
   AllDeposits = AllDeposits + Format(RS1.fields("Bank 3").Value & "", "0.00")
   
   BKDP(3).Text = Format(RS1.fields("Bank 4").Value & "", "0.00")
   AllDeposits = AllDeposits + Format(RS1.fields("Bank 4").Value & "", "0.00")
End If
DbClose DB1, RS1
SafeCash.Text = Format((GetAuditDeposits - AllDeposits) + GetAuditSkims, "0.00")
Me.Refresh
PollingTime = ""
GetActiveWeek
GetPayrollPeriods
GetActiveBatch
SafeAdjust.Text = Format(ReturnChangeFundAdj(AW(1)), "0.00")
Timer1_Timer
Timer1.Enabled = True
Exit Sub
ERH:
End Sub
Public Function HotScheduleExport() As String
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
'Dim ZipFile       As Zip
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
'Set ZipFile = New Zip
FileName = Format(Date, "_YYYYMMDD") & Format(Time, "_HHMMSS")
If CPANYID = "" Or CCEPTID = "" Or LOCID = "" Then
   HotScheduleExport = "No Setup Information"
   Exit Function
End If
 If ConnectADODB(SvrPath.TerminalPath & ClosingHistory, Db) = False Then
   HotScheduleExport = "Error Connecting to Check Stats"
   Exit Function
 End If
 Sql = "Delete from [Sales] where not isdate([Date])"
 Db.Execute Sql
 
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
 
 If ConnectADODB(ProductHistory, Db) = False Then
   HotScheduleExport = "Error Connecting to Product History"
   Exit Function
 End If
 
 Sql = "Select * from [Check Header]" & ReturnDateString(" where TRY_CAST([date] AS datetime) between '" & Format(CDate(CDate(SalesDate) + CDate(GetStartTime)), "yyyy-mm-dd hh:nn:ss") & "' and '" & Format(CDate(CDate(SetEndingDate(SalesDate)) + CDate(SetEndingTime)), "yyyy-mm-dd hh:nn:ss") & "'")
 
 If ConnectADORS(Sql, Db, RS) = False Then
    HotScheduleExport = "Error Connecting to Employee Records"
   Exit Function
 End If
 Dim MyTotal As Currency
 ff = FreeFile
 SalesDate = Format(SalesDate, "YYYY-MM-DD hh:mm:ss.000")
 Open SvrPath.TerminalPath & "Hot Schedules\" & "Sales" & FileName & ".TXT" For Output As ff
 Do While Not RS.EOF
   Print #ff, CPANYID & "|" & CCEPTID & "|" & LOCID & "|" & SalesDate & "|" & Format(RS.fields("Date").Value, "YYYY-MM-DD hh:mm:ss.000") & "|" & Format(((RS.fields("Check Total").Value - RS.fields("Tax Total").Value) + Abs(RS.fields("Discount Total").Value)) - Abs(RS.fields("Non Sales Total").Value), "0.00") & "|" & "1" & "|" & GetOrderDestNo(RS.fields("Order Destination").Value & "") & "|" & "1"
   RS.MoveNext
 Loop
 Close ff
 DbClose Db, RS
 If ConnectADODB(SvrPath.TerminalPath & "EMPLOYEE.MDB", Db) = False Then
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
'Set ZipFile = Nothing
'Dim FTP As MSaleData.Publish
'Set FTP = New MSaleData.Publish
'    FTP.SFTP FilePrefix, TXTFileInfo(2), TXTFileInfo(3), SvrPath.TerminalPath & "Hot Schedules\" & FilePrefix & FileName & ".ZIP", FilePrefix & FileName & ".ZIP", SvrPath.TerminalPath & "Hot Schedules DownLoad\Shifts.TXT"
'Set FTP = Nothing
'Set HSFTP = New SFTPLib.SFTP
'With HSFTP
'        .TimeOut = 120
'        .RemotePath = ""
'        .RemotePath = ""
'        .RemoteFile = ""
'        .SSHAuthMode = amPassword
'        .SSHUser = TXTFileInfo(2)'
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
See_Msg Err.Description, 0, "Error sending hot schedule files"
End Function

Public Function GetAuditDeposits() As Currency
On Error GoTo ErrorHandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Function
End If
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
  If ConnectADODB(SvrPath.TerminalPath & "Audits.mdb", Db) = False Then Exit Function
    Sql = "Select Distinct Sum([Bank Deposit]) as TBank,Sum([Skim]) as TSkim from [Audits] where [Close Out Day] = 0"
  If ConnectADORS(Sql, Db, Tb) = False Then Exit Function
   If Not Tb.EOF Then
    GetAuditDeposits = Val(Tb![TBank] & "") + Val(Tb![TSkim] & "")
   Else
    GetAuditDeposits = 0
   End If
  DbClose Db, Tb
Exit Function
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Edit Audits")
   Select Case ErrorRespose
     Case "Abort"
       Exit Function
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Function
Public Function GetAuditSkims() As Currency
Dim GA  As Integer
Dim ff  As Integer
Dim Reg As Integer
ff = FreeFile
GA = 0
GetAuditSkims = 0
  Open SvrPath.TerminalPath & "DAT Files\" & "CashDeposits.Dat" For Random Shared As ff Len = Len(Sdeposit)
     Do
       GA = GA + 1
       Get #ff, GA, Sdeposit
       GetAuditSkims = GetAuditSkims + Sdeposit.DepositAmt
     Loop Until GA >= LOF(ff) / Len(Sdeposit)
  Close ff
End Function
Private Sub RunLabor(LBRTotal As Currency)
If TotalNetSales = 0 Then Exit Sub
Dim TotalSales As Currency
If Check_Option(91) = 1 Then
   TotalSales = TotalNetSales + NetDiscounts
Else
   TotalSales = TotalNetSales
End If
Dim Data()
Dim Labels()
Dim Colors()
  With LaborData
    Data = Array(LBRTotal, ((TotalSales + OpenCk) - LBRTotal))   '
    Labels = Array("Labor", "Net Sales")
  End With
  createChart LD, Data(), Labels(), Colors(), "Sales vs. Labor"
End Sub
Private Sub CalLabor(BusinessDay As String)
LaborData.cols = 4
LaborData.rows = Total_depts + 2
With LaborData
.Clear
   .FixedCols = 1
   .FixedRows = 1
   .ColWidth(0) = 2500
   .ColWidth(1) = 1100
   If GetConfigInfo("Misc", "Display Office Screen Dollars", "Yes", "FUNCTIONS.INI") = "DISABLED" Then
    .ColWidth(2) = 0  '1100
   Else
    .ColWidth(2) = 1100
   End If
   .ColWidth(3) = 1100
   .TextMatrix(0, 0) = "Dept Name"
   .TextMatrix(0, 1) = "Hours"
   .TextMatrix(0, 2) = "Dollars"
   .TextMatrix(0, 3) = "Percent"
End With
On Error GoTo ErrorHandler
Dim PRTb        As ADODB.Recordset
Dim JCS         As Integer
Dim EDb         As ADODB.Connection
Dim Id          As ADODB.Recordset
Dim CHDB        As ADODB.Connection
Dim CHRS        As ADODB.Recordset
Dim Db          As ADODB.Connection
Dim Tb          As ADODB.Recordset
Dim Ck          As Integer
Dim WrkTime     As Single
Dim PRate       As Currency
Dim Total_Time  As Single
Dim Ttl_P       As Single
Dim Ttl_L       As Single
Dim TWRKTIME    As Single
Dim EmpTwt      As Single
Dim LabPer      As Single
Dim Pl          As Integer
Dim Pen         As Integer
Dim SearchName  As String
Dim TotalSales  As Currency
TWRKTIME = 0
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
OpenDb SvrPath.LocalTerminalPath & "Employee.mdb", EDb
  Sql = "Select * from [General Data] Where ISNULL(TRY_CONVERT(int,[Terminated]),0) = 0 order by [Last Name] asc"
ConnectADORS Sql, EDb, Id
OpenDb TimeRecordsDb, Db
Do While Not Id.EOF
SearchName = Id(0) & " " & Id(2)
Pen = 0: EmpTwt = 0
Sql = "Select * From [Job Codes] Where [First Name] = '" & Id(0) & "'"
Sql = Sql & " and [Last Name] = '" & Id(2) & "'"
ConnectADORS Sql, EDb, PRTb
     If Not PRTb.EOF Then PRTb.MoveLast
     ReDim PayRate(PRTb.RecordCount)
     If Not PRTb.BOF Then PRTb.MoveFirst
    JCS = 0
    Do While Not PRTb.EOF
         PayRate(JCS).JobCodes = PRTb![Job Code] & ""
         PayRate(JCS).JobRate = Val(PRTb![rate] & "")
         PayRate(JCS).JobHourly = Val(PRTb![Hourly] & "")
         JCS = JCS + 1
      PRTb.MoveNext
    Loop
    SetNothing PRTb
    Sql = "Select * From [Employee Records] Where [Employee Name] = '" & SearchName & "'"
    Sql = Sql & " AND TRY_CAST([Shift Date] AS datetime) + TRY_CAST([Time In] AS datetime) Between '" & CDate(BusinessDay) & " " & GetStartTime & "'"
    Sql = Sql & " and '" & SetEndingDate(CDate(BusinessDay)) & " " & SetEndingTime & "'"
    ConnectADORS Sql, Db, Tb
      If Not Tb.BOF Then Tb.MoveFirst
       Do While Not Tb.EOF
         Call CalculateTimeRecord(CStr(Tb(2)) & "", CStr(Tb(7)) & "", CStr(Tb(3)) & "", CStr(Tb![Break Out] & ""), CStr(Tb![Break In] & ""), CStr(Tb(6)) & "", Total_Time)
         'HW = Hoursworked
          For Ck = 0 To JCS - 1
            If Trim(Tb(1)) = Trim(PayRate(Ck).JobCodes) Then
              PRate = PayRate(Ck).JobRate 'PayRate(Ck).JobRate
              Exit For
            End If
          Next Ck
       WrkTime = ConvertTimeToMinutes(Total_Time)
       CalDayLabor Tb(1) & "", PRate, WrkTime, Abs(PayRate(Ck).JobHourly), 1, OvertimeHrs(EmpTwt, WrkTime)
       If IsNull(WrkTime) Then WrkTime = 0
       TWRKTIME = TWRKTIME + Val(WrkTime)
       EmpTwt = EmpTwt + Val(WrkTime)
       WrkTime = 0
       Tb.MoveNext
     Loop
  Id.MoveNext
Loop
DbClose Db, Tb
DbClose EDb, Id
If Check_Option(91) = 1 Then
   TotalSales = TotalNetSales + NetDiscounts
Else
   TotalSales = TotalNetSales
End If
For Pl = 0 To Total_depts - 1
    If Trim(Jobs(Pl).JobTitle) > "" And Jobs(Pl).JobRate > 0 And TotalSales > 0 Then
      LabPer = (Jobs(Pl).JobRate / TotalSales) * 100
    End If
      With LaborData
      .TextMatrix(Pl + 1, 0) = Jobs(Pl).JobTitle
      .TextMatrix(Pl + 1, 1) = Format(Jobs(Pl).JobHours, "0.00")
      .TextMatrix(Pl + 1, 2) = Format(Jobs(Pl).JobRate, "0.00")
      .TextMatrix(Pl + 1, 3) = Format(LabPer, "0.00")
      LabPer = 0
      Ttl_L = Ttl_L + Jobs(Pl).JobRate
     End With
     Jobs(Pl).JobHours = 0: Jobs(Pl).JobRate = 0
Next Pl
      If Val(Ttl_L) > 0 And TotalSales > 0 Then Ttl_P = Ttl_L / TotalSales * 100
      With LaborData
      .TextMatrix(Total_depts + 1, 0) = "Totals"
      .TextMatrix(Total_depts + 1, 2) = Format(Ttl_L, "0.00")
      .TextMatrix(Total_depts + 1, 3) = Format(Ttl_P, "0.00")
      .TextMatrix(Total_depts + 1, 1) = Format(TWRKTIME, "0.00")
      End With
DbClose CHDB, CHRS
If Val(Ttl_L & "") <> 0 Then
 RunLabor Val(Ttl_L)
End If
If Val(Input_Option(80)) > 0 And Ttl_P > 0 Then
  If (Ttl_P / (Val(Input_Option(80)))) * 100 > 100 Then
  LaborWatch.ProgressBarValue = 100
  Else
  LaborWatch.ProgressBarValue = (Ttl_P / (Val(Input_Option(80)))) * 100
  End If
  Select Case LaborWatch.ProgressBarValue
     Case Is < 50
       LaborWatch.ProgressBarColor = vbGreen
     Case 50 To 75
       LaborWatch.ProgressBarColor = vbYellow
     Case Else
       LaborWatch.ProgressBarColor = vbRed
       EmailAlerts "LaborAlert.Alt", "Labor Percentage Exceeded..." & vbCrLf & "The Current Labor Perentage Is " & Ttl_P
     End Select
End If
Ttl_L = 0
Ttl_P = 0
TWRKTIME = 0
Exit Sub
ErrorHandler:
If Err.Number = 9 Then
   Err.Clear
   Exit Sub
End If
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Calculate Labor Office Menu ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub Form_Deactivate()
Set Cd = Nothing
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
ResizeScreen1024 Me
If System_config = FullService Then
   SVRRPTS.Visible = True
Else
   SVRRPTS.Caption = "Cashier Deposits"
   SVRRPTS.Visible = True
End If
If Check_Option(89) = 1 Then
 TBLLAY.Visible = True
End If
'If Input_Option(45) = "Yes" Then
' SafeAdj.Visible = False
'Else
 SafeAdj.Visible = True
'End If
TerminalFailToRead = False
'If EMVDeviceIsPresent Then
  BPL(0).Visible = False
  BPL(1).Visible = False
  BPL(2).Visible = False
  ABT(0).Visible = False
  ABT(1).Visible = False
'End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Timing = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
ReleaseRegisterLock
End Sub
Private Function MakeMeaDec(DataIn As String) As String
 Dim SV() As String
 If InStr(DataIn, "/") > 0 Then
   SV = Split(DataIn, "/")
   MakeMeaDec = Format((Format(Val(SV(1)) / 100, "0.00") / Val(SV(0))), "0.00")
 Else
   MakeMeaDec = "." & DataIn
 End If
End Function
Private Function StripQ(DataIn As String) As String
        DataIn = Replace(DataIn, "$", "")
        If Left(DataIn, 1) = """" Then
         If Len(DataIn) - 2 < 0 Then
          StripQ = ""
         Else
          StripQ = Trim$(Mid(DataIn, 2, Len(DataIn) - 2))
         End If
        Else
          StripQ = Trim$(DataIn)
        End If
        If InStr(StripQ, "¢") > 0 Then
          StripQ = Replace(StripQ, "¢", "")
          StripQ = MakeMeaDec(StripQ)
        End If
End Function
Private Sub DownLoadFile()
If FileCheck(SvrPath.LocalTerminalPath & "INI Files\" & "ChampsImport.INI") = False Then Exit Sub
Dim IPDir           As String
Dim BatchFileName   As String
Dim ZipName         As String
Dim CSVFileName     As String
IPDir = GetConfigInfo("FTP", "Directory", "", "ChampsImport.INI")
If IPDir = "" Then IPDir = SvrPath.LocalTerminalPath
BatchFileName = GetConfigInfo("FTP", "Batch File Name", "", "ChampsImport.INI")
ZipName = GetConfigInfo("FTP", "Name of zip", "", "ChampsImport.INI")
CSVFileName = GetConfigInfo("FTP", "1st File Name", "", "ChampsImport.INI")
If CSVFileName = "" Then
   See_Msg "No File Name Found...", 5, "Cannot Continue."
   Exit Sub
End If
DeleteFile IPDir & ZipName
ChDir IPDir
Shell IPDir & BatchFileName, vbNormalFocus
See_Msg "Down Load complete....", 5, ""
Me.Refresh
End Sub
Private Function GetPriceValue(DataIn As String) As Currency
 Dim SV() As String
 If InStr(DataIn, "/") > 0 Then
   SV = Split(DataIn, "/")
   If InStr(SV(1), ".") = 0 Then
     GetPriceValue = Format((Format(Val(SV(1)) / 100, "0.00") / Val(SV(0))), "0.00")
   Else
     GetPriceValue = Format((Format(Val(SV(1)), "0.00") / Val(SV(0))), "0.00")
   End If
 Else
   GetPriceValue = Val(DataIn)
 End If
End Function
Private Sub HSFTP_SSHServerAuthentication(HostKey As String, FingerPrint As String, KeyAlgorithm As String, Accept As Boolean)
Accept = True
End Sub

Private Sub ImportMenuFile_Click()
On Error GoTo ERH
If FileCheck(SvrPath.LocalTerminalPath & "INI Files\" & "ChampsImport.INI") = False Then Exit Sub
DownLoadFile
If See_Msg("Import File?", 4, "") <> 6 Then Exit Sub
Dim IPDir           As String
Dim BatchFileName   As String
Dim ZipName         As String
Dim CSVFileName     As String
IPDir = GetConfigInfo("FTP", "Directory", "", "ChampsImport.INI")
If IPDir = "" Then IPDir = SvrPath.LocalTerminalPath
BatchFileName = GetConfigInfo("FTP", "Batch File Name", "", "ChampsImport.INI")
ZipName = GetConfigInfo("FTP", "Name of zip", "", "ChampsImport.INI")
CSVFileName = GetConfigInfo("FTP", "1st File Name", "", "ChampsImport.INI")
If CSVFileName = "" Then
   See_Msg "No File Name Found...", 5, "Cannot Continue."
   Exit Sub
End If
'Dim UnZip As MSaleData.Publish
'Set UnZip = New MSaleData.Publish
'    UnZip.ExtractZipFile ZipName, IPDir
'Set UnZip = Nothing
ChDir SvrPath.LocalTerminalPath
Dim Db          As ADODB.Connection
Dim RS          As ADODB.Recordset
Dim lDb         As ADODB.Connection
Dim TM          As Integer
Dim AL          As Integer
Dim Dstr        As String
Dim ff          As Integer
Dim IList       As String
Dim ListI()     As String
Dim L           As Long
Dim TR          As Long
Dim DptName     As String
Dim CPrice      As Currency
Dim RetailItem  As String
Dim RM          As Integer
Dim UPC_code    As String
Dim RawUPC      As String
Dim LZRO        As Boolean
Dim BCLen       As Integer
Dim TfoO_price  As Currency '2 for what ever
If ConnectADODB("MSaleData", lDb) = False Then Exit Sub
If ConnectADODB(SvrPath.LocalTerminalPath & "CRMENU.MDB", Db) = False Then Exit Sub
  Sql = "delete from [Menu Items] where [Quick Pick]= 2"
  Db.Execute Sql
  SetNothing Db
  'Compact_DataBase SvrPath.LocalTerminalPath, "CrMenu.mdb"
If ConnectADODB(SvrPath.LocalTerminalPath & "CRMENU.MDB", Db) = False Then Exit Sub

  Dstr = BuildSql_Table("[Menu Items]", Db, 1)
  Sql = Dstr & " Values ("
  ff = FreeFile
  L = -1
On Error Resume Next
CenterObject Me, Progression
Progression.Visible = True
Me.Refresh
TR = 0
BCLen = GetBarCodeLen
Open IPDir & CSVFileName For Input As ff
Do While Not EOF(ff)
 TR = TR + 1
 Line Input #ff, IList
Loop
Close ff
Open IPDir & CSVFileName For Input As ff
Do While Not EOF(ff)
  Line Input #ff, IList
   L = L + 1
   If L > 0 Then
    ListI = Split(IList, ",")
    If StripQ(ListI(12)) > "" Then
        RawUPC = StripQ(ListI(0))
        If Left(RawUPC, 1) = "2" Then
          UPC_code = Mid(RawUPC, 2, BCLen)
        Else
          UPC_code = RawUPC
        End If
        'UPC_code = "": LZRO = False
        'For RM = 1 To Len(RawUPC)
        '   If MID(RawUPC, RM, 1) = 0 And LZRO = False Then
        '   Else
        '     LZRO = True
        '     UPC_code = UPC_code & MID(RawUPC, RM, 1)
        '   End If
        'Next RM
        'If UPC_code = "" Then UPC_code = RawUPC
        
        DptName = GetConfigInfo("Department LookUp", Val(StripQ(ListI(16))), "Not On File", "ChampsImport.INI")  'Sales Tax Regular
        Progress.ProgressBarValue = (L / TR) * 100: DoEvents
        Sql = Dstr & " Values ("
        RetailItem = Replace(StripQ(ListI(12)), "'", "")
        RetailItem = Replace(RetailItem, """", "")
        
        Sql = Sql & "'" & RetailItem & "'"          'Menu Item
        Sql = Sql & ",'Retail'"                     'Category Name
        'CPrice
        'If InStr(RetailItem, "LEMON LIME") > 0 Then Stop
        If GetPriceValue(StripQ(ListI(79))) > 0 Then 'TPR = temp price reduction
         Sql = Sql & "," & Format(GetPriceValue(StripQ(ListI(79))), "0.00")
        ElseIf GetPriceValue(StripQ(ListI(59))) > 0 Then   'TPR = temp price reduction
         Sql = Sql & "," & Format(GetPriceValue(StripQ(ListI(59))), "0.00")
        Else
         Sql = Sql & "," & Format(GetPriceValue(StripQ(ListI(24))), "0.00")   'Price
        End If
        
        
        'If IsNumeric(StripQ(ListI(79))) Then 'TPR = temp price reduction
        ' Sql = Sql & "," & Format(StripQ(ListI(79)), "0.00")
        'ElseIf IsNumeric(StripQ(ListI(59))) Then   'TPR = temp price reduction
        ' Sql = Sql & "," & Format(StripQ(ListI(59)), "0.00")
        'Else
        ' Sql = Sql & "," & Format(StripQ(ListI(24)), "0.00")   'Price
        'End If
        Sql = Sql & "," & 9999                      'Count Down
        Sql = Sql & "," & 1                         'available Daily
        Sql = Sql & "," & 1                         'Taxable
        Sql = Sql & "," & 0                         'Expand Print
        Sql = Sql & "," & 0                         'Price Levels
        Sql = Sql & "," & 0                         'Remote Printing
        Sql = Sql & ",'None'"                       'Timed Events 1
        Sql = Sql & ",'None'"                       'Timed Events 2
        Sql = Sql & ",'None'"                       'Timed Events 3
        Sql = Sql & ",'None'"                       'Timed Events 4
        Sql = Sql & ",'None'"                       'Timed Events 5
        Sql = Sql & "," & 0                         'Timed Price 1
        Sql = Sql & "," & 0                         'Timed Price 2
        Sql = Sql & "," & 0                         'Timed Price 3
        Sql = Sql & "," & 0                         'Timed Price 4
        Sql = Sql & "," & 0                         'Timed Price 5
        Sql = Sql & ",'None'"                       'Modifier 1
        Sql = Sql & ",'None'"                       'Modifier 2
        Sql = Sql & ",'None'"                       'Modifier 3
        Sql = Sql & ",'None'"                       'Modifier 4
        Sql = Sql & ",'None'"                       'Modifier 5
        Sql = Sql & ",'None'"                       'Remote 1
        Sql = Sql & ",'None'"                       'Remote 2
        Sql = Sql & ",'None'"                       'Remote 3
        Sql = Sql & ",'None'"                       'Remote 4
        Sql = Sql & ",'None'"                       'Remote 5
        Sql = Sql & "," & Val(GetConfigInfo("Department", DptName, 0, "ChampsImport.INI"))   'Sales Tax Regular
        Sql = Sql & "," & 0                         'Special Tax
        Sql = Sql & "," & 0                         'Liquor Tax
        Sql = Sql & "," & 0                         'Use Modifiers
        Sql = Sql & "," & -1                         'PLU
        Sql = Sql & "," & 0                         'Option Mods
        Sql = Sql & "," & 0                         'Record No
        Sql = Sql & ",'" & DptName & "'"            'Department Name
        Sql = Sql & "," & 0                         'Mod Options 1
        Sql = Sql & "," & 0                         'Mod Options 2
        Sql = Sql & "," & 0                         'Mod Options 3
        Sql = Sql & "," & 0                         'Mod Options 4
        Sql = Sql & "," & 0                         'Mod Options 5
        Sql = Sql & ",'" & UPC_code & "'"           'Bar Code
        Sql = Sql & ",'None'"                       'Alt Print
        Sql = Sql & ",'None'"                       'Item Info
        Sql = Sql & "," & 0                         'Bar Tax
        Sql = Sql & "," & 0                         'Open Price
        If StripQ(ListI(110)) = "Y" Then
         Sql = Sql & "," & 1                         'Weight
        Else
         Sql = Sql & "," & 0                         'Weight
        End If
        Sql = Sql & "," & 0                         'Timed Price
        Sql = Sql & "," & 0                         'Cost
        Sql = Sql & ",'0'"                          'Fore Col
        Sql = Sql & ",'0'"                          'Back Col
        Sql = Sql & ",''"                           'Icon
        Sql = Sql & ",'None'"                       'Remote 6
        Sql = Sql & ",'None'"                       'Proceedure
        Sql = Sql & "," & 0                         'Price 1
        Sql = Sql & "," & 0                         'Price 2
        Sql = Sql & "," & 0                         'Price 3
        Sql = Sql & "," & 0                         'Price 4
        Sql = Sql & "," & 0                         'Price 5
        Sql = Sql & "," & 0                         'Minute Cose
        Sql = Sql & "," & 0                         'Price Lock
        Sql = Sql & "," & 0                         'Manager Access
        Sql = Sql & "," & 2                         'Quick Pick
        Sql = Sql & "," & 0                         'Quick Button
        Sql = Sql & "," & 0                         'Lrft
        Sql = Sql & "," & 0                         'Top
        Sql = Sql & "," & 0                         'Up Charge
        Sql = Sql & ",'None'"                       'Tare
        Sql = Sql & "," & 0                         'Print Ticket
        Sql = Sql & ",'None'"                       'Age Requirement
        Sql = Sql & "," & 0                         'Min Max Choice
        Sql = Sql & ",'None'"                       'Access
        Sql = Sql & "," & 0                         'Expires In
        Sql = Sql & ",'None'"                       'Shift Event 1
        Sql = Sql & ",'None'"                       'Shift Event 2
        Sql = Sql & ",'None'"                       'Shift Event 3
        Sql = Sql & ",'None'"                       'Shift Event 4
        Sql = Sql & ",'None'"                       'Shift Event 5
        Sql = Sql & "," & 0                         'Shift Price 1
        Sql = Sql & "," & 0                         'Shift Price 2
        Sql = Sql & "," & 0                         'Shift Price 3
        Sql = Sql & "," & 0                         'Shift Price 4
        Sql = Sql & "," & 0                         'Shift Price 5
        Sql = Sql & ",'None'"                       'Shift Mod 1
        Sql = Sql & ",'None'"                       'Shift Mod 2
        Sql = Sql & ",'None'"                       'Shift Mod 3
        Sql = Sql & ",'None'"                       'Shift Mod 4
        Sql = Sql & ",'None'"                       'Shift Mod 5
        Sql = Sql & "," & 0                         'Height
        Sql = Sql & "," & 0                         'Width
        Sql = Sql & "," & 0                         'Font Size
        Sql = Sql & "," & 0                         'EModShift 1
        Sql = Sql & "," & 0                         'EModShift 2
        Sql = Sql & "," & 0                         'PModShift 3
        Sql = Sql & "," & 0                         'PModShift 4
        Sql = Sql & "," & 0                         'PModShift 5
        Sql = Sql & "," & 0                         'P Mod Shift 1
        Sql = Sql & "," & 0                         'P Mod Shift 2
        Sql = Sql & "," & 0                         'P Mod Shift 3
        Sql = Sql & "," & 0                         'P Mod Shift 4
        Sql = Sql & "," & 0                         'P Mod Shift 5
        Sql = Sql & "," & 0                         'Open Mod
        Sql = Sql & "," & 0                         'On Line
        Sql = Sql & ",'None'"                       'Modifier 6
        Sql = Sql & "," & 0                         'Modifier 6 optional
        Sql = Sql & "," & 0                         'P Mod Shift 1
        Sql = Sql & "," & 0                         'P Mod Shift 2
        Sql = Sql & "," & 0                         'P Mod Shift 3
        Sql = Sql & "," & 0                         'P Mod Shift 4
        Sql = Sql & "," & 0 & ")"                   'p Mod Shift 5
        Db.Execute Sql

        Sql = "Delete from [Menu Links] Where [Menu Name] = '" & StripQ(ListI(3)) & "'"
        lDb.Execute Sql
     End If
   End If
Loop
Close ff
DbClose Db, Nothing
DbClose lDb, Nothing
Progression.Visible = False
Exit Sub
ERH:
  See_Msg Err.Description, 0, "Error In Download"
  Err.Clear
End Sub

Private Sub ListMenu_Click()
Dim Menu     As ADODB.Connection
Dim MRecords As ADODB.Recordset

ConnectADODB "MSaleData", Menu
  Sql = "Select * from [Menu Items] Where [menu item name] = '" & ListMenu.List(ListMenu.ListIndex) & "'"

ConnectADORS Sql, Menu, MRecords
If Not MRecords.EOF Then
   MenuName.Text = MRecords![Menu Item Name] & ""
   Inventory.Text = Val(MRecords![Quick Button] & "")
   CountDown.Text = GetCountDownQty(ListMenu.List(ListMenu.ListIndex))
End If
DbClose Menu, MRecords
End Sub
Private Sub MGRMENU_Click()
If Can_Continue(MRGMenu.Caption) = False Then Exit Sub
Transaction.Show
End Sub
Private Sub MMain_Click()
If Can_Continue(MMain.Caption) = False Then Exit Sub
    Dim AMM As MaintenanceMenus
    Set AMM = New MaintenanceMenus
    AMM.MenuMaintenance LoggedOnAs(0).EMPName, LoggedOnAs(0).EmpAccessLevel
    Set AMM = Nothing
End Sub
Private Sub OfficeData_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 6 And KeyAscii = 13 Then Timer1_Timer
End Sub

Private Sub PSheet_Click()
If Dir(SvrPath.LocalTerminalPath & "V10Prep.exe") = "" Then
   See_Msg "No PREP EXE Found", 5, ""
   Exit Sub
End If
Shell SvrPath.LocalTerminalPath & " ep.exe", vbNormalFocus

End Sub
Private Sub Quit_Click()
If TerminalFailToRead Then
  SETvariables
End If
Call MemoryCheck(Office_Menu)
End Sub
Private Sub ReviewLog_Click()
End Sub
Private Sub RPM_Click()
If Can_Continue(RPM.Caption) = False Then Exit Sub
Shell SvrPath.LocalTerminalPath & "V11RPM.EXE", vbMaximizedFocus
End Sub
Private Sub RSETWL_Click()
If Can_Continue(RSETWL.Caption) = False Then Exit Sub
Dim Db As ADODB.Connection
If See_Msg("Reset Customer Waiting List", 4, "Erases all seated customer on the List") = 6 Then
            ConnectADODB TableLayoutDb, Db
            Sql = "Update [Wait List] Set [day] = " & GetDateToNo(CDate(Date))
            Sql = Sql & " where [day] = 0 and [Seated] = 1;"
            Db.Execute Sql
            SetNothing Db
            See_Msg "Update Complete", 5, ""
End If
End Sub
Private Sub RSVATION_Click()
Reservations.Show 1
End Sub

Private Sub save_Click()
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
ConnectADODB FinancialDb, Db
 Sql = "Select * from [Bank Deposits] "
ConnectADORS Sql, Db, RS
If RS.EOF Then RS.AddNew
   RS.fields("Bank 1").Value = Format(BKDP(0).Text, "0.00")
   RS.fields("Bank 2").Value = Format(BKDP(1).Text, "0.00")
   RS.fields("Bank 3").Value = Format(BKDP(2).Text, "0.00")
   RS.fields("Bank 4").Value = Format(BKDP(3).Text, "0.00")
  RS.Update
DbClose Db, RS
See_Msg "Saved....", 5, "Done"
End Sub
Private Sub Sch_Click()
If Can_Continue(Sch.Caption) = False Then Exit Sub
Shell SvrPath.LocalTerminalPath & "EmpSchedule.exe"
End Sub

Private Sub Set86_Click()
Dim Menu     As ADODB.Connection
Dim MRecords As ADODB.Recordset
ConnectADODB "MSaleData", Menu
  Sql = "Select * from [Menu Items] Where [Quick Pick] = 0 order by [Menu Item Name] asc"
ConnectADORS Sql, Menu, MRecords
ListMenu.Clear
 
 Do While Not MRecords.EOF
    ListMenu.AddItem MRecords![Menu Item Name] & ""
    MRecords.MoveNext
 Loop
DbClose Menu, MRecords

CenterObject Me, MenuList
 MenuList.Visible = True
 MenuList.ZOrder 0
 Inventory.Text = 0
 CountDown.Text = 9999
 MenuName.Text = ""
End Sub
Private Sub ShowFinancial_Click()
If Can_Continue(FReports.Caption) = False Then Exit Sub
Financal_Reports.Show
End Sub
Private Sub SiteName_Click()
IpAddress.Text = IpAddress.List(SiteName.ListIndex)
User.Text = User.List(SiteName.ListIndex)
PW.Text = PW.List(SiteName.ListIndex)
End Sub
Private Sub SVRRPTS_Click()
If System_config = ExpressService Then
 If Can_Continue(SVRRPTS.Caption) = False Then Exit Sub
 'AuditDeposit.Show 1
 Dim AD As VER11SystemMaintenance.MaintenanceMenus
 Set AD = New VER11SystemMaintenance.MaintenanceMenus
     AD.AUDITdeposits LoggedOnAs(0).EmpLoggedOnAs, LoggedOnAs(0).EmpAccessLevel, ReturnSystemType
 Set AD = Nothing
Else
 If Can_Continue(SVRRPTS.Caption) = False Then Exit Sub
 Server_Report_Menu.Show
End If
End Sub
Private Sub TBLLAY_Click()
If Can_Continue(TBLLAY.Caption) = False Then Exit Sub
EditLayout = True
isManager = True
TableLayout.LoadTableScreen
TableLayout.Show 1
End Sub
Public Sub Timer1_Timer()
On Error GoTo ERH
Dim DB1         As ADODB.Connection
Dim RS1         As ADODB.Recordset
Dim AllDeposits As Currency
If ClosingProgess Then Exit Sub
If Val(Input_Option(24)) > 0 Then
   Timing = Timing + 1
   If Timing >= Val(Input_Option(24)) Then
    Timing = 0
    Unload Me
    Exit Sub
   End If
End If

If IsDate(Input_Option(46)) Then
 If Hour(Time) = Hour(CDate(Input_Option(46))) Then
   Trans_Log "Exit Office Screen Auto Closing System"
   See_Msg "Waiting For Auto Close", 5, "Software will exit"
   Unload Me
   Exit Sub
 End If
End If
If IsDate(PollingTime) Then
  If DateDiff("n", CDate(PollingTime), Now) < 5 Then Exit Sub
End If
If IsDate(Input_Option(46)) Then
 If Format(CDate(Input_Option(46)), "HH:MM") = Format(Time, "HH:MM") Then
   Quit_Click
   Exit Sub
 End If
End If
SafeAdjust.Text = Format(ReturnChangeFundAdj(AW(1)), "0.00")
If Check_Option(32) = 1 And Check_Option(17) = 0 Then
 If GetPendingOrders Then
    ButtonFunction = BTNFNCT.LoadPending
    Store_Grid Order.Ticket(0)
    Clear_Grid Order.Ticket(0)
    Order.Load_Grid
    NoLoadActivate = True
    PhoneOrders.Show
    Call MemoryCheck(PhoneOrders)
    DoEvents
    NoLoadActivate = False
 End If
End If 'cannot use pending orders with drive thru causes problem with check numbers
'See_Msg "Updating Sales Data................", 5, "One Moment Please"
Sleep 0
Timer1.Enabled = False
PollingTime = CStr(Now)
Office_Menu.Caption = "POS Office Reading Terminals for Sales....."
Sleep 0
Sales_Breakdown True, 0
OfficeData(0).Text = Format(TotalNetSales + OpenCk, "0.00")
OfficeData(1).Text = Format(NetDiscounts, "0.00")
OfficeData(2).Text = Format(VoidSales, "0.00")
OfficeData(3).Text = Format(WtdTotalNetSales, "0.00")
OfficeData(4).Text = Format(RegCounts.Transaction, "#")
If RegCounts.Transaction > 0 And TotalNetSales > 0 Then
  OfficeData(5).Text = Format(TotalNetSales / RegCounts.Transaction, "0.00")
End If
Dim GETSrvTime As Long
If Val(Input_Option(2)) > 0 Then
 'If FileCheck(SvrPath.TerminalPath & "ServeTimeAlert.ALT") = False Then
  GETSrvTime = GetAvgServingTime
  If Val(Input_Option(2)) > 0 And GETSrvTime > 0 Then
   If (GETSrvTime / (Val(Input_Option(2)))) * 100 > 100 Then
    ServeTimeWatch.ProgressBarValue = 100
   Else
    ServeTimeWatch.ProgressBarValue = (GETSrvTime / (Val(Input_Option(2)))) * 100
   End If
   Select Case ServeTimeWatch.ProgressBarValue
     Case Is < 50
       ServeTimeWatch.ProgressBarColor = vbGreen
     Case 50 To 75
       ServeTimeWatch.ProgressBarColor = vbYellow
     Case Else
       ServeTimeWatch.ProgressBarColor = vbRed
       EmailAlerts "ServeTimeAlert.ALT", "Serving Time Exceeded..." & vbCrLf & "The Current Average Serving Time Is " & GETSrvTime & " Minutes"
     End Select
   End If
 'End If
End If
If Val(Input_Option(78)) > 0 And VoidSales > 0 Then
 If FileCheck(SvrPath.TerminalPath & "VoidAlert.ALT") = False Then
  If (VoidSales / (Val(Input_Option(78)))) * 100 > 100 Then
  VoidWatch.ProgressBarValue = 100
  Else
  VoidWatch.ProgressBarValue = (VoidSales / (Val(Input_Option(78)))) * 100
  End If
  Select Case VoidWatch.ProgressBarValue
     Case Is < 50
       VoidWatch.ProgressBarColor = vbGreen
     Case 50 To 75
       VoidWatch.ProgressBarColor = vbYellow
     Case Else
       VoidWatch.ProgressBarColor = vbRed
       EmailAlerts "VoidAlert.ALT", "Void Dollars Exceeded..." & vbCrLf & "The Current Void Dollars Is " & Format(VoidSales, "0.00")
       'DeleteFile SvrPath.TerminalPath & "VoidAlert.ALT"
     End Select
  End If
End If
If Val(Input_Option(79)) > 0 And NetDiscounts > 0 Then
 If FileCheck(SvrPath.TerminalPath & "DiscountAlert.Alt") = False Then
  If (NetDiscounts / (Val(Input_Option(79)))) * 100 > 100 Then
   DiscWatch.ProgressBarValue = 100
  Else
   DiscWatch.ProgressBarValue = (NetDiscounts / (Val(Input_Option(79)))) * 100
  End If
  Select Case DiscWatch.ProgressBarValue
     Case Is < 50
       DiscWatch.ProgressBarColor = vbGreen
     Case 50 To 75
       DiscWatch.ProgressBarColor = vbYellow
     Case Else
       DiscWatch.ProgressBarColor = vbRed
       EmailAlerts "DiscountAlert.Alt", "Discount Dollars Exceeded..." & vbCrLf & "The Current Discount Dollars Is " & Format(NetDiscounts, "0.00")
       'DeleteFile SvrPath.TerminalPath & "DiscountAlert.ALT"
     End Select
 End If
End If
If InStr(OfficeData(6).Text, "/") > 0 Then
 If IsDate(OfficeData(6).Text) Then
  CalLabor OfficeData(6).Text
 End If
End If
Dim Td As Integer
Dim DeptReports() As DeptSale
Td = CurrentDeptSales(DeptReports())
Dim Data()
Dim Labels()
Dim Colors()
ReDim Data(Td)
ReDim Labels(Td)
ReDim Colors(Td)
Td = 0
For Printall = 0 To UBound(DeptReports, 1) - 1
    If Trim(DeptReports(Printall).Department & "") = "" Then Exit For
    If DeptReports(Printall).DollarsSold <> 0 Then
      Labels(Td) = DeptReports(Printall).Department
      Data(Td) = DeptReports(Printall).DollarsSold
      Td = Td + 1
    End If
Next Printall
createChart LaborBreakdown, Data(), Labels(), Colors(), "Department Sales"
ConnectADODB FinancialDb, DB1
 Sql = "Select * from [Bank Deposits] "
ConnectADORS Sql, DB1, RS1
If Not RS1.EOF Then
   BKDP(0).Text = Format(RS1.fields("Bank 1").Value & "", "0.00")
   AllDeposits = Format(RS1.fields("Bank 1").Value & "", "0.00")
   
   BKDP(1).Text = Format(RS1.fields("Bank 2").Value & "", "0.00")
   AllDeposits = AllDeposits + Format(RS1.fields("Bank 2").Value & "", "0.00")
   
   BKDP(2).Text = Format(RS1.fields("Bank 3").Value & "", "0.00")
   AllDeposits = AllDeposits + Format(RS1.fields("Bank 3").Value & "", "0.00")
   
   BKDP(3).Text = Format(RS1.fields("Bank 4").Value & "", "0.00")
   AllDeposits = AllDeposits + Format(RS1.fields("Bank 4").Value & "", "0.00")
End If
DbClose DB1, RS1
SafeCash.Text = Format((GetAuditDeposits - AllDeposits) + GetAuditSkims, "0.00")
Timer1.Enabled = True
Office_Menu.Caption = "POS Office Menu"
Sleep 0
Exit Sub
ERH:
Err.Clear
End Sub
Private Sub TRMain_Click()
If Can_Continue(TRMain.Caption) = False Then Exit Sub
Edit_Time = True
Time_Sheet.Show 1
Edit_Time = False
End Sub

Private Sub UPHS_Click()
See_Msg HotScheduleExport, 5, "Upload to Hot Schedules"

End Sub

Private Sub WeekClose_Click()
If Can_Continue(WeekClose.Caption) = False Then Exit Sub
Financal_Reports.Show
End Sub
