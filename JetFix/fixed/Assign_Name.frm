VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Assign_Name 
   Appearance      =   0  'Flat
   BackColor       =   &H002A231D&
   BorderStyle     =   0  'None
   ClientHeight    =   10755
   ClientLeft      =   195
   ClientTop       =   825
   ClientWidth     =   14955
   DrawMode        =   4  'Mask Not Pen
   FillColor       =   &H002A231D&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H002A231D&
   Icon            =   "Assign_Name.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MemberPage 
      Appearance      =   0  'Flat
      BackColor       =   &H00170F04&
      BorderStyle     =   0  'None
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
      Height          =   6465
      Left            =   -6720
      ScaleHeight     =   6465
      ScaleWidth      =   7320
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CheckBox APPLYPAY 
         Appearance      =   0  'Flat
         BackColor       =   &H002E241D&
         Caption         =   "Apply Charge Amount To Payments"
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
         Height          =   326
         Left            =   2160
         TabIndex        =   38
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3859
      End
      Begin VB.TextBox ChgAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Left            =   3480
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   1800
         Width           =   2475
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "Assign_Name.frx":FC8F
         Left            =   3480
         List            =   "Assign_Name.frx":FC9F
         TabIndex        =   34
         Text            =   "House Account"
         Top             =   1200
         Width           =   2460
      End
      Begin BTNENHLib4.BtnEnh sscommand5 
         Height          =   675
         Left            =   2640
         TabIndex        =   6
         Top             =   5400
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1191
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":FCD9
         textLT          =   "Assign_Name.frx":FD41
         textCT          =   "Assign_Name.frx":FD59
         textRT          =   "Assign_Name.frx":FD71
         textLM          =   "Assign_Name.frx":FD89
         textRM          =   "Assign_Name.frx":FDA1
         textLB          =   "Assign_Name.frx":FDB9
         textCB          =   "Assign_Name.frx":FDD1
         textRB          =   "Assign_Name.frx":FDE9
         colorBack       =   "Assign_Name.frx":FE01
         colorIntern     =   "Assign_Name.frx":FE2B
         colorMO         =   "Assign_Name.frx":FE55
         colorFocus      =   "Assign_Name.frx":FE7F
         colorDisabled   =   "Assign_Name.frx":FEA9
         colorPressed    =   "Assign_Name.frx":FED3
         HollowFrame     =   -1  'True
         LightDirection  =   8
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
         Office2003ColorNorth=   7493961
         Office2003ColorSouth=   5587767
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   7493961
         Office2003ColorSouthMouseOver=   5587767
         Office2003ColorBorderMouseOver=   11045503
         Office2003ColorNorthPressed=   11045503
         Office2003ColorSouthPressed=   11045503
         Office2003ColorBorderPressed=   11045503
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh20 
         Height          =   675
         Left            =   3720
         TabIndex        =   33
         Top             =   3120
         Width           =   2235
         _Version        =   589828
         _ExtentX        =   3942
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Charge Accounts"
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
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   94
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":FEFD
         textLT          =   "Assign_Name.frx":FF7B
         textCT          =   "Assign_Name.frx":FF93
         textRT          =   "Assign_Name.frx":FFAB
         textLM          =   "Assign_Name.frx":FFC3
         textRM          =   "Assign_Name.frx":FFDB
         textLB          =   "Assign_Name.frx":FFF3
         textCB          =   "Assign_Name.frx":1000B
         textRB          =   "Assign_Name.frx":10023
         colorBack       =   "Assign_Name.frx":1003B
         colorIntern     =   "Assign_Name.frx":10065
         colorMO         =   "Assign_Name.frx":1008F
         colorFocus      =   "Assign_Name.frx":100B9
         colorDisabled   =   "Assign_Name.frx":100E3
         colorPressed    =   "Assign_Name.frx":1010D
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
      Begin BTNENHLib4.BtnEnh BtnEnh21 
         Height          =   675
         Left            =   1320
         TabIndex        =   39
         Top             =   3120
         Width           =   2235
         _Version        =   589828
         _ExtentX        =   3942
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Check Exp Dates"
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
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":10137
         textLT          =   "Assign_Name.frx":101B5
         textCT          =   "Assign_Name.frx":101CD
         textRT          =   "Assign_Name.frx":101E5
         textLM          =   "Assign_Name.frx":101FD
         textRM          =   "Assign_Name.frx":10215
         textLB          =   "Assign_Name.frx":1022D
         textCB          =   "Assign_Name.frx":10245
         textRB          =   "Assign_Name.frx":1025D
         colorBack       =   "Assign_Name.frx":10275
         colorIntern     =   "Assign_Name.frx":1029F
         colorMO         =   "Assign_Name.frx":102C9
         colorFocus      =   "Assign_Name.frx":102F3
         colorDisabled   =   "Assign_Name.frx":1031D
         colorPressed    =   "Assign_Name.frx":10347
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh22 
         Height          =   675
         Left            =   1320
         TabIndex        =   40
         Top             =   4080
         Width           =   4635
         _Version        =   589828
         _ExtentX        =   8176
         _ExtentY        =   1191
         _StockProps     =   66
         Caption         =   "Email - Print and Pay Statements"
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
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":10371
         textLT          =   "Assign_Name.frx":10411
         textCT          =   "Assign_Name.frx":10429
         textRT          =   "Assign_Name.frx":10441
         textLM          =   "Assign_Name.frx":10459
         textRM          =   "Assign_Name.frx":10471
         textLB          =   "Assign_Name.frx":10489
         textCB          =   "Assign_Name.frx":104A1
         textRB          =   "Assign_Name.frx":104B9
         colorBack       =   "Assign_Name.frx":104D1
         colorIntern     =   "Assign_Name.frx":104FB
         colorMO         =   "Assign_Name.frx":10525
         colorFocus      =   "Assign_Name.frx":1054F
         colorDisabled   =   "Assign_Name.frx":10579
         colorPressed    =   "Assign_Name.frx":105A3
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
         Office2003ColorNorth=   9156159
         Office2003ColorSouth=   6259495
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   9156159
         Office2003ColorSouthMouseOver=   6259495
         Office2003ColorBorderMouseOver=   60534
         Office2003ColorNorthPressed=   7845449
         Office2003ColorSouthPressed=   7845449
         Office2003ColorBorderPressed=   60534
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Amount"
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
         Height          =   330
         Left            =   1200
         TabIndex        =   37
         Top             =   1800
         Width           =   2130
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Account Type"
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
         Height          =   330
         Left            =   600
         TabIndex        =   35
         Top             =   1200
         Width           =   2700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Account Information"
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
         Height          =   330
         Index           =   26
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   7320
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H009EF027&
         Height          =   4545
         Left            =   240
         Top             =   600
         Width           =   6840
      End
      Begin VB.Shape Shape20 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H002E241D&
         FillStyle       =   0  'Solid
         Height          =   4500
         Left            =   240
         Top             =   600
         Width           =   6810
      End
   End
   Begin VB.PictureBox AcctTypeSetup 
      Appearance      =   0  'Flat
      BackColor       =   &H00170F04&
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
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5505
      ScaleWidth      =   6690
      TabIndex        =   23
      Top             =   -5040
      Visible         =   0   'False
      Width           =   6720
      Begin BTNENHLib4.BtnEnh Command1 
         Height          =   615
         Left            =   1320
         TabIndex        =   182
         Top             =   3840
         Width           =   3855
         _Version        =   589828
         _ExtentX        =   6800
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Clean Database"
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
         Surface         =   2
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":105CD
         textLT          =   "Assign_Name.frx":10649
         textCT          =   "Assign_Name.frx":10661
         textRT          =   "Assign_Name.frx":10679
         textLM          =   "Assign_Name.frx":10691
         textRM          =   "Assign_Name.frx":106A9
         textLB          =   "Assign_Name.frx":106C1
         textCB          =   "Assign_Name.frx":106D9
         textRB          =   "Assign_Name.frx":10789
         colorBack       =   "Assign_Name.frx":107A1
         colorIntern     =   "Assign_Name.frx":107CB
         colorMO         =   "Assign_Name.frx":107F5
         colorFocus      =   "Assign_Name.frx":1081F
         colorDisabled   =   "Assign_Name.frx":10849
         colorPressed    =   "Assign_Name.frx":10873
      End
      Begin VB.ComboBox Default 
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
         ItemData        =   "Assign_Name.frx":1089D
         Left            =   3240
         List            =   "Assign_Name.frx":108AD
         TabIndex        =   59
         Text            =   "House Account"
         Top             =   3120
         Width           =   2938
      End
      Begin VB.TextBox BillAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3240
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox GPName 
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
         ItemData        =   "Assign_Name.frx":108E7
         Left            =   3240
         List            =   "Assign_Name.frx":108F7
         TabIndex        =   26
         Text            =   "House Account"
         Top             =   600
         Width           =   2938
      End
      Begin VB.ComboBox GPTender 
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
         ItemData        =   "Assign_Name.frx":10931
         Left            =   3240
         List            =   "Assign_Name.frx":10941
         TabIndex        =   25
         Text            =   "House Account"
         Top             =   1080
         Width           =   2938
      End
      Begin VB.ComboBox TipIsOn 
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
         ItemData        =   "Assign_Name.frx":1097B
         Left            =   3240
         List            =   "Assign_Name.frx":10985
         TabIndex        =   24
         Text            =   "Yes"
         Top             =   1560
         Width           =   2938
      End
      Begin BTNENHLib4.BtnEnh BtnEnh6 
         Height          =   645
         Left            =   4440
         TabIndex        =   27
         Top             =   4680
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":10992
         textLT          =   "Assign_Name.frx":109FA
         textCT          =   "Assign_Name.frx":10A12
         textRT          =   "Assign_Name.frx":10A2A
         textLM          =   "Assign_Name.frx":10A42
         textRM          =   "Assign_Name.frx":10A5A
         textLB          =   "Assign_Name.frx":10A72
         textCB          =   "Assign_Name.frx":10A8A
         textRB          =   "Assign_Name.frx":10AA2
         colorBack       =   "Assign_Name.frx":10ABA
         colorIntern     =   "Assign_Name.frx":10AE4
         colorMO         =   "Assign_Name.frx":10B0E
         colorFocus      =   "Assign_Name.frx":10B38
         colorDisabled   =   "Assign_Name.frx":10B62
         colorPressed    =   "Assign_Name.frx":10B8C
         HollowFrame     =   -1  'True
         LightDirection  =   8
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
         Office2003ColorNorth=   7493961
         Office2003ColorSouth=   5587767
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   7493961
         Office2003ColorSouthMouseOver=   5587767
         Office2003ColorBorderMouseOver=   11045503
         Office2003ColorNorthPressed=   11045503
         Office2003ColorSouthPressed=   11045503
         Office2003ColorBorderPressed=   11045503
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   645
         Left            =   2280
         TabIndex        =   28
         Top             =   4680
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":10BB6
         textLT          =   "Assign_Name.frx":10C1E
         textCT          =   "Assign_Name.frx":10C36
         textRT          =   "Assign_Name.frx":10C4E
         textLM          =   "Assign_Name.frx":10C66
         textRM          =   "Assign_Name.frx":10C7E
         textLB          =   "Assign_Name.frx":10C96
         textCB          =   "Assign_Name.frx":10CAE
         textRB          =   "Assign_Name.frx":10CC6
         colorBack       =   "Assign_Name.frx":10CDE
         colorIntern     =   "Assign_Name.frx":10D08
         colorMO         =   "Assign_Name.frx":10D32
         colorFocus      =   "Assign_Name.frx":10D5C
         colorDisabled   =   "Assign_Name.frx":10D86
         colorPressed    =   "Assign_Name.frx":10DB0
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
         Office2003ColorNorth=   9156159
         Office2003ColorSouth=   6259495
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   9156159
         Office2003ColorSouthMouseOver=   6259495
         Office2003ColorBorderMouseOver=   60534
         Office2003ColorNorthPressed=   7845449
         Office2003ColorSouthPressed=   7845449
         Office2003ColorBorderPressed=   60534
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh8 
         Height          =   645
         Left            =   120
         TabIndex        =   29
         Top             =   4680
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":10DDA
         textLT          =   "Assign_Name.frx":10E46
         textCT          =   "Assign_Name.frx":10E5E
         textRT          =   "Assign_Name.frx":10E76
         textLM          =   "Assign_Name.frx":10E8E
         textRM          =   "Assign_Name.frx":10EA6
         textLB          =   "Assign_Name.frx":10EBE
         textCB          =   "Assign_Name.frx":10ED6
         textRB          =   "Assign_Name.frx":10EEE
         colorBack       =   "Assign_Name.frx":10F06
         colorIntern     =   "Assign_Name.frx":10F30
         colorMO         =   "Assign_Name.frx":10F5A
         colorFocus      =   "Assign_Name.frx":10F84
         colorDisabled   =   "Assign_Name.frx":10FAE
         colorPressed    =   "Assign_Name.frx":10FD8
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
         Office2003ColorNorth=   1708747
         Office2003ColorSouth=   1708467
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   1708747
         Office2003ColorSouthMouseOver=   1708467
         Office2003ColorBorderMouseOver=   3024108
         Office2003ColorNorthPressed=   1182862
         Office2003ColorSouthPressed=   1182862
         Office2003ColorBorderPressed=   3024108
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Set Default Account Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00170F04&
         Height          =   330
         Left            =   120
         TabIndex        =   58
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Amount "
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
         Height          =   330
         Left            =   120
         TabIndex        =   41
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type Setup"
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
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tender Name Assigned"
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
         Height          =   330
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Accepted "
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
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00271D16&
         FillColor       =   &H008AD739&
         FillStyle       =   0  'Solid
         Height          =   945
         Left            =   120
         Top             =   2760
         Width           =   6420
      End
   End
   Begin MSComCtl2.MonthView AccountExp 
      Height          =   3210
      Left            =   9960
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   5662
      _Version        =   393216
      ForeColor       =   1511172
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16776960
      StartOfWeek     =   136970241
      TitleBackColor  =   1511172
      TitleForeColor  =   16777215
      CurrentDate     =   38012
   End
   Begin VB.Frame LSetup 
      Appearance      =   0  'Flat
      BackColor       =   &H00ECDED7&
      Caption         =   "Location Setup"
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
      Height          =   6015
      Left            =   -4920
      TabIndex        =   76
      Top             =   -1800
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox Clocation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Current Location"
         Height          =   255
         Left            =   1800
         TabIndex        =   85
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save List"
         Height          =   400
         Left            =   1920
         TabIndex        =   84
         Top             =   5520
         Width           =   1335
      End
      Begin VB.ListBox LocationList 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   83
         Top             =   1920
         Width           =   4455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   400
         Left            =   1800
         TabIndex        =   81
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add New"
         Height          =   400
         Left            =   3480
         TabIndex        =   80
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox LName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   79
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   400
         Left            =   3480
         TabIndex        =   77
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Location Name"
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
         Left            =   360
         TabIndex        =   82
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Location Names"
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
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.PictureBox TransferBalances 
      Appearance      =   0  'Flat
      BackColor       =   &H00170F04&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4845
      Left            =   -8280
      ScaleHeight     =   4815
      ScaleWidth      =   8685
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   8715
      Begin VB.ComboBox TargetCard 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3000
         TabIndex        =   13
         Top             =   1200
         Width           =   3552
      End
      Begin VB.ComboBox SourceCard 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3000
         TabIndex        =   12
         Top             =   2280
         Width           =   3552
      End
      Begin BTNENHLib4.BtnEnh BtnEnh16 
         Height          =   645
         Left            =   6360
         TabIndex        =   14
         Top             =   3960
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":11002
         textLT          =   "Assign_Name.frx":1106A
         textCT          =   "Assign_Name.frx":11082
         textRT          =   "Assign_Name.frx":1109A
         textLM          =   "Assign_Name.frx":110B2
         textRM          =   "Assign_Name.frx":110CA
         textLB          =   "Assign_Name.frx":110E2
         textCB          =   "Assign_Name.frx":110FA
         textRB          =   "Assign_Name.frx":11112
         colorBack       =   "Assign_Name.frx":1112A
         colorIntern     =   "Assign_Name.frx":11154
         colorMO         =   "Assign_Name.frx":1117E
         colorFocus      =   "Assign_Name.frx":111A8
         colorDisabled   =   "Assign_Name.frx":111D2
         colorPressed    =   "Assign_Name.frx":111FC
         HollowFrame     =   -1  'True
         LightDirection  =   8
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
         Office2003ColorNorth=   7493961
         Office2003ColorSouth=   5587767
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   7493961
         Office2003ColorSouthMouseOver=   5587767
         Office2003ColorBorderMouseOver=   11045503
         Office2003ColorNorthPressed=   11045503
         Office2003ColorSouthPressed=   11045503
         Office2003ColorBorderPressed=   11045503
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh17 
         Height          =   645
         Left            =   4200
         TabIndex        =   15
         Top             =   3960
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Transfer Balance"
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
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":11226
         textLT          =   "Assign_Name.frx":112A6
         textCT          =   "Assign_Name.frx":112BE
         textRT          =   "Assign_Name.frx":112D6
         textLM          =   "Assign_Name.frx":112EE
         textRM          =   "Assign_Name.frx":11306
         textLB          =   "Assign_Name.frx":1131E
         textCB          =   "Assign_Name.frx":11336
         textRB          =   "Assign_Name.frx":1134E
         colorBack       =   "Assign_Name.frx":11366
         colorIntern     =   "Assign_Name.frx":11390
         colorMO         =   "Assign_Name.frx":113BA
         colorFocus      =   "Assign_Name.frx":113E4
         colorDisabled   =   "Assign_Name.frx":1140E
         colorPressed    =   "Assign_Name.frx":11438
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Target Card"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   960
         TabIndex        =   22
         Top             =   1200
         Width           =   1950
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Source Card"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   960
         TabIndex        =   21
         Top             =   2280
         Width           =   1950
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSOLIDATE ACCOUNTS"
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
         Height          =   330
         Left            =   135
         TabIndex        =   20
         Top             =   480
         Width           =   8370
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Target Account Name"
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
         Height          =   330
         Left            =   960
         TabIndex        =   19
         Top             =   1680
         Width           =   1950
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Source Account Name"
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
         Height          =   330
         Left            =   960
         TabIndex        =   18
         Top             =   2880
         Width           =   1965
      End
      Begin VB.Label TargetName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3000
         TabIndex        =   17
         Top             =   1680
         Width           =   3555
      End
      Begin VB.Label SourceName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3000
         TabIndex        =   16
         Top             =   2760
         Width           =   3555
      End
      Begin VB.Shape Shape22 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H002E241D&
         FillStyle       =   0  'Solid
         Height          =   3675
         Left            =   120
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.Frame AccountbalanceEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H002E241D&
      Caption         =   "Adjust Account Balance"
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
      Height          =   4575
      Left            =   7800
      TabIndex        =   64
      Top             =   6480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox NBal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "0.00"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox CBal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox AdjustBal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   2760
         TabIndex        =   66
         Text            =   "0.00"
         Top             =   960
         Width           =   1815
      End
      Begin BTNENHLib4.BtnEnh BtnEnh33 
         Height          =   645
         Left            =   360
         TabIndex        =   71
         Top             =   2880
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Add Credit"
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
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   94
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":11462
         textLT          =   "Assign_Name.frx":114D6
         textCT          =   "Assign_Name.frx":114EE
         textRT          =   "Assign_Name.frx":11506
         textLM          =   "Assign_Name.frx":1151E
         textRM          =   "Assign_Name.frx":11536
         textLB          =   "Assign_Name.frx":1154E
         textCB          =   "Assign_Name.frx":11566
         textRB          =   "Assign_Name.frx":1157E
         colorBack       =   "Assign_Name.frx":11596
         colorIntern     =   "Assign_Name.frx":115C0
         colorMO         =   "Assign_Name.frx":115EA
         colorFocus      =   "Assign_Name.frx":11614
         colorDisabled   =   "Assign_Name.frx":1163E
         colorPressed    =   "Assign_Name.frx":11668
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
      Begin BTNENHLib4.BtnEnh BtnEnh34 
         Height          =   645
         Left            =   2520
         TabIndex        =   72
         Top             =   2880
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Deduct Credit"
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
         textCaption     =   "Assign_Name.frx":11692
         textLT          =   "Assign_Name.frx":1170C
         textCT          =   "Assign_Name.frx":11724
         textRT          =   "Assign_Name.frx":1173C
         textLM          =   "Assign_Name.frx":11754
         textRM          =   "Assign_Name.frx":1176C
         textLB          =   "Assign_Name.frx":11784
         textCB          =   "Assign_Name.frx":1179C
         textRB          =   "Assign_Name.frx":117B4
         colorBack       =   "Assign_Name.frx":117CC
         colorIntern     =   "Assign_Name.frx":117F6
         colorMO         =   "Assign_Name.frx":11820
         colorFocus      =   "Assign_Name.frx":1184A
         colorDisabled   =   "Assign_Name.frx":11874
         colorPressed    =   "Assign_Name.frx":1189E
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
         Office2003ColorNorth=   1708747
         Office2003ColorSouth=   1708467
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   1708747
         Office2003ColorSouthMouseOver=   1708467
         Office2003ColorBorderMouseOver=   3024108
         Office2003ColorNorthPressed=   1182862
         Office2003ColorSouthPressed=   1182862
         Office2003ColorBorderPressed=   3024108
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh35 
         Height          =   765
         Left            =   2520
         TabIndex        =   73
         Top             =   3600
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1349
         _StockProps     =   66
         Caption         =   "Close Window"
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
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":118C8
         textLT          =   "Assign_Name.frx":11940
         textCT          =   "Assign_Name.frx":11958
         textRT          =   "Assign_Name.frx":11970
         textLM          =   "Assign_Name.frx":11988
         textRM          =   "Assign_Name.frx":119A0
         textLB          =   "Assign_Name.frx":119B8
         textCB          =   "Assign_Name.frx":119D0
         textRB          =   "Assign_Name.frx":119E8
         colorBack       =   "Assign_Name.frx":11A00
         colorIntern     =   "Assign_Name.frx":11A2A
         colorMO         =   "Assign_Name.frx":11A54
         colorFocus      =   "Assign_Name.frx":11A7E
         colorDisabled   =   "Assign_Name.frx":11AA8
         colorPressed    =   "Assign_Name.frx":11AD2
         HollowFrame     =   -1  'True
         LightDirection  =   8
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
         Office2003ColorNorth=   7493961
         Office2003ColorSouth=   5587767
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   7493961
         Office2003ColorSouthMouseOver=   5587767
         Office2003ColorBorderMouseOver=   11045503
         Office2003ColorNorthPressed=   11045503
         Office2003ColorSouthPressed=   11045503
         Office2003ColorBorderPressed=   11045503
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh36 
         Height          =   765
         Left            =   360
         TabIndex        =   74
         Top             =   3600
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1349
         _StockProps     =   66
         Caption         =   "Save Adjustment"
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
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":11AFC
         textLT          =   "Assign_Name.frx":11B7A
         textCT          =   "Assign_Name.frx":11B92
         textRT          =   "Assign_Name.frx":11BAA
         textLM          =   "Assign_Name.frx":11BC2
         textRM          =   "Assign_Name.frx":11BDA
         textLB          =   "Assign_Name.frx":11BF2
         textCB          =   "Assign_Name.frx":11C0A
         textRB          =   "Assign_Name.frx":11C22
         colorBack       =   "Assign_Name.frx":11C3A
         colorIntern     =   "Assign_Name.frx":11C64
         colorMO         =   "Assign_Name.frx":11C8E
         colorFocus      =   "Assign_Name.frx":11CB8
         colorDisabled   =   "Assign_Name.frx":11CE2
         colorPressed    =   "Assign_Name.frx":11D0C
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
         Office2003ColorNorth=   9156159
         Office2003ColorSouth=   6259495
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   9156159
         Office2003ColorSouthMouseOver=   6259495
         Office2003ColorBorderMouseOver=   60534
         Office2003ColorNorthPressed=   7845449
         Office2003ColorSouthPressed=   7845449
         Office2003ColorBorderPressed=   60534
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New Balance"
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
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   1920
         Width           =   2385
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance"
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
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1440
         Width           =   2385
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Dollar Amount "
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
         Height          =   330
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   2385
      End
   End
   Begin VB.PictureBox AccountEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00170F04&
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
      Height          =   7935
      Left            =   1680
      ScaleHeight     =   7905
      ScaleWidth      =   5745
      TabIndex        =   43
      Top             =   1560
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox NewData 
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
         Height          =   405
         Index           =   1
         Left            =   360
         TabIndex        =   50
         Top             =   4440
         Width           =   5055
      End
      Begin VB.TextBox NewData 
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
         ForeColor       =   &H00404000&
         Height          =   405
         Index           =   0
         Left            =   360
         TabIndex        =   49
         Top             =   2760
         Width           =   5055
      End
      Begin BTNENHLib4.BtnEnh BtnEnh24 
         Height          =   660
         Left            =   360
         TabIndex        =   44
         Top             =   6960
         Width           =   5055
         _Version        =   589828
         _ExtentX        =   8916
         _ExtentY        =   1164
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":11D36
         textLT          =   "Assign_Name.frx":11D9E
         textCT          =   "Assign_Name.frx":11DB6
         textRT          =   "Assign_Name.frx":11DCE
         textLM          =   "Assign_Name.frx":11DE6
         textRM          =   "Assign_Name.frx":11DFE
         textLB          =   "Assign_Name.frx":11E16
         textCB          =   "Assign_Name.frx":11E2E
         textRB          =   "Assign_Name.frx":11E46
         colorBack       =   "Assign_Name.frx":11E5E
         colorIntern     =   "Assign_Name.frx":11E88
         colorMO         =   "Assign_Name.frx":11EB2
         colorFocus      =   "Assign_Name.frx":11EDC
         colorDisabled   =   "Assign_Name.frx":11F06
         colorPressed    =   "Assign_Name.frx":11F30
         HollowFrame     =   -1  'True
         LightDirection  =   8
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
         Office2003ColorNorth=   7493961
         Office2003ColorSouth=   5587767
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   7493961
         Office2003ColorSouthMouseOver=   5587767
         Office2003ColorBorderMouseOver=   11045503
         Office2003ColorNorthPressed=   11045503
         Office2003ColorSouthPressed=   11045503
         Office2003ColorBorderPressed=   11045503
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh26 
         Height          =   645
         Left            =   360
         TabIndex        =   54
         Top             =   4920
         Width           =   5055
         _Version        =   589828
         _ExtentX        =   8916
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Change Account Number"
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
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":11F5A
         textLT          =   "Assign_Name.frx":11FE4
         textCT          =   "Assign_Name.frx":11FFC
         textRT          =   "Assign_Name.frx":12014
         textLM          =   "Assign_Name.frx":1202C
         textRM          =   "Assign_Name.frx":12044
         textLB          =   "Assign_Name.frx":1205C
         textCB          =   "Assign_Name.frx":12074
         textRB          =   "Assign_Name.frx":1208C
         colorBack       =   "Assign_Name.frx":120A4
         colorIntern     =   "Assign_Name.frx":120CE
         colorMO         =   "Assign_Name.frx":120F8
         colorFocus      =   "Assign_Name.frx":12122
         colorDisabled   =   "Assign_Name.frx":1214C
         colorPressed    =   "Assign_Name.frx":12176
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
         Office2003ColorNorth=   45307
         Office2003ColorSouth=   39676
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   3129599
         Office2003ColorSouthMouseOver=   39676
         Office2003ColorBorderMouseOver=   4901375
         Office2003ColorNorthPressed=   3129599
         Office2003ColorSouthPressed=   3129599
         Office2003ColorBorderPressed=   3129599
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh27 
         Height          =   645
         Left            =   360
         TabIndex        =   56
         Top             =   3240
         Width           =   5055
         _Version        =   589828
         _ExtentX        =   8916
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Change Account Name"
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
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":121A0
         textLT          =   "Assign_Name.frx":12226
         textCT          =   "Assign_Name.frx":1223E
         textRT          =   "Assign_Name.frx":12256
         textLM          =   "Assign_Name.frx":1226E
         textRM          =   "Assign_Name.frx":12286
         textLB          =   "Assign_Name.frx":1229E
         textCB          =   "Assign_Name.frx":122B6
         textRB          =   "Assign_Name.frx":122CE
         colorBack       =   "Assign_Name.frx":122E6
         colorIntern     =   "Assign_Name.frx":12310
         colorMO         =   "Assign_Name.frx":1233A
         colorFocus      =   "Assign_Name.frx":12364
         colorDisabled   =   "Assign_Name.frx":1238E
         colorPressed    =   "Assign_Name.frx":123B8
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
         Office2003ColorNorth=   45307
         Office2003ColorSouth=   39676
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   3129599
         Office2003ColorSouthMouseOver=   39676
         Office2003ColorBorderMouseOver=   4901375
         Office2003ColorNorthPressed=   3129599
         Office2003ColorSouthPressed=   3129599
         Office2003ColorBorderPressed=   3129599
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.TextBox MEMONOTE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   53
         Top             =   6000
         Width           =   5055
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Memo / Notes"
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
         Height          =   375
         Left            =   360
         TabIndex        =   183
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Account Editor "
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
         Left            =   0
         TabIndex        =   55
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label OldData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   52
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label OldData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   51
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label NewAcctNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "New Master Account Number"
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
         Height          =   405
         Left            =   360
         TabIndex        =   48
         Top             =   4080
         Width           =   4965
      End
      Begin VB.Label NewAcct 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "New Account Name"
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
         Height          =   285
         Left            =   360
         TabIndex        =   47
         Top             =   2400
         Width           =   2445
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Master Account Number"
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
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   46
         Top             =   1560
         Width           =   3285
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   45
         Top             =   720
         Width           =   3045
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00E2821A&
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00170F04&
         FillStyle       =   0  'Solid
         Height          =   3255
         Left            =   360
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame AttachPhoto 
      Appearance      =   0  'Flat
      BackColor       =   &H002E241D&
      Caption         =   "Attach Photo"
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
      Height          =   4575
      Left            =   1920
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox InmatePic 
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
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   63
         Text            =   "None"
         Top             =   3240
         Width           =   3735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3480
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BTNENHLib4.BtnEnh BtnEnh30 
         Height          =   720
         Left            =   2040
         TabIndex        =   61
         Top             =   3720
         Width           =   1770
         _Version        =   589828
         _ExtentX        =   3122
         _ExtentY        =   1270
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
         CornerFactor    =   10
         Surface         =   15
         BackColorContainer=   16777215
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":123E2
         textLT          =   "Assign_Name.frx":1244A
         textCT          =   "Assign_Name.frx":12462
         textRT          =   "Assign_Name.frx":1247A
         textLM          =   "Assign_Name.frx":12492
         textRM          =   "Assign_Name.frx":124AA
         textLB          =   "Assign_Name.frx":124C2
         textCB          =   "Assign_Name.frx":124DA
         textRB          =   "Assign_Name.frx":124F2
         colorBack       =   "Assign_Name.frx":1250A
         colorIntern     =   "Assign_Name.frx":12534
         colorMO         =   "Assign_Name.frx":1255E
         colorFocus      =   "Assign_Name.frx":12588
         colorDisabled   =   "Assign_Name.frx":125B2
         colorPressed    =   "Assign_Name.frx":125DC
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
         Office2003ColorNorth=   9156159
         Office2003ColorSouth=   6259495
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   9156159
         Office2003ColorSouthMouseOver=   6259495
         Office2003ColorBorderMouseOver=   60534
         Office2003ColorNorthPressed=   7845449
         Office2003ColorSouthPressed=   7845449
         Office2003ColorBorderPressed=   60534
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh31 
         Height          =   720
         Left            =   120
         TabIndex        =   62
         Top             =   3720
         Width           =   1770
         _Version        =   589828
         _ExtentX        =   3122
         _ExtentY        =   1270
         _StockProps     =   66
         Caption         =   "Find Photo"
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
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":12606
         textLT          =   "Assign_Name.frx":1267A
         textCT          =   "Assign_Name.frx":12692
         textRT          =   "Assign_Name.frx":126AA
         textLM          =   "Assign_Name.frx":126C2
         textRM          =   "Assign_Name.frx":126DA
         textLB          =   "Assign_Name.frx":126F2
         textCB          =   "Assign_Name.frx":1270A
         textRB          =   "Assign_Name.frx":12722
         colorBack       =   "Assign_Name.frx":1273A
         colorIntern     =   "Assign_Name.frx":12764
         colorMO         =   "Assign_Name.frx":1278E
         colorFocus      =   "Assign_Name.frx":127B8
         colorDisabled   =   "Assign_Name.frx":127E2
         colorPressed    =   "Assign_Name.frx":1280C
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin VB.Image InmateId 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      ForeColor       =   &H80000008&
      Height          =   2620
      Left            =   11400
      TabIndex        =   156
      Top             =   8075
      Width           =   3435
      Begin VB.CheckBox delete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close Account    "
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
         Height          =   326
         Left            =   360
         TabIndex        =   160
         Top             =   1200
         Width           =   2190
      End
      Begin VB.CheckBox Suspend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Suspend Account         "
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
         Height          =   330
         Left            =   360
         TabIndex        =   159
         Top             =   1500
         Width           =   2310
      End
      Begin VB.TextBox AExpiration 
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
         ForeColor       =   &H00000080&
         Height          =   350
         Left            =   360
         TabIndex        =   158
         Text            =   "Never"
         Top             =   850
         Width           =   2100
      End
      Begin BTNENHLib4.BtnEnh SSCommand4 
         Height          =   645
         Left            =   360
         TabIndex        =   161
         Top             =   1920
         Width           =   2100
         _Version        =   589828
         _ExtentX        =   3704
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Remove Account"
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
         BackColorContainer=   16777215
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":12836
         textLT          =   "Assign_Name.frx":128B2
         textCT          =   "Assign_Name.frx":128CA
         textRT          =   "Assign_Name.frx":128E2
         textLM          =   "Assign_Name.frx":128FA
         textRM          =   "Assign_Name.frx":12912
         textLB          =   "Assign_Name.frx":1292A
         textCB          =   "Assign_Name.frx":12942
         textRB          =   "Assign_Name.frx":1295A
         colorBack       =   "Assign_Name.frx":12972
         colorIntern     =   "Assign_Name.frx":1299C
         colorMO         =   "Assign_Name.frx":129C6
         colorFocus      =   "Assign_Name.frx":129F0
         colorDisabled   =   "Assign_Name.frx":12A1A
         colorPressed    =   "Assign_Name.frx":12A44
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
         Office2003ColorNorth=   5398752
         Office2003ColorSouth=   5065682
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
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "END ACCOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002A231D&
         Height          =   255
         Left            =   360
         TabIndex        =   168
         Top             =   120
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00ECDED7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00DBD4CE&
         FillColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   -120
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Expiration Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   360
         TabIndex        =   157
         Top             =   575
         Width           =   2535
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00DBD4CE&
         Height          =   2175
         Left            =   0
         Top             =   480
         Width           =   3435
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   11400
      TabIndex        =   149
      Top             =   5230
      Width           =   3435
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECDED7&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         TabIndex        =   179
         Top             =   0
         Width           =   3495
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "ACTIVITY RANGE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002A231D&
            Height          =   285
            Left            =   360
            TabIndex        =   180
            Top             =   120
            Width           =   3075
         End
      End
      Begin BTNENHLib4.BtnEnh SSCommand10 
         Height          =   645
         Left            =   1800
         TabIndex        =   150
         ToolTipText     =   "Resets Charge Data and Rolls Account Balance Foward."
         Top             =   2040
         Width           =   1515
         _Version        =   589828
         _ExtentX        =   2672
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Reset Account Activity"
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
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":12A6E
         textLT          =   "Assign_Name.frx":12AFA
         textCT          =   "Assign_Name.frx":12B12
         textRT          =   "Assign_Name.frx":12B2A
         textLM          =   "Assign_Name.frx":12B42
         textRM          =   "Assign_Name.frx":12B5A
         textLB          =   "Assign_Name.frx":12B72
         textCB          =   "Assign_Name.frx":12B8A
         textRB          =   "Assign_Name.frx":12BA2
         colorBack       =   "Assign_Name.frx":12BBA
         colorIntern     =   "Assign_Name.frx":12BE4
         colorMO         =   "Assign_Name.frx":12C0E
         colorFocus      =   "Assign_Name.frx":12C38
         colorDisabled   =   "Assign_Name.frx":12C62
         colorPressed    =   "Assign_Name.frx":12C8C
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   645
         Left            =   120
         TabIndex        =   151
         ToolTipText     =   "Will Remove all Starting Balances Except Gift Cards."
         Top             =   2040
         Width           =   1635
         _Version        =   589828
         _ExtentX        =   2884
         _ExtentY        =   1138
         _StockProps     =   66
         Caption         =   "Reset Beginning Balance - Debit Balance Only"
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
         ButtonRaiseFactor=   0
         CaptionWordWrapPerc=   96
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":12CB6
         textLT          =   "Assign_Name.frx":12D6E
         textCT          =   "Assign_Name.frx":12D86
         textRT          =   "Assign_Name.frx":12D9E
         textLM          =   "Assign_Name.frx":12DB6
         textRM          =   "Assign_Name.frx":12DCE
         textLB          =   "Assign_Name.frx":12DE6
         textCB          =   "Assign_Name.frx":12DFE
         textRB          =   "Assign_Name.frx":12E16
         colorBack       =   "Assign_Name.frx":12E2E
         colorIntern     =   "Assign_Name.frx":12E58
         colorMO         =   "Assign_Name.frx":12E82
         colorFocus      =   "Assign_Name.frx":12EAC
         colorDisabled   =   "Assign_Name.frx":12ED6
         colorPressed    =   "Assign_Name.frx":12F00
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin MSComCtl2.DTPicker StartDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   153
         Top             =   840
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "/dd"
         Format          =   136970241
         CurrentDate     =   37998
      End
      Begin MSComCtl2.DTPicker EndDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   155
         Top             =   1560
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "mm/dd"
         Format          =   136970241
         CurrentDate     =   37998
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   154
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   330
         Left            =   360
         TabIndex        =   152
         Top             =   600
         Width           =   1515
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00DBD4CE&
         Height          =   2295
         Left            =   0
         Top             =   480
         Width           =   3435
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   7320
      TabIndex        =   113
      Top             =   6120
      Width           =   3975
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECDED7&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         TabIndex        =   173
         Top             =   0
         Width           =   3975
         Begin VB.Label Label37 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SPENDING RESTRICTIONS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002A231D&
            Height          =   375
            Left            =   240
            TabIndex        =   174
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.ComboBox Scan_Data 
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
         Index           =   2
         Left            =   360
         Style           =   1  'Simple Combo
         TabIndex        =   119
         Text            =   "0.00"
         Top             =   1050
         Width           =   1402
      End
      Begin VB.ComboBox Scan_Data 
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
         Index           =   1
         Left            =   360
         Style           =   1  'Simple Combo
         TabIndex        =   117
         Text            =   "0.00"
         Top             =   3260
         Width           =   1402
      End
      Begin VB.TextBox MealLimit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   360
         TabIndex        =   115
         Text            =   "0"
         Top             =   2050
         Width           =   960
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum spending amount for the open period. IE. If you reset accounts monthly, the period is 30 days."
         ForeColor       =   &H00463B36&
         Height          =   615
         Left            =   360
         TabIndex        =   172
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum number of transactions for the day"
         ForeColor       =   &H00463B36&
         Height          =   495
         Left            =   360
         TabIndex        =   171
         Top             =   2450
         Width           =   3135
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum spending amount for the day."
         ForeColor       =   &H00463B36&
         Height          =   255
         Left            =   360
         TabIndex        =   170
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   330
         Index           =   3
         Left            =   360
         TabIndex        =   118
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   116
         Top             =   750
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Meals Per Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   330
         Index           =   16
         Left            =   360
         TabIndex        =   114
         Top             =   1800
         Width           =   1710
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H00DBD4CE&
         Height          =   4575
         Left            =   0
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   11400
      TabIndex        =   88
      Top             =   840
      Width           =   3435
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Suspend"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   2040
         TabIndex        =   111
         Top             =   3655
         Width           =   975
      End
      Begin VB.TextBox SPCINT 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   110
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox SPCINT 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   108
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Suspend"
         ForeColor       =   &H80000008&
         Height          =   326
         Index           =   2
         Left            =   2040
         TabIndex        =   107
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox SPCINT 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   105
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Suspend"
         ForeColor       =   &H80000008&
         Height          =   326
         Index           =   1
         Left            =   2040
         TabIndex        =   104
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox SPCINT 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   360
         TabIndex        =   102
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Suspend"
         ForeColor       =   &H00000000&
         Height          =   326
         Index           =   0
         Left            =   2040
         TabIndex        =   101
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Op1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate #4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   109
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Label Op1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate #3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   106
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label Op1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   103
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label Op1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   100
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ALTERNATE ACCOUNT #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002A231D&
         Height          =   330
         Index           =   19
         Left            =   360
         TabIndex        =   99
         Top             =   120
         Width           =   3000
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00ECDED7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00DBD4CE&
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.ComboBox AccountBackUp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7440
      TabIndex        =   9
      Text            =   "None"
      Top             =   0
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.TextBox Address 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Index           =   7
      Left            =   5400
      TabIndex        =   0
      Top             =   10560
      Visible         =   0   'False
      Width           =   645
   End
   Begin BTNENHLib4.BtnEnh SSCommand1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2115
      _Version        =   589828
      _ExtentX        =   3731
      _ExtentY        =   1085
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
      CornerFactor    =   10
      Surface         =   15
      BackColorContainer=   16777215
      ButtonRaiseFactor=   0
      CaptionWordWrapPerc=   96
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Assign_Name.frx":12F2A
      textLT          =   "Assign_Name.frx":12F92
      textCT          =   "Assign_Name.frx":12FAA
      textRT          =   "Assign_Name.frx":12FC2
      textLM          =   "Assign_Name.frx":12FDA
      textRM          =   "Assign_Name.frx":12FF2
      textLB          =   "Assign_Name.frx":1300A
      textCB          =   "Assign_Name.frx":13022
      textRB          =   "Assign_Name.frx":1303A
      colorBack       =   "Assign_Name.frx":13052
      colorIntern     =   "Assign_Name.frx":1307C
      colorMO         =   "Assign_Name.frx":130A6
      colorFocus      =   "Assign_Name.frx":130D0
      colorDisabled   =   "Assign_Name.frx":130FA
      colorPressed    =   "Assign_Name.frx":13124
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
      Office2003ColorNorth=   9156159
      Office2003ColorSouth=   6259495
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   9156159
      Office2003ColorSouthMouseOver=   6259495
      Office2003ColorBorderMouseOver=   60534
      Office2003ColorNorthPressed=   7845449
      Office2003ColorSouthPressed=   7845449
      Office2003ColorBorderPressed=   60534
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh BtnEnh3 
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   2115
      _Version        =   589828
      _ExtentX        =   3731
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Print Activity"
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
      MovePictureOnPress=   0   'False
      ButtonRaiseFactor=   0
      CaptionWordWrapPerc=   96
      LogPixels       =   96
      TextureMode     =   0
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "Assign_Name.frx":1314E
      textLT          =   "Assign_Name.frx":131CA
      textCT          =   "Assign_Name.frx":131E2
      textRT          =   "Assign_Name.frx":131FA
      textLM          =   "Assign_Name.frx":13212
      textRM          =   "Assign_Name.frx":1322A
      textLB          =   "Assign_Name.frx":13242
      textCB          =   "Assign_Name.frx":1325A
      textRB          =   "Assign_Name.frx":13272
      colorBack       =   "Assign_Name.frx":1328A
      colorIntern     =   "Assign_Name.frx":132B4
      colorMO         =   "Assign_Name.frx":132DE
      colorFocus      =   "Assign_Name.frx":13308
      colorDisabled   =   "Assign_Name.frx":13332
      colorPressed    =   "Assign_Name.frx":1335C
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
      Office2003ColorNorth=   15523543
      Office2003ColorSouth=   14273476
      Office2003ColorBorder=   16777215
      Office2003ColorNorthMouseOver=   15523543
      Office2003ColorSouthMouseOver=   14273476
      Office2003ColorBorderMouseOver=   13782573
      Office2003ColorNorthPressed=   13147786
      Office2003ColorSouthPressed=   13147786
      Office2003ColorBorderPressed=   13782573
      Office2007OuterBorderWidth=   1
      Office2007InnerBorderWidth=   1
      Office2003BorderWidth=   1
   End
   Begin BTNENHLib4.BtnEnh BtnEnh4 
      Height          =   375
      Left            =   12600
      TabIndex        =   87
      Top             =   120
      Visible         =   0   'False
      Width           =   915
      _Version        =   589828
      _ExtentX        =   1614
      _ExtentY        =   661
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
      textCaption     =   "Assign_Name.frx":13386
      textLT          =   "Assign_Name.frx":133EE
      textCT          =   "Assign_Name.frx":13406
      textRT          =   "Assign_Name.frx":1341E
      textLM          =   "Assign_Name.frx":13436
      textRM          =   "Assign_Name.frx":1344E
      textLB          =   "Assign_Name.frx":13466
      textCB          =   "Assign_Name.frx":1347E
      textRB          =   "Assign_Name.frx":13496
      colorBack       =   "Assign_Name.frx":134AE
      colorIntern     =   "Assign_Name.frx":134D8
      colorMO         =   "Assign_Name.frx":13502
      colorFocus      =   "Assign_Name.frx":1352C
      colorDisabled   =   "Assign_Name.frx":13556
      colorPressed    =   "Assign_Name.frx":13580
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   7320
      TabIndex        =   89
      Top             =   840
      Width           =   3975
      Begin VB.CheckBox CreditBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prepaid balance only"
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
         Height          =   330
         Left            =   360
         TabIndex        =   112
         Top             =   720
         Value           =   1  'Checked
         Width           =   2970
      End
      Begin VB.ComboBox Scan_Data 
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
         Height          =   375
         Index           =   6
         Left            =   360
         Style           =   1  'Simple Combo
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Width           =   1395
      End
      Begin VB.ComboBox Scan_Data 
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
         ForeColor       =   &H0065713C&
         Height          =   375
         Index           =   5
         Left            =   2160
         Style           =   1  'Simple Combo
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3600
         Width           =   1402
      End
      Begin VB.ComboBox Scan_Data 
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
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   4
         Left            =   360
         Style           =   1  'Simple Combo
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3600
         Width           =   1402
      End
      Begin VB.ComboBox Scan_Data 
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
         Height          =   375
         Index           =   7
         Left            =   360
         Style           =   1  'Simple Combo
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   4420
         Width           =   1402
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00DBD4CE&
         Height          =   2775
         Left            =   -240
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Non-prepaid accounts work like a credit card; allowing negative balances to be paid at a later time."
         ForeColor       =   &H00463B36&
         Height          =   735
         Left            =   360
         TabIndex        =   175
         Top             =   1520
         Width           =   3015
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepaid accounts must have money on the account to pay for purchases."
         ForeColor       =   &H00463B36&
         Height          =   495
         Left            =   360
         TabIndex        =   169
         Top             =   1005
         Width           =   3015
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00DBD4CE&
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   3975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   97
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   95
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Charges"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   93
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   315
         Index           =   12
         Left            =   360
         TabIndex        =   91
         Top             =   4150
         Width           =   2160
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT BALANCE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002A231D&
         Height          =   330
         Left            =   360
         TabIndex        =   90
         Top             =   120
         Width           =   2295
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00ECDED7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00DBD4CE&
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   120
      TabIndex        =   120
      Top             =   840
      Width           =   7095
      Begin VB.ComboBox AcctType 
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
         ForeColor       =   &H00EB5325&
         Height          =   450
         ItemData        =   "Assign_Name.frx":135AA
         Left            =   360
         List            =   "Assign_Name.frx":135BA
         TabIndex        =   166
         Text            =   "House Account"
         Top             =   700
         Width           =   4290
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         TabIndex        =   140
         Top             =   6000
         Width           =   7095
         Begin VB.TextBox Address 
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
            Height          =   375
            Index           =   6
            Left            =   4440
            TabIndex        =   176
            Top             =   240
            Width           =   2025
         End
         Begin VB.TextBox Address 
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
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   144
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox Address 
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
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   143
            Top             =   960
            Width           =   2145
         End
         Begin VB.TextBox Address 
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
            Height          =   375
            Index           =   2
            Left            =   2640
            TabIndex        =   142
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Address 
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
            Height          =   375
            Index           =   3
            Left            =   4440
            TabIndex        =   141
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00552E08&
            Height          =   285
            Index           =   13
            Left            =   4440
            TabIndex        =   177
            Top             =   0
            Width           =   1140
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00552E08&
            Height          =   330
            Index           =   7
            Left            =   360
            TabIndex        =   148
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00552E08&
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   147
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00552E08&
            Height          =   285
            Index           =   9
            Left            =   2640
            TabIndex        =   146
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00552E08&
            Height          =   285
            Index           =   10
            Left            =   4440
            TabIndex        =   145
            Top             =   720
            Width           =   1290
         End
      End
      Begin VB.TextBox Memo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1800
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   139
         Top             =   7920
         Width           =   3960
      End
      Begin VB.ComboBox AttachDiscount 
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
         Height          =   375
         Left            =   360
         TabIndex        =   135
         Text            =   "None"
         Top             =   9360
         Width           =   2340
      End
      Begin VB.TextBox AnvDate 
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
         Height          =   375
         Left            =   360
         TabIndex        =   133
         Text            =   "None"
         Top             =   8640
         Width           =   2340
      End
      Begin VB.TextBox Bday 
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
         Height          =   375
         Left            =   360
         TabIndex        =   131
         Text            =   "None"
         Top             =   7920
         Width           =   2340
      End
      Begin VB.TextBox Address 
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
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   129
         Top             =   5400
         Width           =   2145
      End
      Begin VB.TextBox Address 
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
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   127
         Top             =   5400
         Width           =   2145
      End
      Begin VB.ComboBox Scan_Data 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   360
         Style           =   1  'Simple Combo
         TabIndex        =   125
         Top             =   4080
         Width           =   4455
      End
      Begin VB.ComboBox Scan_Data 
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   122
         Top             =   2880
         Width           =   4455
      End
      Begin BTNENHLib4.BtnEnh BtnEnh18 
         Height          =   645
         Left            =   5040
         TabIndex        =   136
         Top             =   2880
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1147
         _StockProps     =   66
         Caption         =   "Search by Name"
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
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":135F4
         textLT          =   "Assign_Name.frx":13670
         textCT          =   "Assign_Name.frx":13688
         textRT          =   "Assign_Name.frx":136A0
         textLM          =   "Assign_Name.frx":136B8
         textRM          =   "Assign_Name.frx":136D0
         textLB          =   "Assign_Name.frx":136E8
         textCB          =   "Assign_Name.frx":13700
         textRB          =   "Assign_Name.frx":13718
         colorBack       =   "Assign_Name.frx":13730
         colorIntern     =   "Assign_Name.frx":1375A
         colorMO         =   "Assign_Name.frx":13784
         colorFocus      =   "Assign_Name.frx":137AE
         colorDisabled   =   "Assign_Name.frx":137D8
         colorPressed    =   "Assign_Name.frx":13802
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   645
         Left            =   5040
         TabIndex        =   137
         Top             =   4225
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1147
         _StockProps     =   66
         Caption         =   "Change Account Data"
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
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":1382C
         textLT          =   "Assign_Name.frx":138B2
         textCT          =   "Assign_Name.frx":138CA
         textRT          =   "Assign_Name.frx":138E2
         textLM          =   "Assign_Name.frx":138FA
         textRM          =   "Assign_Name.frx":13912
         textLB          =   "Assign_Name.frx":1392A
         textCB          =   "Assign_Name.frx":13942
         textRB          =   "Assign_Name.frx":1395A
         colorBack       =   "Assign_Name.frx":13972
         colorIntern     =   "Assign_Name.frx":1399C
         colorMO         =   "Assign_Name.frx":139C6
         colorFocus      =   "Assign_Name.frx":139F0
         colorDisabled   =   "Assign_Name.frx":13A1A
         colorPressed    =   "Assign_Name.frx":13A44
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh BtnEnh15 
         Height          =   645
         Left            =   5040
         TabIndex        =   163
         Top             =   3550
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1147
         _StockProps     =   66
         Caption         =   "Search By Account #"
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
         textCaption     =   "Assign_Name.frx":13A6E
         textLT          =   "Assign_Name.frx":13AF4
         textCT          =   "Assign_Name.frx":13B0C
         textRT          =   "Assign_Name.frx":13B24
         textLM          =   "Assign_Name.frx":13B3C
         textRM          =   "Assign_Name.frx":13B54
         textLB          =   "Assign_Name.frx":13B6C
         textCB          =   "Assign_Name.frx":13B84
         textRB          =   "Assign_Name.frx":13B9C
         colorBack       =   "Assign_Name.frx":13BB4
         colorIntern     =   "Assign_Name.frx":13BDE
         colorMO         =   "Assign_Name.frx":13C08
         colorFocus      =   "Assign_Name.frx":13C32
         colorDisabled   =   "Assign_Name.frx":13C5C
         colorPressed    =   "Assign_Name.frx":13C86
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
         Office2003ColorNorth=   15523543
         Office2003ColorSouth=   14273476
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   15523543
         Office2003ColorSouthMouseOver=   14273476
         Office2003ColorBorderMouseOver=   13782573
         Office2003ColorNorthPressed=   13147786
         Office2003ColorSouthPressed=   13147786
         Office2003ColorBorderPressed=   13782573
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand14 
         Height          =   615
         Left            =   360
         TabIndex        =   164
         Top             =   1200
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Load Only"
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
         Picture         =   "C:\Users\ben\Desktop\Icons\filter1act.bmp"
         PictureXOffset  =   10
         PicturePosition =   4
         BackColorContainer=   15523543
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         dibPicture      =   "Assign_Name.frx":13CB0
         textCaption     =   "Assign_Name.frx":13EB6
         textLT          =   "Assign_Name.frx":13F28
         textCT          =   "Assign_Name.frx":13F40
         textRT          =   "Assign_Name.frx":13F58
         textLM          =   "Assign_Name.frx":13F70
         textRM          =   "Assign_Name.frx":13F88
         textLB          =   "Assign_Name.frx":13FA0
         textCB          =   "Assign_Name.frx":13FB8
         textRB          =   "Assign_Name.frx":13FD0
         colorBack       =   "Assign_Name.frx":13FE8
         colorIntern     =   "Assign_Name.frx":14012
         colorMO         =   "Assign_Name.frx":1403C
         colorFocus      =   "Assign_Name.frx":14066
         colorDisabled   =   "Assign_Name.frx":14090
         colorPressed    =   "Assign_Name.frx":140BA
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
         Office2003ColorNorth=   16015959
         Office2003ColorSouth=   14694958
         Office2003ColorBorder=   15523543
         Office2003ColorNorthMouseOver=   16015959
         Office2003ColorSouthMouseOver=   14694958
         Office2003ColorBorderMouseOver=   16711730
         Office2003ColorNorthPressed=   10625559
         Office2003ColorSouthPressed=   10625559
         Office2003ColorBorderPressed=   16711730
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   1
      End
      Begin BTNENHLib4.BtnEnh SSCommand15 
         Height          =   615
         Left            =   2520
         TabIndex        =   165
         Top             =   1200
         Width           =   2115
         _Version        =   589828
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Load All Accounts"
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
         BackColorContainer=   15523543
         MovePictureOnPress=   0   'False
         ButtonRaiseFactor=   0
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Assign_Name.frx":140E4
         textLT          =   "Assign_Name.frx":14166
         textCT          =   "Assign_Name.frx":1417E
         textRT          =   "Assign_Name.frx":14196
         textLM          =   "Assign_Name.frx":141AE
         textRM          =   "Assign_Name.frx":141C6
         textLB          =   "Assign_Name.frx":141DE
         textCB          =   "Assign_Name.frx":141F6
         textRB          =   "Assign_Name.frx":1420E
         colorBack       =   "Assign_Name.frx":14226
         colorIntern     =   "Assign_Name.frx":14250
         colorMO         =   "Assign_Name.frx":1427A
         colorFocus      =   "Assign_Name.frx":142A4
         colorDisabled   =   "Assign_Name.frx":142CE
         colorPressed    =   "Assign_Name.frx":142F8
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
         Office2003ColorNorth=   16024350
         Office2003ColorSouth=   16666919
         Office2003ColorBorder=   15523543
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
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TIP: Use the customer's phone number as the account number to prevent forgotten information."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F260B&
         Height          =   450
         Left            =   360
         TabIndex        =   181
         Top             =   4440
         Width           =   4095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT ACCOUNT TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002A231D&
         Height          =   285
         Left            =   360
         TabIndex        =   167
         Top             =   120
         Width           =   2955
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the dropdown under NAME to view all accounts associated with the selected Account Type."
         ForeColor       =   &H00463B36&
         Height          =   495
         Left            =   360
         TabIndex        =   162
         Top             =   3240
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Memo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Left            =   2880
         TabIndex        =   138
         Top             =   7680
         Width           =   3720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Attached Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   23
         Left            =   360
         TabIndex        =   134
         Top             =   9120
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Anniversary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   15
         Left            =   360
         TabIndex        =   132
         Top             =   8400
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   14
         Left            =   360
         TabIndex        =   130
         Top             =   7680
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cell #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   210
         Index           =   11
         Left            =   2640
         TabIndex        =   128
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   21
         Left            =   360
         TabIndex        =   126
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "* Main Account Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   22
         Left            =   360
         TabIndex        =   124
         Top             =   3840
         Width           =   4080
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER INFORMATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002A231D&
         Height          =   255
         Left            =   360
         TabIndex        =   123
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00552E08&
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   121
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Shape Shape9 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00DBD4CE&
         FillColor       =   &H00ECDED7&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   0
         Top             =   1920
         Width           =   7590
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00DBD4CE&
         FillColor       =   &H00ECDED7&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   7590
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00DBD4CE&
         Height          =   5505
         Left            =   0
         Top             =   1920
         Width           =   7530
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00DBD4CE&
         FillColor       =   &H00ECDED7&
         FillStyle       =   0  'Solid
         Height          =   2565
         Left            =   -2040
         Top             =   4320
         Width           =   7455
      End
   End
   Begin VB.Label NoAcct 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Number Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F7F5F4&
      Height          =   435
      Left            =   4440
      TabIndex        =   178
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H002A231D&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   8160
      Top             =   7440
      Width           =   915
   End
   Begin VB.Label MyLoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "My Location Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8280
      TabIndex        =   86
      Top             =   360
      Width           =   3435
   End
   Begin VB.Label SQLConnected 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "My Sql"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12600
      TabIndex        =   75
      Top             =   360
      Width           =   2205
   End
   Begin VB.Label LoadOpts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select  Load Options From Below"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002A231D&
      Height          =   210
      Left            =   7920
      TabIndex        =   57
      Top             =   10560
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fax"
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
      Height          =   300
      Index           =   17
      Left            =   4680
      TabIndex        =   10
      Top             =   10560
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "View Back Up Accounts"
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   20
      Left            =   480
      TabIndex        =   8
      Top             =   10560
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* Required Fields"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002A231D&
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   10560
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Menu Ext 
      Caption         =   "Exit"
   End
   Begin VB.Menu MyReports 
      Caption         =   "Reports"
      Begin VB.Menu PBD 
         Caption         =   "Print Birthdays"
      End
      Begin VB.Menu PAVS 
         Caption         =   "Print Anniversaries"
      End
      Begin VB.Menu PED 
         Caption         =   "Print Expiration Dates"
      End
      Begin VB.Menu LNAN 
         Caption         =   "List Names and Account #"
      End
   End
   Begin VB.Menu Util 
      Caption         =   "Utilities"
      Begin VB.Menu ATP 
         Caption         =   "Attach Photo"
      End
      Begin VB.Menu LSA 
         Caption         =   "Load Suspended Accounts"
      End
      Begin VB.Menu GAN 
         Caption         =   "Generate Account #"
      End
      Begin VB.Menu EAB 
         Caption         =   "Edit Account Balances"
      End
      Begin VB.Menu AEL 
         Caption         =   "Edit/Add Locations"
      End
      Begin VB.Menu LTTA 
         Caption         =   "Link Tenders To Accounts"
      End
      Begin VB.Menu CXPT 
         Caption         =   "Create Export File"
      End
      Begin VB.Menu PEACT 
         Caption         =   "Purge Expired Accounts"
      End
      Begin VB.Menu PZBGC 
         Caption         =   "Purge 0 Balance Gift Cards"
      End
      Begin VB.Menu ImportData 
         Caption         =   "Import Accounts"
      End
   End
   Begin VB.Menu ACSM 
      Caption         =   "Account Summary"
   End
   Begin VB.Menu ACTPAY 
      Caption         =   "Account Payment"
   End
   Begin VB.Menu ATCS 
      Caption         =   "Account Type Charges"
   End
   Begin VB.Menu PAS 
      Caption         =   "Print Account Statements"
   End
   Begin VB.Menu EMSTM 
      Caption         =   "Email Statements"
   End
   Begin VB.Menu TFBAL 
      Caption         =   "Transfer Balance"
   End
   Begin VB.Menu CRGEMAI 
      Caption         =   "Charge and Email"
   End
End
Attribute VB_Name = "Assign_Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StatementDate   As String
Private AddCredit       As Boolean
Private IAmAdmin        As Boolean
Private AdminLevel      As Integer
Private Function ValidDate(InDate As String) As String
  If IsDate(InDate) Then
    ValidDate = InDate
  Else
    ValidDate = "None"
  End If
End Function
Public Sub GetAccount()
On Error GoTo ErrorHandler
Dim R    As Integer
Dim SD   As Integer
If Scan_Data(0).Text = "" And Scan_Data(3).Text = "" Then Exit Sub
Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 0
Check1(3).Value = 0
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
If Scan_Data(0).Text = "" Then
If ConnectADORS("Select * from [HouseAccountData] Where [Acct Number] = '" & Scan_Data(3).Text & "'", Db, Tb) = False Then Exit Sub
Else
If ConnectADORS("Select * from [HouseAccountData] Where [Scan Name] = '" & Scan_Data(0).Text & "'", Db, Tb) = False Then Exit Sub
End If
    If Tb.EOF Then
        R = See_Msg("Information not found", 5, "")
        DbClose Db, Tb
        Exit Sub
    Else
    
    Scan_Data(1).Text = Format(Tb![Max Daily Amount], "0.00")
    Scan_Data(2).Text = Format(Tb![Max Weekly Amount], "0.00")
    Scan_Data(0).Text = Tb![scan Name] & ""
    Scan_Data(3).Text = Tb![Acct Number]
    Suspend.Value = Abs(Val(Tb![Suspend] & ""))
    delete.Value = Abs(Val(Tb![Delete Record] & ""))
    CreditBalance.Value = Abs(Val(Tb![Credit Balance] & ""))
        Address(0).Text = Tb![Address] & ""
        Address(1).Text = Tb![City] & ""
        Address(2).Text = Tb![State] & ""
        Address(3).Text = Tb![Zip] & ""
        Address(4).Text = Tb![Phone] & ""
        Address(5).Text = Tb![Cell] & ""
        Address(6).Text = Tb![Email] & ""
        Address(7).Text = Tb![Fax] & ""
        SPCINT(0).Text = Tb![Spc Interest 1] & ""
        If Left(Tb![Spc Interest 1] & "", 1) = "*" Then Check1(0).Value = 1
        SPCINT(1).Text = Tb![Spc Interest 2] & ""
        If Left(Tb![Spc Interest 1] & "", 2) = "*" Then Check1(1).Value = 1
        SPCINT(2).Text = Tb![Spc Interest 3] & ""
        If Left(Tb![Spc Interest 3] & "", 1) = "*" Then Check1(2).Value = 1
        SPCINT(3).Text = Tb![Spc Interest 4] & ""
        If Left(Tb![Spc Interest 4] & "", 1) = "*" Then Check1(3).Value = 1
        BDay.Text = ValidDate(Tb![Birthday] & "")
        AnvDate.Text = ValidDate(Tb![Anniversary] & "")
        Memo.Text = Tb![Memo] & ""
        MealLimit.Text = Val(Tb![Meal Limit] & "")
        If Tb![Discount] & "" > "" Then
         AttachDiscount.Text = Tb![Discount] & ""
        End If
        AExpiration.Text = ValidDate(Tb![Account Expires On] & "")
        If Tb![Account Type] & "" = "" Then
          AcctType.Text = "House Account"
        Else
          AcctType.Text = Tb![Account Type] & ""
        End If
    End If
    SetNothing Tb
    Sql = "Select Distinct sum([Adj Total]) as BalanceDue "
    Sql = Sql & ", sum([Tip Amount]) as CTip "
    Sql = Sql & " From [HouseCharges] where [charged to] = '" & Scan_Data(0).Text & "'"
    Sql = Sql & " and [Voided] = 0"
    
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    
    If Not Tb.EOF Then Scan_Data(4).Text = Format(Tb!BalanceDue, "0.00")   '+ Tb(1)
    SetNothing Tb
    Sql = "Select Distinct sum([Amount]) as BalanceDue "
    
    Sql = Sql & "From [HousePayments] where [Account Name] = '" & Scan_Data(0).Text & "'"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    If Not Tb.EOF Then Scan_Data(5).Text = Format(Tb!BalanceDue, "0.00")
    SetNothing Tb
    Sql = "Select Distinct sum([Beginning Balance]) as BegBal from [Starting Balance] Where [Account Name] = '" & Scan_Data(0).Text & "'"

    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub

    If Tb.EOF Then
     Scan_Data(7).Text = "0.00"
    Else
     Scan_Data(7).Text = Format(Val(Tb![BegBal] & ""), "0.00")
    End If
     Sql = "Select * from [Photos] Where [Account Name] = '" & Scan_Data(0).Text & "'"
     Sql = Sql & " and [Account Id] = '" & Scan_Data(3).Text & "'"
    If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
    If Not Tb.EOF Then
     Dim lSize As Long
     Dim varChunk() As Byte
     Dim lOffset As Long
     Dim sPath As String
     Dim nHandle As Integer
     Dim iChunks As Integer
     Dim nFragmentOffset As Integer
     Dim i As Integer
     Dim sFile As String
            nHandle = FreeFile
            sPath = App.path
            sFile = sPath & "\output.bin"
           Open sFile For Binary Access Write As nHandle
           lOffset = 0
           Do While lOffset < Tb.fields.Item("Photo").ActualSize
                ReDim varChunk(conChunkSize) As Byte
                varChunk() = Tb.fields.Item("Photo").GetChunk(conChunkSize)
                Put nHandle, , varChunk()
                lOffset = lOffset + conChunkSize
           Loop
       Close nHandle
       Set InmateId.Picture = LoadPicture(sFile, , vbLPColor)
       'BtnEnh29.ForeColor = &HFF00&
    Else
      InmateId.Picture = Nothing
      'BtnEnh29.ForeColor = &HFF&
    End If
    DbClose Db, Tb
    Scan_Data(6).Text = Format(Val(Scan_Data(7).Text) + Val(Scan_Data(4).Text) - Val(Scan_Data(5).Text), "0.00")
    CBal.Text = Format(Abs(Scan_Data(6).Text), "0.00")
    NBal.Text = "0.00"
    If Val(Scan_Data(6).Text) < 0 Then
     Label1(6).Caption = "Credit Balance"
    Else
     Label1(6).Caption = "Balance Owed"
    End If
    InmatePic.Text = "None"
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Account Assign Name ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub AccountBackUp_Click()
If AccountBackUp.Text <> "None" Then
 TurnFunctionsOff
 HouseAccountHistory = AccountBackUp.Text
Else
 TurnFunctionsOn
 HouseAccountHistory = "None"
End If
End Sub
Private Sub TurnFunctionsOn()
TFBAL.Enabled = True 'transfer balance
LNAN.Enabled = True 'list names and accounts
PZBGC.Enabled = True 'purge 0 balance
SSCommand10.Enabled = True  'reset account activity
BtnEnh1.Enabled = True     'change account data
PEACT.Enabled = True    'purge expired accounts
BtnEnh2.Enabled = True     'reset begin balance
LTTA.Enabled = True     'link tenders
End Sub
Private Sub TurnFunctionsOff()
TFBAL.Enabled = False
LNAN.Enabled = False
PZBGC.Enabled = False
SSCommand10.Enabled = False
BtnEnh1.Enabled = False
PEACT.Enabled = False
BtnEnh2.Enabled = False
LTTA.Enabled = False
End Sub

Private Sub AccountExp_DateClick(ByVal DateClicked As Date)
AccountExp.Visible = False
End Sub

Private Sub ACSM_Click()
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
End If
GETEXPDATE "Account Summary", False, False, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
End Sub
Private Sub ACTPAY_Click()
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
End If
tryagain:
Dim vDb As ADODB.Connection
Dim vRc As ADODB.Recordset
Dim VRs As ADODB.Recordset
Dim TotalExposure As Currency
If ConnectADODB(HouseAcct, vDb) = False Then
   Exit Sub
End If
    Sql = "Select Distinct [Payment Type], sum([Amount]) as BalanceDue "
    Sql = Sql & "From [HousePayments] where "
    Sql = Sql & " TRY_CAST([Date] AS datetime) between '" & CDate(SearchDates.Item(1).StartDate) & "' and '" & CDate(SearchDates.Item(1).EndDate) & "'"
    Sql = Sql & " Group By [Payment Type] Order by [Payment Type] asc"
If ConnectADORS(Sql, vDb, vRc) = False Then Exit Sub
If vRc.EOF Then
  Call See_Msg("Nothing to Print", 5, "")
  DbClose vDb, vRc
  Exit Sub
End If
SetPage ReportView.View, "Payments Summary Between " & SearchDates.Item(1).StartDate & " and " & SearchDates.Item(1).EndDate
 With ReportView.View
  .FontBold = False
  .FontSize = 8
  .TextAlign = taLeftBottom
 End With
 DrawLine
 TotalExposure = 0
 PrintLine "Payment Type", 4500
 PrintLine "Payment Amount", -1
 DrawLine
 Do While Not vRc.EOF
   PrintLine vRc![Payment Type] & "", 4500
   PrintLine Format(vRc!BalanceDue, "0.00"), -1
   TotalExposure = TotalExposure + vRc!BalanceDue
  vRc.MoveNext
 Loop
 DrawLine
 PrintLine "Total All Payments", 4500
 PrintLine Format(TotalExposure, "0.00") & "", -1
 ReportView.View.EndDoc
 ReportView.Show 1
DbClose vDb, vRc
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Assign Name BTn14 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub AEL_Click()
If HouseAcct <> "MSaleData" Then
  See_Msg "Sql Function Only", 5, ""
  Exit Sub
End If
If LoggedOnAs(0).EmpAccessLevel < AdminLevel Then
    See_Msg "Access To Low", 5, ""
End If
MsgBox "This Function is for Multi Unit Data Management", vbInformation, "Not for use with single site data"
Dim RC As ADODB.Recordset
Dim Db As ADODB.Connection
If ConnectADODB(HouseAcct, Db) Then
   Sql = "Select * from [Locations] order by [Location Name] asc "
 ConnectADORS Sql, Db, RC
   LName.Clear
   LocationList.Clear
   Do While Not RC.EOF
     LName.AddItem RC![Location Name] & ""
     LocationList.AddItem RC![Location Name] & ""
     If RC![Location Charge Privilege] = "True" Then
      LocationList.Selected(LocationList.NewIndex) = RC![Location Charge Privilege]
     End If
     RC.MoveNext
   Loop
End If
DbClose Db, RC
LName.Text = GetSetting("Micro$ale", "Register Setup", "Site")
Clocation.Value = 1
CenterObject Me, LSetup
LSetup.Visible = True
LSetup.ZOrder 0
End Sub
Private Sub AExpiration_Click()
If Can_Continue(Label1.Item(18).Caption) = False Then Exit Sub
    With AccountExp
       CenterObject Me, AccountExp
       .Value = Format(Date, "MM/DD/YYYY")
       .Visible = True
    End With
    Do While AccountExp.Visible = True
      DoEvents
    Loop
    If AccountExp.Value <> Format(Date, "MM/DD/YYYY") Then
       AExpiration.Text = Format(AccountExp.Value, "MM/DD/YYYY")
    Else
       AExpiration.Text = "Never"
    End If
End Sub
Private Sub AnvDate_Click()
If Can_Continue(Label1.Item(15).Caption) = False Then Exit Sub
With AccountExp
   CenterObject Me, AccountExp
   .Value = Format(Date, "MM/DD/YYYY")
   .Visible = True
End With
Do While AccountExp.Visible = True
  DoEvents
Loop
AnvDate.Text = Format(AccountExp.Value, "MM/DD")
End Sub
Private Sub ATCS_Click()
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
End If
tryagain:
Dim vDb As ADODB.Connection
Dim vRc As ADODB.Recordset
Dim VRs As ADODB.Recordset
Dim TotalExposure As Currency
If ConnectADODB(HouseAcct, vDb) = False Then
   Exit Sub
End If
    Sql = "Select Distinct [Account Type], sum([Adj total]) as BalanceDue "
    Sql = Sql & "From [HouseCharges] where [Voided] = 0"
    Sql = Sql & " and TRY_CAST([Check Date] AS datetime) between '" & CDate(SearchDates.Item(1).StartDate) & "' and '" & CDate(SearchDates.Item(1).EndDate) & "'"
    Sql = Sql & " Group By [Account Type] Order by [Account Type] asc"
If ConnectADORS(Sql, vDb, vRc) = False Then Exit Sub
If vRc.EOF Then
  Call See_Msg("Nothing to Print", 5, "")
  DbClose vDb, vRc
  Exit Sub
End If
SetPage ReportView.View, "Charge Summary Between " & SearchDates.Item(1).StartDate & " and " & SearchDates.Item(1).EndDate
 With ReportView.View
  .FontBold = False
  .FontSize = 8
  .TextAlign = taLeftBottom
 End With
 DrawLine
 TotalExposure = 0
 PrintLine "Account Type", 4500
 PrintLine "Amount Charged", -1
 DrawLine
 Do While Not vRc.EOF
   PrintLine vRc![Account Type] & "", 4500
   PrintLine Format(vRc!BalanceDue, "0.00"), -1
   TotalExposure = TotalExposure + vRc!BalanceDue
  vRc.MoveNext
 Loop
 DrawLine
 PrintLine "Total All Charges", 4500
 PrintLine "" & Format(TotalExposure, "0.00") & "", -1
 ReportView.View.EndDoc
 ReportView.Show 1
DbClose vDb, vRc
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Assign Name BTN13 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub ATP_Click()
CenterObject Me, AttachPhoto
AttachPhoto.Visible = True
AttachPhoto.ZOrder 0
End Sub
Private Sub AttachDiscount_Click()
 TrackHouseEvents LoggedOnAs(0).EMPName, "Discount Add/Edit " & Scan_Data(0).Text & " " & Scan_Data(3).Text & " Disc " & AttachDiscount.Text
End Sub
Private Sub Bday_Click()
If Can_Continue(Label1.Item(14).Caption) = False Then Exit Sub
With AccountExp
   .Font.Size = 15
   .Value = Format(Date, "MM/DD/YYYY")
   CenterObject Me, AccountExp
   .Visible = True
   .ZOrder 0
End With
Do While AccountExp.Visible = True
  DoEvents
Loop
BDay.Text = Format(AccountExp.Value, "MM/DD")
End Sub
Private Sub BtnEnh1_Click()
If Can_Continue(BtnEnh1.Caption) = False Then Exit Sub
If Scan_Data(0).Text = "" Then
  See_Msg "Select an Account", 5, ""
  Exit Sub
End If
OldData(0).Caption = Scan_Data(0).Text
OldData(1).Caption = Scan_Data(3).Text
NewData(0).Text = ""
NewData(1).Text = ""
MEMONOTE.Text = Memo.Text
CenterObject Me, AccountEdit
AccountEdit.Visible = True
AccountEdit.ZOrder 0
End Sub
Private Sub BtnEnh15_Click()
Dim AcctNo As String
Clear_data
Find_What = "Select Account"
Load_What = "Charge"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
AcctNo = GetSelected
GetSelected = ""
If AcctNo = "" Then Exit Sub
Scan_Data(0).Text = AcctNo
GetAccount
Me.Refresh
End Sub
Private Sub BtnEnh16_Click()
TransferBalances.Visible = False
End Sub
Private Sub BtnEnh17_Click()
If Can_Continue(BtnEnh17.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim RC As ADODB.Recordset
Dim SourceAcctName As String
Dim TargetAcctName As String
If TargetCard.Text = "" Then
  See_Msg "No Account Number", 5, "No Target Account"
  Exit Sub
End If
If SourceCard.Text = "" Then
  See_Msg "No Account Number", 5, "No Source Account"
  Exit Sub
End If

If See_Msg("Transfer Account Balance", 4, "Are You Sure ??") <> 6 Then Exit Sub
'may be bug here
Dim TargetAcctType  As String
Dim StartBal        As Currency
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & TargetCard.Text & "'"
If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
If RC.EOF Then
  See_Msg "Account Number Not Found", 5, ""
  DbClose Db, RC
  Exit Sub
Else
  TargetAcctType = RC![Account Type] & ""
  TargetAcctName = RC![scan Name] & ""
End If
SetNothing RC
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & SourceCard.Text & "'"
If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
If RC.EOF Then
  See_Msg "Account Number Not Found", 5, ""
  DbClose Db, RC
  Exit Sub
Else
  SourceAcctName = RC![scan Name] & ""
End If
SetNothing RC
Sql = "Select * from [HouseCharges] Where [Scan Number] = '" & SourceCard.Text & "'"
If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
Do While Not RC.EOF
   RC![Scan Number] = TargetCard.Text
   RC![Charged To] = TargetAcctName
   RC.Update
  RC.MoveNext
Loop
SetNothing RC
Sql = "Select * from [HousePayments] Where [Account Name] = '" & SourceAcctName & "'"
If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
 If Not RC.EOF Then
  Do While Not RC.EOF
   RC![Account Name] = TargetAcctName
   RC![Account Number] = TargetCard.Text
   RC.Update
   RC.MoveNext
 Loop
 End If
 Sql = "Select * from [Starting Balance] Where [Account Name] = '" & SourceAcctName & "'"
 StartBal = 0
 If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
 If Not RC.EOF Then
  StartBal = Val(RC![Beginning Balance])
  RC![Beginning Balance] = 0
  RC.Update
 End If
 SetNothing RC
 If StartBal <> 0 Then
    Sql = "Select * from [Starting Balance] Where [Account Name] = '" & TargetAcctName & "'"
    If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
     If Not RC.EOF Then
       RC![Beginning Balance] = Val(RC![Beginning Balance] & "") + StartBal
       RC.Update
     Else
      RC.AddNew
      RC![Account Type] = TargetAcctType
      RC![Account Name] = TargetAcctName
      RC![Account Number] = TargetCard.Text
      RC![Date] = Format(Date, "mm/dd/yyyy")
      RC![Beginning Balance] = StartBal
      RC.Update
     End If
 End If
DbClose Db, RC
TrackHouseEvents LoggedOnAs(0).EMPName, "Transfer Balance Target " & TargetCard.Text & " Transfer Amt " & StartBal
TrackHouseEvents LoggedOnAs(0).EMPName, "Transfer Balance Source " & SourceCard.Text
Call See_Msg("Account Transfered", 5, "")
SourceCard.Text = ""
SourceName.Caption = ""
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Consolidate Cards ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub

Public Sub BtnEnh18_Click()
Dim AcctNo As String
Clear_data
Find_What = "Select Account": Load_What = "Charge"
RequestListForm.Show 1
If GetSelected = "" Then Exit Sub
AcctNo = GetSelected
GetSelected = ""
If AcctNo = "" Then Exit Sub
Scan_Data(0).Text = AcctNo
GetAccount
End Sub
Public Function GenerateAccount(DataStr As String) As Long
Dim R As Integer
GenerateAccount = 0
For R = 1 To Len(DataStr)
   GenerateAccount = GenerateAccount + Asc(Mid(DataStr, R, 1))
Next R
End Function
Private Sub BtnEnh2_Click()
If Can_Continue(BtnEnh2.Caption) = False Then Exit Sub
If AcctType.Text = "Gift Card" Then
   See_Msg "Cannot Reset Gift Card Balances.", 5, "Not Allowed."
   Exit Sub
End If
DatePicker.StartDate.Caption = "Select a Ending Date"
DatePicker.StartingDate.Text = Format(Date, "MM/DD/YYYY")
DatePicker.EndingDate.Visible = False
DatePicker.EndDate.Visible = False
DatePicker.Show 1
If SearchDates Is Nothing Then
   Exit Sub
End If
If Not IsDate(SearchDates.Item(1).StartDate) Then Exit Sub
If See_Msg("Reset all Starting Balances Thru " & SearchDates.Item(1).StartDate, 4, "Credit Balance Accounts will NOT Be Reset") <> 6 Then Exit Sub
    Dim vDb As ADODB.Connection
    Dim RS  As ADODB.Recordset
      If ConnectADODB(HouseAcct, vDb) = False Then
        Exit Sub
      End If
      If UCase(AcctType.Text) <> "ALL" Then
        Sql = "Select * from [HouseAccountData] Where [Credit Balance] = 0  and [Account Type] = '" & AcctType.Text & "'"
      Else
        Sql = "Select * from [HouseAccountData] Where [Credit Balance] = 0 "
      End If
     ConnectADORS Sql, vDb, RS
      Do While Not RS.EOF
        Sql = "Delete from [Starting Balance] where [Account Name] = '" & RS![scan Name] & "'"
        Sql = Sql & " and TRY_CAST([date] AS datetime) Between '01/01/2000' and '" & CDate(SearchDates.Item(1).StartDate) & "'"
       vDb.Execute Sql
      RS.MoveNext
     Loop
    SetNothing vDb
    See_Msg "Completed....", 5, "File Purge Completed."
End Sub

Private Sub BtnEnh22_Click()
If Can_Continue(BtnEnh4.Caption) = False Then Exit Sub
If Combo3.Text = "" Then
   See_Msg "Select Account Type", 5, ""
   Exit Sub
End If
If Combo3.Text <> "ALL" Then
   If See_Msg("Statement for " & Combo3.Text & " Only", 4, "Continue ?") <> 6 Then Exit Sub
End If

If See_Msg("Have you Reviewed The Statements ?", 4, "") <> 6 Then Exit Sub

If See_Msg("Email All Statements", 4, "Send ?") <> 6 Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
End If
Dim LT As Integer
If CDate(SearchDates.Item(1).StartDate) > CDate(StartDate.Value) Then
  See_Msg "Invalid Start Date.", 0, "Current Period Start is " & Format(StartDate.Value, "MM/DD/YYYY")
  Exit Sub
End If

See_Msg "Sending Statements.....", 5, ""
ReDim FailedPayments(Scan_Data(0).ListCount)
 For LT = 0 To Scan_Data(0).ListCount - 1
  PrintStatments Combo3.Text, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, Scan_Data(0).List(LT), LT, "Pay Statements"
 Next LT
Dim Data$
Dim TPaid As Currency
Dim TDec  As Currency
Dim Cr    As Integer
ReportView.View.Orientation = orPortrait
SetPage ReportView.View, "Bill Report " & "From " & Format(SearchDates.Item(1).StartDate, "MM/DD/YY") & " To " & Format(SearchDates.Item(1).EndDate, "MM/DD/YY")
Dim F$
 F = "+2500|+>2000|+2000;"
 Dim H$
 TPaid = 0: TDec = 0
 H = "Name|Payment Amount|Payment Results|;"
 For Cr = 0 To Scan_Data(0).ListCount
   Data$ = Data$ & FailedPayments(Cr).AccountName & "|"
   Data$ = Data$ & Format(FailedPayments(Cr).AccountName, "0.00") & "|"
   Data$ = Data$ & FailedPayments(Cr).AccountName & "|;"
   TPaid = TPaid + FailedPayments(Cr).Paid
   TDec = TDec + FailedPayments(Cr).Failed
 Next Cr
Data$ = Data$ & "Total Paid" & "|"
Data$ = Data$ & Format(TPaid, "0.00") & "|"
Data$ = Data$ & "" & "|;"
Data$ = Data$ & "Total Declined" & "|"
Data$ = Data$ & Format(TDec, "0.00") & "|"
Data$ = Data$ & "" & "|;"
ReportView.View.TableBorder = tbBoxRows
ReportView.View.addTable F, H, Data, vbBlack, vbWhite
FinishJob 1
End Sub
Private Sub BtnEnh24_Click()
AccountEdit.Visible = False
AccountEdit.Left = -5000
End Sub
Private Sub BtnEnh25_Click()
End Sub
Private Sub BtnEnh26_Click()
If Can_Continue(BtnEnh26.Caption) = False Then Exit Sub
If NewData(1).Text = "" Then NewData(1).Text = OldData(1).Caption
On Error GoTo ErrorHandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
If See_Msg("Change Account Information ?", 4, "") <> 6 Then Exit Sub
Dim Rdb  As ADODB.Connection
Dim RRs  As ADODB.Recordset
Dim Sql  As String
If ConnectADODB(HouseAcct, Rdb) = False Then
  Exit Sub
End If

  Sql = "Select * from [HouseAccountData] "
  Sql = Sql & "Where [Acct Number] = '" & NewData(1).Text & "'"
  If ConnectADORS(Sql, Rdb, RRs) = False Then Exit Sub
  If RecordsCount(RRs) > 1 Then
    See_Msg "Duplicate Account Number Found!!!", 0, ""
    DbClose Rdb, RRs
    Exit Sub
  End If
  SetNothing RRs
  Sql = "Select * from [HouseAccountData] "
  Sql = Sql & " Where [Acct Number] = '" & OldData(1).Caption & "'"
  If ConnectADORS(Sql, Rdb, RRs) = False Then Exit Sub
  If Not RRs.EOF Then
    RRs![Acct Number] = NewData(1).Text
   RRs.Update
  End If
  SetNothing RRs
  Sql = "Update [HouseCharges] set [Scan Number] = '" & NewData(1).Text & "'"
  Sql = Sql & " Where [Scan Number] = '" & OldData(1).Caption & "'"
  Rdb.Execute Sql
  Sql = "Update HouseChargeDetail set [Account Number] = '" & NewData(1).Text & "'"
  Sql = Sql & " Where [Account Number] = '" & OldData(1).Caption & "'"
  Rdb.Execute Sql
  Sql = "Update [Starting Balance] set [Account Number] = '" & NewData(1).Text & "'"
  Sql = Sql & " Where [Account Number] = '" & OldData(1).Caption & "'"
  Rdb.Execute Sql
  Sql = "Update [HousePayments] set [Account Number] = '" & NewData(1).Text & "'"
  Sql = Sql & " Where [Account Number] = '" & OldData(1).Caption & "'"
  Rdb.Execute Sql
DbClose Rdb, RRs
TrackHouseEvents LoggedOnAs(0).EMPName, "Account Number Change OLD " & OldData(1).Caption & " " & OldData(0).Caption
TrackHouseEvents LoggedOnAs(0).EMPName, "Account Number Change NEW " & NewData(1).Text & " " & NewData(0).Text
See_Msg "Update complete.....", 5, ""
If NewData(1).Text > "" Then
  Scan_Data(3).AddItem NewData(1).Text
  Scan_Data(3).Text = NewData(1).Text
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Calculate Balance ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub BtnEnh27_Click()
If Can_Continue(BtnEnh27.Caption) = False Then Exit Sub
If NewData(0).Text = "" Then NewData(0).Text = OldData(0).Caption
On Error GoTo ErrorHandler
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
  Exit Sub
End If
If See_Msg("Change Account Information ?", 4, "") <> 6 Then Exit Sub
Dim Rdb  As ADODB.Connection
Dim RRs  As ADODB.Recordset
Dim Sql  As String
 If ConnectADODB(HouseAcct, Rdb) = False Then
  Exit Sub
 End If
 Sql = "Select * from [HouseAccountData] "
 Sql = Sql & "Where [Scan Name] = '" & NewData(0).Text & "'"
 If ConnectADORS(Sql, Rdb, RRs) = False Then Exit Sub
  If RecordsCount(RRs) > 0 Then
    See_Msg "Duplicate Account Name Found!!!", 0, ""
    DbClose Rdb, RRs
    Exit Sub
  End If
  SetNothing RRs
  Sql = "Select * from [HouseAccountData] "
  Sql = Sql & " Where [Scan Name] = '" & OldData(0).Caption & "'"
  If ConnectADORS(Sql, Rdb, RRs) = False Then Exit Sub
  If Not RRs.EOF Then
    RRs![Card Number] = ""
    RRs![Card Expires] = ""
    RRs![scan Name] = NewData(0).Text
   RRs.Update
  End If
  SetNothing RRs
  Sql = "Update [HouseCharges] set [Charged To] = '" & NewData(0).Text & "'"
  Sql = Sql & " Where [Charged To] = '" & OldData(0).Caption & "'"
  Rdb.Execute Sql
  Sql = "Update HouseChargeDetail set [Account Name] = '" & NewData(0).Text & "'"
  Sql = Sql & " Where [Account Name] = '" & OldData(0).Caption & "'"
  Rdb.Execute Sql
  Sql = "Update [HousePayments] set [Account Name] = '" & NewData(0).Text & "'"
  Sql = Sql & " Where [Account Name] = '" & OldData(0).Caption & "'"
  Rdb.Execute Sql
  Sql = "Update [Starting Balance] set [Account Name] = '" & NewData(0).Text & "'"
  Sql = Sql & " Where [Account Name] = '" & OldData(0).Caption & "'"
  Rdb.Execute Sql

DbClose Rdb, RRs
If NewData(0).Text > "" Then
  Scan_Data(0).AddItem NewData(0).Text
  Scan_Data(0).Text = NewData(0).Text
End If
TrackHouseEvents LoggedOnAs(0).EMPName, "Account Name Change OLD " & OldData(1).Caption & " " & OldData(0).Caption
TrackHouseEvents LoggedOnAs(0).EMPName, "Account Name Change NEW " & NewData(1).Text & " " & NewData(0).Text
See_Msg "Update complete.....", 5, "Name Update Done."
'LoadAccounts
Me.ZOrder 0
'Me.Visible = True
'Me.Refresh
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Change Name")
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
If AccountBackUp.Text <> "None" Then
  See_Msg "Print Statements Only", 5, ""
  Exit Sub
End If
If Scan_Data(0).Text = "" Then
  See_Msg "Please Select Account Name", 5, ""
  Exit Sub
End If
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
Else
  ReDim FailedPayments(1)
  PrintStatments AcctType.Text, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, Scan_Data(0).Text, 0, "Account Activity"
  Set SearchDates = Nothing
End If
End Sub

Private Sub BtnEnh30_Click()
AttachPhoto.Visible = False
AttachPhoto.Left = -5000
End Sub

Private Sub BtnEnh31_Click()
With CommonDialog1
     .ShowSave
     If .FileName > "" Then
         InmatePic.Text = .FileName
         'InmatePic.Text = File1.FileName
     End If
End With
On Error Resume Next
If InmatePic.Text = "None" Or InmatePic.Text = "" Then

Else
 InmateId.Picture = LoadPicture(InmatePic.Text)
End If
ChDir SvrPath.LocalTerminalPath
End Sub

Private Sub BtnEnh33_Click()
If Val(AdjustBal.Text) = 0 Then
   See_Msg "Nothing To Adjust", 5, ""
   Exit Sub
End If
NBal.Text = Format(Val(Scan_Data(6).Text) + Val(AdjustBal.Text) * -1, "0.00")
If Val(NBal.Text) < 0 Then
  Label29.Caption = "New Credit Balance"
Else
  Label29.Caption = "New Balance Owed"
End If
NBal.Text = Format(Abs(NBal.Text), "0.00")
AddCredit = True
End Sub
Private Sub BtnEnh34_Click()
If Val(AdjustBal.Text) = 0 Then
   See_Msg "Nothing To Adjust", 5, ""
   Exit Sub
End If
NBal.Text = Format(Val(Scan_Data(6).Text) + Val(AdjustBal.Text), "0.00")
If Val(NBal.Text) < 0 Then
  Label29.Caption = "New Credit Balance"
Else
  Label29.Caption = "New Balance Owed"
End If
NBal.Text = Format(Abs(NBal.Text), "0.00")
AddCredit = False
End Sub
Private Sub BtnEnh35_Click()
GetAccount
AccountbalanceEdit.Visible = False
AccountbalanceEdit.Left = -5000
End Sub
Private Sub BtnEnh36_Click()
If See_Msg("Are You Sure", 4, Label29.Caption & Format(NBal.Text, "0.00")) <> 6 Then Exit Sub
Check_no = -1111
If AddCredit Then
  EditGiftCards Scan_Data(3).Text, Val(AdjustBal.Text), "Mananger Account Adjust", AcctType.Text, "", Scan_Data(0).Text
Else
  EditGiftCards Scan_Data(3).Text, Val(AdjustBal.Text) * -1, "Mananger Account Adjust", AcctType.Text, "", Scan_Data(0).Text
End If
Check_no = 0
End Sub
Public Sub BtnEnh4_Click()
Set HouseAccounts = Nothing
SetMemberToNothing
Call MemoryCheck(Assign_Name)
End Sub
Private Sub BtnEnh6_Click()
AcctTypeSetup.Visible = False
AcctTypeSetup.Left = -5000
End Sub
Private Sub BtnEnh7_Click()
On Error GoTo ErrorHandler
If GPName.Text = "" Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If

 Sql = "Select * from [HouseGroups] Where [Group Name] = '" & GPName.Text & "'"
ConnectADORS Sql, Db, Tb
 If Tb.EOF Then Tb.AddNew
 Tb![Group Name] = GPName.Text
 Tb![Group Tender] = GPTender.Text
 Tb![Accept Tip] = TipIsOn.Text
 Tb.fields("Billing Amount").Value = Format(Val(BillAmt.Text), "0.00")
 Tb.Update
DbClose Db, Tb
SaveSetting "Micro$ale", "Register Setup", "Default Account Type", Default.Text
TrackHouseEvents LoggedOnAs(0).EMPName, "Change to Tender Link for House Accounts " & GPName.Text & " Group Tender " & GPTender.Text
See_Msg "Saved....", 5, ""
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Assign Name BTN7 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub BtnEnh8_Click()
If GPName.Text = "" Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
 Sql = "delete from [HouseGroups] Where [Group Name] = '" & GPName.Text & "'"
Db.Execute Sql
DbClose Db, Tb
SaveSetting "Micro$ale", "Register Setup", "Default Account Type", ""
See_Msg "Removed....", 5, ""
End Sub
Private Sub BtnEnh9_Click()
End Sub


Private Sub Check1_Click(Index As Integer)
If SPCINT(Index).Text = "None" Or SPCINT(Index).Text = "" Then Check1(Index).Value = 0: Exit Sub
If Check1(Index).Value = 1 Then
 If InStr(SPCINT(Index).Text, "*") = 0 Then
   SPCINT(Index).Text = "*" & SPCINT(Index).Text
 End If
Else
 If InStr(SPCINT(Index).Text, "*") > 0 Then
   SPCINT(Index).Text = Mid(SPCINT(Index).Text, 2, Len(SPCINT(Index).Text))
 End If
End If
End Sub
Private Sub Clocation_Click()
If LName.Text = "" Then Exit Sub
If Clocation.Value = 0 Then Exit Sub
'Dim Up_date As Integer
'For Up_date = 1 To Number_of_Registers - 1
'    WriteTxtFile MyRegisters(Up_date).RegLoc & "INI Files\" & "Location.INI", "[CONFIG]", OutPutFile
'    WriteTxtFile MyRegisters(Up_date).RegLoc & "INI Files\" & "Location.INI", "Site Name=" & LName.Text, AppendFile
'Next Up_date
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Combo3_Click()
Dim TipOn As String
Dim BillValue As Currency
Call GetGpTender(Combo3.Text, TipOn, BillValue)
ChgAmt.Text = Format(BillValue, "0.00")
End Sub
Private Sub Command1_Click()
CleanData
End Sub
Private Sub Command2_Click()
LSetup.Visible = False
LSetup.Left = -5000
End Sub
Private Sub Command3_Click()
On Error GoTo ErrorHandler
If LName.Text = "" Then Exit Sub
Dim RC As ADODB.Recordset
Dim Db As ADODB.Connection
If ConnectADODB(HouseAcct, Db) Then
   Sql = "Select * from [Locations] Where [Location Name] = '" & LName.Text & "'"
   ConnectADORS Sql, Db, RC
     If RC.EOF Then
       RC.AddNew
       LocationList.AddItem LName.Text
     End If
     RC![Location Name] = LName.Text
     RC.fields("Location Charge Privilege").Value = "False"
     RC.Update
     See_Msg "Saved...", 5, ""
End If
DbClose Db, RC
Exit Sub
ErrorHandler:

End Sub

Private Sub Command4_Click()
If LName.Text = "" Then Exit Sub
Dim RC As ADODB.Recordset
Dim Db As ADODB.Connection
If ConnectADODB(HouseAcct, Db) Then
   Sql = "delete from [Locations] Where [Location Name] = '" & LName.Text & "'"
   Db.Execute Sql
   See_Msg "Removed...", 5, ""
End If
DbClose Db, RC
End Sub
Private Sub Command5_Click()
Dim RC As ADODB.Recordset
Dim Db As ADODB.Connection
Dim PR As Integer
If ConnectADODB(HouseAcct, Db) Then
  Sql = "Delete from [Locations]"
  Db.Execute Sql
  Sql = "Select * from [Locations] "
 ConnectADORS Sql, Db, RC
   
   For PR = 0 To LocationList.ListCount - 1
     If LocationList.List(PR) > "" Then
      RC.AddNew
      RC![Location Name] = LocationList.List(PR)
      RC![Location Charge Privilege] = LocationList.Selected(PR)
      RC![Location Number] = Encrypt(LocationList.List(PR), "PW")
      RC.Update
     End If
   Next PR
End If
DbClose Db, RC
See_Msg "Saved...", 5, ""
End Sub
Private Sub CreditBalance_Click()
If CreditBalance.Value = 0 Then
 If Can_Continue("Pre-Paid Balance Only") = False Then
  CreditBalance.Value = 1
  Exit Sub
 End If
End If
End Sub
Private Sub CRGEMAI_Click()
If Can_Continue(CRGEMAI.Caption) = False Then Exit Sub

Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If
 Sql = "Select * from [HouseGroups]"
ConnectADORS Sql, Db, Tb
 Combo3.Clear
 Combo3.AddItem "ALL"
 Combo3.AddItem "House Account"
 Combo3.AddItem "Gift Card"
 Combo3.AddItem "Frequent Diner"
 Combo3.AddItem "Membership"
 If Not Tb.EOF Then
   Do While Not Tb.EOF
     If IsCombo(AcctType, Tb![Group Name] & "") = False Then
       Combo3.AddItem Tb![Group Name] & ""
     End If
    Tb.MoveNext
   Loop
 End If
DbClose Db, Tb
CenterObject Me, MemberPage
With MemberPage
   .ZOrder 0
   .Visible = True
End With

End Sub
Private Sub CXPT_Click()
On Error GoTo ERH
If Dir(SvrPath.TerminalPath & "AccountExport", vbDirectory) = "" Then
 MkDir SvrPath.TerminalPath & "AccountExport"
End If
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim ff              As Integer
Dim AccountExposure As Currency
Dim Cr              As Long
Dim RTotal          As Long
ff = FreeFile
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
If ConnectADORS("Select * from [HouseAccountData]  order by [Scan Name] asc", Db, RS) = False Then Exit Sub
Open SvrPath.TerminalPath & "AccountExport\HouseAccount.TXT" For Output As ff
Print #ff, "Account Name" & "," & "Account Number" & "," & "Account Type" & "," & "Balance"
Cr = 0
RTotal = RecordsCount(RS)
Do While Not RS.EOF
   Cr = Cr + 1
   NoAcct.Caption = "Calculating " & RS![scan Name] & "" & " Record " & CStr(Cr) & " Of " & CStr(RTotal)
   NoAcct.Refresh
   AccountExposure = ReturnAllAccountCharges(RS![scan Name])
   Print #ff, RS![scan Name] & "," & RS![Acct Number] & "," & RS![Account Type] & "," & Format(AccountExposure, "0.00")
  RS.MoveNext
Loop
DbClose Db, RS
Close ff
See_Msg "File exported to", 5, SvrPath.TerminalPath & "AccountExport\HouseAccount.TXT"
Exit Sub
ERH:
Err.Clear

End Sub

Private Sub delete_Click()
On Error GoTo ErrorHandler
If Abs(delete.Value) = 1 Then
If Scan_Data(0).Text = "" Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
Sql = "select * from [HouseCharges] where [Scan Number] = '" & Scan_Data(3).Text & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub

If Tb.EOF Then
    If See_Msg("No Records Found", 4, "Delete Records Now ?") = 6 Then
        SetNothing Tb
        'Set Tb = Db.OpenRecordset("select * from [HouseAccountData] where [Scan Name] = """ & Scan_Data(0).Text & """")
        Sql = "select * from [HouseAccountData] where [Scan Name] = '" & Scan_Data(0).Text & "'"
        If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
     If Not Tb.EOF Then
        Tb.delete
        DbClose Db, Tb
        Dim W As Integer
        For W = 0 To Scan_Data(0).ListCount - 1
            If Trim(Scan_Data(0).List(W)) = Trim(Scan_Data(0).Text) Then
                Scan_Data(0).RemoveItem W
                Exit For
            End If
        Next W
        Clear_data
    End If
End If
See_Msg "Account has open charges will remove on closeout !", 5, ""
End If
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Delete Click Assign Names ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Public Sub UserPrvilage(Off As Boolean)
Select Case Off
   Case Is = True
     PAS.Enabled = False   'print Account Statements
     SSCommand4.Enabled = False   'delete acccount
     SSCommand10.Enabled = False  'reset account activity
     CXPT.Enabled = False  'Create Export file
     CRGEMAI.Enabled = False  'report / Util
     EMSTM.Enabled = False      'email statements
     BtnEnh3.Enabled = False      'print account act
     AttachDiscount.Enabled = False
     CreditBalance.Value = 1
     delete.Enabled = False
End Select
End Sub
Public Sub AddToExpiration(aExpDate As String)
If IsDate(aExpDate) Then
 If IsDate(AExpiration.Text) Then
   AExpiration.Text = CDate(AExpiration.Text) + DateDiff("d", Now, CDate(aExpDate))
 Else
   AExpiration.Text = Format(aExpDate, "MM/DD/YYYY")
 End If
Else
  If Not IsDate(AExpiration.Text) Then
   AExpiration.Text = "Never"
  End If
End If
End Sub
Private Sub LoadAccounts()
On Error GoTo ErrorHandler
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim Tb As ADODB.Recordset
Dim LocationId As Long
LocationId = GetLocationId
   If LocationId <> 0 And LoggedOnAs(0).EmpAccessLevel < AdminLevel Then
    Sql = "Select [Scan Name],[Acct Number] from [HouseAccountData] Where [Meal Plan] = " & LocationId & " order by [Account Type] desc,[Scan Name] asc"
   Else
    Sql = "Select [Scan Name],[Acct Number] from [HouseAccountData] order by [Account Type] desc,[Scan Name] asc"
   End If
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
    If Not Tb Is Nothing Then
      If Not Tb.BOF Then Tb.MoveFirst
    End If
    Scan_Data(0).Clear
    TargetCard.Clear
    SourceCard.Clear
    Scan_Data(3).Clear
    NoAcct.Caption = "Total Accounts " & RecordsCount(Tb)
    Do While Not Tb.EOF
     If Not IsNull(Tb![scan Name]) Then
      Scan_Data(0).AddItem Tb![scan Name] & ""
      TargetCard.AddItem Tb![Acct Number] & ""
      SourceCard.AddItem Tb![Acct Number] & ""
     End If
     Tb.MoveNext
    Loop
    SetNothing Tb
    Sql = "Select * from [HouseGroups]"
    Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Groups")
    If Not Tb Is Nothing Then
      If Not Tb.BOF Then Tb.MoveFirst
    End If
    If Not Tb.EOF Then
        Do While Not Tb.EOF
          AcctType.AddItem Tb![Group Name] & ""
          Tb.MoveNext
        Loop
    End If
    DbClose Nothing, Tb
    Set GRS = Nothing
    Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "LoadAccounts ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
    
End Sub

Private Sub EAB_Click()
If Can_Continue(EAB.Caption) = False Then Exit Sub
If Scan_Data(0).Text = "" Then
  See_Msg "Select Account", 5, ""
  Exit Sub
End If
Dim MyAccess As String
MyAccess = GetConfigInfo("Functions", "Edit Account Data", "", "FUNCTIONS.INI")
If MyAccess = "ENABLED" Then
  IAmAdmin = True
Else
  IAmAdmin = False
End If
If IAmAdmin = False Then
 If AdminIsOn = False Then
   See_Msg "Administrator Account Required for Edit", 5, "Admin Account Required."
   Exit Sub
 End If
 If AdminIsOn Then
  If GetAdministrator("Edit House Account") = False Then Exit Sub
  IAmAdmin = True
 End If
End If
AdjustBal.Text = "0.00"
If Val(Scan_Data(6).Text) < 0 Then
  Label28.Caption = "Credit Balance"
Else
  Label28.Caption = "Balance Owed"
End If
Label29.Caption = "New Balance"
TrackHouseEvents LoggedOnAs(0).EMPName, "Account" & "Edit House Account"
CenterObject Me, AccountbalanceEdit
AccountbalanceEdit.Visible = True
AccountbalanceEdit.ZOrder 0

End Sub

Private Sub EMSTM_Click()
If CloseWeek Then GoTo DoData
If Can_Continue(EMSTM.Caption) = False Then Exit Sub
If See_Msg("Email All Statements", 4, "Send ?") <> 6 Then Exit Sub
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
End If
Dim LT As Integer
DoData:
If CDate(SearchDates.Item(1).StartDate) > CDate(StartDate.Value) Then
  See_Msg "Invalid Start Date.", 0, "Current Period Start is " & Format(StartDate.Value, "MM/DD/YYYY")
  Exit Sub
End If
See_Msg "Sending Statements.....", 5, ""
ReDim FailedPayments(Scan_Data(0).ListCount)
For LT = 0 To Scan_Data(0).ListCount - 1
  PrintStatments Combo3.Text, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, Scan_Data(0).List(LT), LT, "EMail Account Activity"
Next LT
End Sub

Private Sub Ext_Click()
Set HouseAccounts = Nothing
SetMemberToNothing
Call MemoryCheck(Assign_Name)

End Sub

Private Sub Form_Load()
On Error Resume Next
IAmAdmin = False
PrintingStatements = True
HouseAccountHistory = "None"
'Me.Top = 0
'Me.Left = 0
ResizeScreen1024 Me
Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 0
Check1(3).Value = 0
Dim RSDate As String
Dim REDate As String
GetAccountDateRanges RSDate, REDate
StartDate.Value = Format(RSDate, "MM/DD/YYYY")
EndDate.Value = Format(REDate, "MM/DD/YYYY")

If Check_System_Access("Admin") = 1 Then
   AdminLevel = 9
Else
   AdminLevel = Check_System_Access("Admin")
End If
CreditBalance.Value = 1
SQLConnected.Caption = "Connected To Sql"


Dim Rdb As ADODB.Connection, RTB As ADODB.Recordset
Dim i As Integer
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If LoggedOnAs(0).EmpAccessLevel <= Check_System_Access("Mask") Then
  SPCINT(0).PasswordChar = "X"
  SPCINT(1).PasswordChar = "X"
  SPCINT(2).PasswordChar = "X"
  SPCINT(3).PasswordChar = "X"
  AttachDiscount.Enabled = False
End If
If LoggedOnAs(0).EmpAccessLevel <= Check_System_Access("Mask1") Then
  SPCINT(0).PasswordChar = "X"
  AttachDiscount.Enabled = False
End If
If LoggedOnAs(0).EmpAccessLevel <= Check_System_Access("Mask2") Then
  SPCINT(1).PasswordChar = "X"
  AttachDiscount.Enabled = False
End If
If LoggedOnAs(0).EmpAccessLevel <= Check_System_Access("Mask3") Then
  SPCINT(2).PasswordChar = "X"
  AttachDiscount.Enabled = False
End If
If LoggedOnAs(0).EmpAccessLevel <= Check_System_Access("Mask4") Then
  SPCINT(3).PasswordChar = "X"
  AttachDiscount.Enabled = False
End If
AcctType.Text = "ALL"
If GetSetting("Micro$ale", "Register Setup", "Default Account Type") > "" Then
  AcctType.Text = GetSetting("Micro$ale", "Register Setup", "Default Account Type")
Else
  AcctType.Text = "ALL"
End If
If OkToDo = False Then
  LoadOpts.Visible = True
  LoadAccounts
  If GetSetting("Micro$ale", "Register Setup", "Default Account Type") > "" Then
    SSCommand14_Click
  End If
End If
AttachDiscount.AddItem "None"
If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
If ConnectADORS("Select * from [Discounts] order by [Discount Name] asc", Db, Tb) = False Then Exit Sub
Do While Not Tb.EOF
  AttachDiscount.AddItem Tb![Discount Name] & ""
  Tb.MoveNext
Loop
DbClose Db, Tb
If Not HouseAccounts Is Nothing Then
  If HouseAccounts.Item(1) Is Nothing Then Exit Sub
    Scan_Data(0).Text = HouseAccounts.Item(1).AccountName
    GetAccount
    SSCommand10.Enabled = False
    SSCommand4.Enabled = False
End If
'Dim BkUpFile As String
'AccountBackUp.Clear
'AccountBackUp.AddItem "None"
'AccountBackUp.Text = "None"
'BkUpFile = Dir(SvrPath.TerminalPath & "*House Accounts.mdb")
'If BkUpFile = vbNullString Then Exit Sub
' Do Until BkUpFile = vbNullString
'   AccountBackUp.AddItem BkUpFile
'   BkUpFile = Dir
' Loop
Exit Sub
ErrorHandler:
See_Msg Err.Description, 0, "Error Loading"
End Sub
Private Sub HouseAccountBackUp(BeginDate As String)
On Error GoTo ErrorHandler
Dim Cdb As ADODB.Connection
Dim Hdb As ADODB.Connection
Dim CRS As ADODB.Recordset
Dim HRS As ADODB.Recordset
Dim Er  As Long  'er each record
Dim Destination As String
If Dir(SvrPath.TerminalPath & "HouseAccounts.Mdb") = "" Then
   See_Msg "No Account file found", 5, ""
 Exit Sub
End If
See_Msg "Saving Historical Data", 5, ""
Destination = Format(BeginDate, "MM-DD-YYYY") & " Week " & "HouseAcctBackUp.MDB"
FileCopy SvrPath.TerminalPath & "HouseAccounts.mdb", SvrPath.TerminalPath & Destination
SqlToAccessBackUp SvrPath.TerminalPath & Destination, Destination, SvrPath.TerminalPath & HouseAcct, HouseAcct
See_Msg "Account History Update Complete", 5, ""
Exit Sub
ErrorHandler:
Err.Clear
End Sub
Private Sub GetAccountDateRanges(BegingRange As String, EndingRange As String)
On Error GoTo ERH
    See_Msg "Loading Date Ranges...", 5, ""
    Dim RC As ADODB.Recordset
    Dim GRS As GeneralLib.Utilities
    Set GRS = New GeneralLib.Utilities
    Sql = "Select Top 1 [Check Date] From [HouseCharges] "
    Sql = Sql & "Order by [Check Date] asc "
    Set RC = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Charge Data")
    If RC Is Nothing Then
       See_Msg "Nothing to load", 5, ""
       Exit Sub
    End If
    If Not RC Is Nothing Then
      If Not RC.BOF Then RC.MoveFirst
    End If
    If Not RC.EOF Then
      BegingRange = CStr(Format(RC![Check Date], "mm/dd/yyyy"))
    Else
      BegingRange = Format(Date, "MM/DD/YYYY")
    End If
    EndingRange = Format(Date, "MM/DD/YYYY")
DbClose Nothing, RC
Set GRS = Nothing
Exit Sub
ERH:
BegingRange = Format(Date, "MM/DD/YYYY")
EndingRange = Format(Date, "MM/DD/YYYY")

End Sub
Private Sub Form_Unload(Cancel As Integer)
  PrintingStatements = False
  HouseAccountHistory = "None"
End Sub

Private Sub GAN_Click()
If Can_Continue(GAN.Caption) = False Then Exit Sub

If Scan_Data(0).Text = "" Then
   See_Msg "Name Required", 5, "Cannot Create Account Number"
   Exit Sub
End If
If Address(0).Text = "" Then
   See_Msg "Address Required", 5, "Cannot Create Account Number"
   Exit Sub
End If
Scan_Data(3).Text = GenerateAccount(Scan_Data(0).Text)
Scan_Data(3).Text = Scan_Data(3).Text & GenerateAccount(Address(0).Text)

End Sub

Private Sub GPName_Click()
Dim TipOn As String
Dim BillValue As Currency
GPTender.Text = GetGpTender(GPName.Text, TipOn, BillValue)
TipIsOn.Text = TipOn
BillAmt.Text = Format(BillValue, "0.00")
End Sub
Private Sub ImportData_Click()
On Error GoTo ErrorHandler
Dim Db              As ADODB.Connection
Dim RS              As ADODB.Recordset
Dim PS              As ADODB.Recordset
Dim P1 As String
Dim P2 As String
Dim P3 As String
Dim P4 As String
Dim P5 As String
Dim P6 As String
Dim P7 As String
Dim P8 As String
Dim P9 As String
Dim P10 As String
Dim P11 As String
Dim P12 As String
Dim P13 As String
Dim P14 As String
Dim P15 As String
Dim P16 As String
Dim P17 As String
Dim P18 As String
Dim P19 As String
Dim P20 As String
Dim P21 As String
Dim P22 As String
Dim P23 As String
Dim P24 As String
Dim P25 As String
Dim P26 As String
Dim P27 As String
Dim P28 As String
Dim P29 As String
Dim ff  As Integer
If ConnectADODB(HouseAcct, Db) = False Then Exit Sub
  Sql = "Select * from [HouseAccountData] "  'Change to HouseAccountData
ConnectADORS Sql, Db, RS
  Sql = "Select * from [HousePayments] "
ConnectADORS Sql, Db, PS
ff = FreeFile
Open SvrPath.TerminalPath & "AccountBalances.csv" For Input As #ff
Do
  Input #ff, P1, P2, P3, P4 ', P5, P6, P7, P8, P9, P10, P11, P12   ', P13, P14 ', P15, P16, P17, P18, P19, P20, P21, P22, P23, P24, P25, P26, P27, P28, P29
  RS.AddNew
  RS![Memo] = ""
  RS![Account Type] = "Gift Card"
  RS![scan Name] = P1
  RS![Max Daily Amount] = 0
  RS![Max Weekly Amount] = 0
  RS![Daily Charge] = 0
  RS![Phone] = ""
  RS![Credit Balance] = 0
  RS![Acct Number] = P1
  RS![Address] = ""
  RS![Suspend] = 0
  RS![City] = ""
  RS![Zip] = ""
  RS![State] = ""
  RS![Delete Record] = 0
  RS.Update
  
  PS.AddNew
  PS![Date] = Format(Date, "MM/DD/YYYY")
  PS![Account Name] = P1
  PS![Amount] = Val(P4) ', "0.00")
  PS![Payment Type] = "Import"
  PS![Location Name] = ""
  PS![Check Number] = 100
  PS![Account Number] = P1
  Loop Until EOF(ff)
  Close #ff
  PS.Close
  DbClose Db, RS
  MsgBox "Import Complete", vbInformation, "File Imported."
  Exit Sub
ErrorHandler:
See_Msg "Import Failed..", 0, Err.Description
End Sub
Private Sub Label1_Click(Index As Integer)
'If Can_Continue(Label1.Item(Index).Caption) = False Then Exit Sub
Dim Msg         As String
Dim DataRetuned As String
Select Case Index
   Case Is = 1
     Msg = "Name"
   Case Is = 4
     Msg = "Account Number"
   Case Is = 7
     Msg = "Address"
   Case Is = 8
     Msg = "City"
   Case Is = 9
     Msg = "State"
   Case Is = 10
     Msg = "Zip"
   Case Is = 22
     Msg = "Account Number"
   Case Is = 21
     Msg = "Phone"
   Case Is = 11
     Msg = "Cell Number"
   Case Is = 13
     Msg = "EMail"
   Case Is = 17
     Msg = "Fax"
   Case Is = 14
     BDay.Text = "None"
     Exit Sub
   Case Is = 15
     AnvDate.Text = "None"
     Exit Sub
   Case Is = 18
     AExpiration.Text = "None"
     Exit Sub
   Case Else
     Exit Sub
End Select
DataRetuned = ShowKeyBoard(Msg)
If DataRetuned = "" Then Exit Sub
If Index = 1 Then
  Scan_Data(0).Text = DataRetuned
End If
If Index = 4 Then
  Scan_Data(3).Text = DataRetuned
End If
If Index = 7 Then
  Address(0).Text = DataRetuned
End If
If Index = 8 Then
  Address(1).Text = DataRetuned
End If
If Index = 9 Then
  Address(2).Text = DataRetuned
End If
If Index = 10 Then
  Address(3).Text = DataRetuned
End If
If Index = 21 Then
  Address(4).Text = DataRetuned
End If
If Index = 11 Then
  Address(5).Text = DataRetuned
End If
If Index = 13 Then
  Address(6).Text = DataRetuned
End If
If Index = 17 Then
  Address(7).Text = DataRetuned
End If
End Sub

Private Sub Label27_Click()
AdjustBal.Text = ShowNumberPad("Enter Amount", False, 0, 0)
End Sub

Private Sub LName_Click()
If LName.Text = GetSetting("Micro$ale", "Register Setup", "Site") Then
  Clocation.Value = 1
Else
  Clocation.Value = 0
End If

End Sub

Private Sub LNAN_Click()
tryagain:
Dim Heading As String
Dim vDb As ADODB.Connection
Dim vRc As ADODB.Recordset
Dim VRs As ADODB.Recordset
If ConnectADODB(HouseAcct, vDb) = False Then
   Exit Sub
End If

  Sql = "Select * from [HouseAccountData] order by [Scan Name] asc"
If ConnectADORS(Sql, vDb, vRc) = False Then Exit Sub
If vRc.EOF Then
  Call See_Msg("Nothing to Print", 5, "")
  SetNothing vRc
  Exit Sub
End If
Heading = "All Accounts List"
Dim TCharge As Currency, TTip As Currency, TVoid As Currency
   'SetPrinter 0
   SetPage ReportView.View, Heading
   With ReportView.View
    .FontBold = False
    .FontSize = 10
    .TextAlign = taLeftBottom
   End With
Do While Not vRc.EOF
        With ReportView.View
         DrawLine
                .Text = vRc![scan Name] & "": .CurrentX = 4500: .Text = "Account Number " & vRc![Acct Number] & "" & vbCrLf
                If Trim(vRc![Address]) > "" Then
                .Text = vRc![Address] & " " & vRc![City] & " " & vRc![State] & " " & vRc![Zip] & vbCrLf
                End If
                If vRc![Phone] > "" Then
                .Text = "Phone " & vRc![Phone] & "    " & "Email " & vRc![Email] & "" & vbCrLf
                End If
                If vRc![Account Expires On] & "" > "" And vRc![Account Expires On] & "" <> "None" Then
                 .Text = "Account Expires On " & vRc![Account Expires On] & "" & vbCrLf
                End If
            DrawLine
          .TextAlign = taLeftBottom
        End With
    
nextR:
vRc.MoveNext
Loop
DbClose vDb, vRc
ReportView.View.EndDoc
ReportView.Show 1
DbClose vDb, vRc
End Sub

Private Sub LSA_Click()
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim LocationId As Long
LocationId = GetLocationId
If LocationId <> 0 And LoggedOnAs(0).EmpAccessLevel < AdminLevel Then
  Sql = "Select [Scan Name] from [HouseAccountData] Where [Meal Plan] = " & LocationId & " and [Suspend] <> 0 order by [Account Type] desc,[Scan Name] asc"
Else
  Sql = "Select [Scan Name] from [HouseAccountData] Where [Suspend] <> 0 order by [Account Type] desc,[Scan Name] asc"
End If
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
If Not Tb Is Nothing Then
  If Not Tb.BOF Then Tb.MoveFirst
End If
Scan_Data(0).Clear
NoAcct.Caption = "Total Accounts " & RecordsCount(Tb)
Do While Not Tb.EOF
  Scan_Data(0).AddItem Trim$(Tb![scan Name] & "")
  Tb.MoveNext
Loop
DbClose Nothing, Tb
Set GRS = Nothing
See_Msg "Accounts Loaded", 5, ""
End Sub

Private Sub LTTA_Click()
If Can_Continue(LTTA.Caption) = False Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
GPName.Clear
Default.Clear
GPName.AddItem "House Account"
GPName.AddItem "Gift Card"
GPName.AddItem "Frequent Diner"
GPName.AddItem "Membership"
Default.AddItem "House Account"
Default.AddItem "Gift Card"
Default.AddItem "Frequent Diner"
Default.AddItem "Membership"
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If
 Sql = "Select * from [HouseGroups]"
ConnectADORS Sql, Db, Tb
 If Not Tb.EOF Then
   Do While Not Tb.EOF
    GPName.AddItem Tb![Group Name] & ""
    Default.AddItem Tb![Group Name] & ""
    Tb.MoveNext
   Loop
 End If
DbClose Db, Tb
Dim Dbt As ADODB.Connection, TBT As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Dbt
OpenTb "Select * from [Credit Cards]", TBT, Dbt
GPTender.Clear
GPTender.AddItem "None"
Do While Not TBT.EOF
 GPTender.AddItem Check_Null(TBT(0))
 TBT.MoveNext
Loop
DbClose Dbt, TBT
Default.Text = GetSetting("Micro$ale", "Register Setup", "Default Account Type")
CenterObject Me, AcctTypeSetup
AcctTypeSetup.Visible = True
AcctTypeSetup.ZOrder 0
End Sub
Private Function GetGpTender(DataName As String, AllowTip As String, BillingAmt As Currency) As String
Dim Tb As ADODB.Recordset
AllowTip = "Yes"
If UCase(DataName) = "CASH" Or UCase(DataName) = "CREDIT CARD" Or InStr(UCase(DataName), "CHECK") > 0 Then Exit Function
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
 Sql = "Select * from [HouseGroups] Where [Group Name] = '" & DataName & "'"
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "Groups")
If Not Tb Is Nothing Then
  If Not Tb.BOF Then Tb.MoveFirst
End If
If Not Tb.EOF Then
  GetGpTender = Tb![Group Tender] & ""
  AllowTip = Tb![Accept Tip] & ""
  BillingAmt = Val(Tb.fields("Billing Amount").Value & "")
Else
  GetGpTender = ""
  AllowTip = "Yes"
  BillingAmt = 0
End If
DbClose Nothing, Tb
Set GRS = Nothing
End Function

Private Sub NewAcct_Click()
Dim DataIn As String
DataIn = ShowKeyBoard("Enter New Account Name")
If DataIn > "" Then
 NewData(0).Text = DataIn
End If
End Sub
Private Sub NewAcctNumber_Click()
Dim Results As String
Results = ShowKeyBoard("Enter New Account Number")
If Results > "" Then
 NewData(1).Text = Results
End If
End Sub
Private Sub NewData_DblClick(Index As Integer)
If NewData(Index).Text = "" Then NewData(Index).Text = OldData(Index).Caption
End Sub
Private Sub NewData_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
  KeyAscii = ReturnSentinal(KeyAscii)
End If
End Sub
Private Sub QBChg_Error(ErrorCode As Integer, Description As String)

End Sub

Private Sub Quit_Click()
End Sub

Private Sub PAS_Click()
If AcctType.Text = "" Then
  See_Msg "Must Select Account Type", 5, ""
  Exit Sub
End If
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
   Exit Sub
Else
  If HouseAccountHistory = "None" Then
   If CDate(SearchDates.Item(1).StartDate) > CDate(StartDate.Value) Then
    See_Msg "Invalid Start Date.", 0, "Current Period Start is " & Format(StartDate.Value, "MM/DD/YYYY")
    Exit Sub
   End If
  End If
 If Check_Option(96) = 0 Then
    HouseAccount AcctType.Text, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate
 Else
    ReDim FailedPayments(1)
    PrintStatments AcctType.Text, SearchDates.Item(1).StartDate, SearchDates.Item(1).EndDate, "", 0, ""
 End If
  StatementDate = Format(SearchDates.Item(1).StartDate, "MM-DD-YYYY")
  Set SearchDates = Nothing
End If

End Sub

Private Sub PAVS_Click()
EventDates "Anniversary", "Anniversary"

End Sub

Private Sub PBD_Click()
EventDates "Birthday", "Birthday"
End Sub

Private Sub PEACT_Click()
GETEXPDATE "Expired Accounts", True, False, "", ""
End Sub

Private Sub PED_Click()
EventDates "Expiration", "Account Expires On"

End Sub

Private Sub PZBGC_Click()
If Can_Continue(PZBGC.Caption) = False Then Exit Sub
GETEXPDATE "0 Balance Gift Cards", False, True, "", ""

End Sub

Private Sub RICC_Click()

End Sub

Private Sub Scan_Data_Click(Index As Integer)
If AccountBackUp.Text = "" Then AccountBackUp.Text = "None"
If AccountBackUp.Text <> "None" Then
  See_Msg "Print Statements Only", 5, ""
  Exit Sub
End If
If Index = 0 Then

Else
  Exit Sub
End If
Set HouseAccounts = Nothing
Set HouseAccounts = New AllAccounts
TenderLink = ""
'If GetAccountData(Scan_Data(0).Text) = False Then
'  Clear_data
'  Exit Sub
'End If
On Error GoTo ErrorHandler
GetAccount
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Scan Data Click ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub Scan_Data_DblClick(Index As Integer)
If Index = 3 Then
  GetAccountData Scan_Data(3).Text
End If
End Sub
Private Sub Scan_Data_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   Scan_Data(Index).Text = GetValidAcctNo(Scan_Data(Index).Text)
End If
KeyAscii = ReturnSentinal(KeyAscii)
Select Case KeyAscii
    Case Is = 39
      KeyAscii = 96
    Case Is = 33, 34, 44
      KeyAscii = 0
      Exit Sub
End Select
If Index = 4 Then KeyAscii = 0
If Index = 1 And Len(Scan_Data(0).Text) > 40 And KeyAscii <> 8 Then
KeyAscii = 0
End If
If Index = 3 And Len(Scan_Data(3).Text) > 50 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
Select Case Index
   Case Is = 4, 5, 6
     KeyAscii = 0
End Select
End Sub
Private Sub Scan_Data_Validate(Index As Integer, Cancel As Boolean)
If Index = 3 Then
If Scan_Data(Index).Text = "None" Or Scan_Data(Index).Text = "" Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If

Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & Scan_Data(3).Text & "'"
Sql = Sql & " or [Spc Interest 1] = '" & Scan_Data(3).Text & "'"
Sql = Sql & " or [Spc Interest 2] = '" & Scan_Data(3).Text & "'"
Sql = Sql & " or [Spc Interest 3] = '" & Scan_Data(3).Text & "'"
Sql = Sql & " or [Spc Interest 4] = '" & Scan_Data(3).Text & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
If Scan_Data(0).Text > "" And Tb![scan Name] & "" <> Scan_Data(0).Text Then
        'See_Msg Tb![Scan Name], 0, Tb![Acct Number]
        See_Msg "Duplicate Account Number", 5, Tb![scan Name]
        Scan_Data(3).Text = ""
        DbClose Db, Tb
        Exit Sub
End If
End If
DbClose Db, Tb
End If
End Sub
Private Sub SourceCard_Click()
Dim Db As ADODB.Connection
Dim RC As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If
Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & SourceCard.Text & "'"
If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
If RC.EOF Then
  See_Msg "Account Number Not Found", 5, ""
  SourceCard.Text = ""
  SourceName.Caption = ""
  DbClose Db, RC
  Exit Sub
Else
  SourceName.Caption = RC![scan Name] & ""
End If
DbClose Db, RC

End Sub

Private Sub SourceCard_KeyPress(KeyAscii As Integer)
KeyAscii = ReturnSentinal(KeyAscii)
Select Case KeyAscii
    Case Is = 39
      KeyAscii = 96
    Case Is = 33, 34, 44
      KeyAscii = 0
      Exit Sub
End Select
End Sub

Private Sub SPCINT_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = ReturnSentinal(KeyAscii)
Select Case KeyAscii
    Case Is = 39
      KeyAscii = 96
    Case Is = 33, 34, 44
      KeyAscii = 0
      Exit Sub
End Select
End Sub

Private Sub SPCINT_Validate(Index As Integer, Cancel As Boolean)
If SPCINT(Index).Text = "None" Or SPCINT(Index).Text = "" Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
  Exit Sub
End If

Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & SPCINT(Index).Text & "'"
Sql = Sql & " or [Spc Interest 1] = '" & SPCINT(Index).Text & "'"
Sql = Sql & " or [Spc Interest 2] = '" & SPCINT(Index).Text & "'"
Sql = Sql & " or [Spc Interest 3] = '" & SPCINT(Index).Text & "'"
Sql = Sql & " or [Spc Interest 4] = '" & SPCINT(Index).Text & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
If Tb![scan Name] & "" <> Scan_Data(0).Text Then
        'See_Msg Tb![Scan Name], 0, Tb![Acct Number]
        See_Msg "Duplicate Account Number", 5, Tb![scan Name]
        SPCINT(Index).Text = ""
        DbClose Db, Tb
        Exit Sub
End If
End If
DbClose Db, Tb
End Sub
Public Sub SSCommand10_Click()
If CloseWeek Then GoTo DoData
If Can_Continue(SSCommand10.Caption) = False Then Exit Sub
If See_Msg("Have you printed your reports ?", 4, "Reports will not be printed for this function") <> 6 Then Exit Sub
If UCase(AcctType.Text) <> "ALL" Then
   If See_Msg("Reset Only " & AcctType.Text, 4, "OK To Continue ?") <> 6 Then Exit Sub
Else
   If See_Msg("Reset ALL Accounts ", 4, "OK To Continue ?") <> 6 Then Exit Sub
End If
DatePicker.StartDate.Caption = "Select a Ending Date"
DatePicker.StartingDate.Text = Format(Date, "MM/DD/YYYY")
DatePicker.EndingDate.Visible = False
DatePicker.EndDate.Visible = False
DatePicker.Show 1
If SearchDates Is Nothing Then
   Exit Sub
End If
If Not IsDate(SearchDates.Item(1).StartDate) Then Exit Sub
If See_Msg("Reset Account Activity ?", 4, "Thru " & SearchDates.Item(1).StartDate) <> 6 Then Exit Sub
DoData:
   ' If CDate(SearchDates.Item(1).StartDate) > CDate(StartDate.Value) Then
   '   See_Msg "Invalid Start Date.", 0, "Current Period Start is " & Format(StartDate.Value, "MM/DD/YYYY")
   '   Exit Sub
   ' End If
    HouseAccountBackUp SearchDates.Item(1).StartDate
    CalculateBalance SearchDates.Item(1).StartDate, AcctType.Text
    Dim vDb As ADODB.Connection
    If ConnectADODB(HouseAcct, vDb) = False Then
     Exit Sub
    End If

    Sql = "Delete from [HouseCharges]"
    If UCase(AcctType.Text) <> "ALL" Then
      Sql = Sql & " Where [Account Type] = '" & AcctType.Text & "'"
        Sql = Sql & " and TRY_CAST([Check date] AS datetime) Between '01/01/2000' and '" & CDate(SearchDates.Item(1).StartDate) & "'"
    Else
        Sql = Sql & " Where TRY_CAST([Check date] AS datetime) Between '01/01/2000' and '" & CDate(SearchDates.Item(1).StartDate) & "'"
    End If
    vDb.Execute Sql
    '**********************
     If Check_Option(227) = 1 Then
      Sql = "Delete from HouseChargeDetail"
      Sql = Sql & " Where [Account Type] = '" & AcctType.Text & "'"
      Sql = Sql & " and TRY_CAST([Check date] AS datetime) Between '01/01/2000' and '" & CDate(SearchDates.Item(1).StartDate) & "'"
      vDb.Execute Sql
      
     End If
    '**********************
    ResetPayments SearchDates.Item(1).StartDate, AcctType.Text
    SetNothing vDb
    See_Msg "Completed....", 5, "File Purge Completed."
End Sub
Private Sub SSCommand14_Click()
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
Dim LocationId As Long
LocationId = GetLocationId
If LocationId <> 0 And LoggedOnAs(0).EmpAccessLevel < AdminLevel Then
 Sql = "Select [Scan Name] from [HouseAccountData] Where [Meal Plan] = " & LocationId & " and [Account Type] = '" & AcctType.Text & "' order by [Account Type] desc,[Scan Name] asc"
Else
 Sql = "Select [Scan Name] from [HouseAccountData] Where [Account Type] = '" & AcctType.Text & "' order by [Account Type] desc,[Scan Name] asc"
End If
If ConnectADODB(HouseAcct, Db) = False Then Exit Sub
ConnectADORS Sql, Db, Tb

'Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
If Not Tb Is Nothing Then
  If Not Tb.BOF Then Tb.MoveFirst
End If
Scan_Data(0).Clear
NoAcct.Caption = AcctType.Text & " " & RecordsCount(Tb)
Do While Not Tb.EOF
  Scan_Data(0).AddItem Trim(Check_Null(Tb![scan Name]))
  Tb.MoveNext
Loop
DbClose Db, Tb
Set GRS = Nothing
End Sub

Private Sub SSCommand15_Click()
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
'LoadAccounts
Dim LocationId As Long
LocationId = GetLocationId
'Sql = "Select [Scan Name] from [HouseAccountData] order by [Account Type] desc,[Scan Name] asc"
Dim GRS As GeneralLib.Utilities
Set GRS = New GeneralLib.Utilities
If LocationId <> 0 And LoggedOnAs(0).EmpAccessLevel < AdminLevel Then
  Sql = "Select [Scan Name] from [HouseAccountData] where [Meal Plan] = " & LocationId & " order by [Account Type] desc,[Scan Name] asc"
Else
 Sql = "Select [Scan Name] from [HouseAccountData] order by [Account Type] desc,[Scan Name] asc"
End If
Set Tb = GRS.ReturnRecordset(Sql, ReturnAdoConnection(HouseAcct), "HouseAccountData")
If Not Tb Is Nothing Then
  If Not Tb.BOF Then Tb.MoveFirst
End If
Scan_Data(0).Clear
NoAcct.Caption = "Total Accounts " & RecordsCount(Tb)
Do While Not Tb.EOF
  Scan_Data(0).AddItem Trim(Check_Null(Tb![scan Name]))
  Tb.MoveNext
Loop
DbClose Db, Tb
Set GRS = Nothing

End Sub
Private Sub SSCommand16_Click()
End Sub
Private Sub EventDates(EventType As String, FieldName As String)
LoadDates "Select Distinct [Date] from [Sales]", SvrPath.TerminalPath & ClosingHistory, "Date"
If SearchDates Is Nothing Then
  Exit Sub
End If

Dim PrinterPort As String
Printer_Drivers "Receipt"
CenterText = EC(8)
LeftText = EC(9)

Dim PF As Integer
PF = FreeFile
PrinterPort = "Assign Name"
tryagain:
Dim Heading As String
Dim vDb As ADODB.Connection
Dim vRc As ADODB.Recordset
If ConnectADODB(HouseAcct, vDb) = False Then
   Exit Sub
End If

Sql = "Select * from [HouseAccountData] where [" & FieldName & "] Is Not Null and [" & FieldName & "] > '' and [" & FieldName & "] <> 'None' and left([" & FieldName & "],5) >= '" & Format(SearchDates.Item(1).StartDate, "mm/dd") & "' and left([" & FieldName & "],5) <= '" & Format(SearchDates.Item(1).EndDate, "mm/dd") & "'"
If ConnectADORS(Sql, vDb, vRc) = False Then Exit Sub
If vRc.EOF Then
  Call See_Msg("Nothing to Print", 5, "")
  SetNothing vRc
  Exit Sub
End If
If EventType = "Birthday" Then
  Heading = "Birthday's"
ElseIf EventType = "Membership" Then
  Heading = "Membership Anniversary's"
Else
  Heading = "Anniversary's"
End If
Dim TCharge As Currency, TTip As Currency, TVoid As Currency
   'SetPrinter 0
   SetPage ReportView.View, Heading
   With ReportView.View
    .FontBold = False
    .FontSize = 10
    .TextAlign = taLeftBottom
   End With
Do While Not vRc.EOF
    DrawLine
    PrintLine vRc![scan Name], -1
    If Trim(vRc![Address]) > "" Then
      PrintLine vRc![Address] & "   " & vRc![City] & ",   " & vRc![State] & "   " & vRc![Zip], -1
    End If
    If vRc![Phone] > "" Then
      PrintLine "Phone " & vRc![Phone] & "    " & "Email " & vRc![Email] & "", -1
    End If
      PrintLine "Birthday      " & Format(vRc![Birthday], "MM/DD"), 2500: PrintLine "Anniversary      " & Format(vRc![Anniversary], "MM/DD"), -1
    If EventType = "Expiration" Then
      PrintLine "Membership Expiration      " & Format(vRc![Account Expires On], "MM/DD/YYYY"), -1
    End If
    DrawLine
vRc.MoveNext
Loop
ReportView.View.EndDoc
ReportView.Show 1
DbClose vDb, vRc
Set SearchDates = Nothing
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Events Dates ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand17_Click()
End Sub
Private Sub GETEXPDATE(AccountType As String, ExpiredOnly As Boolean, PurgeAll As Boolean, StartDate As String, EndDate As String)
tryagain:
Dim Heading As String
Dim vDb As ADODB.Connection
Dim vRc As ADODB.Recordset
Dim VRs As ADODB.Recordset
Dim PR As Integer
Dim Cr As Long
Dim PurgeName() As String
Dim RidRcd As Integer
ReDim PurgeName(0) As String
Dim TotalExposure As Currency
Dim AccountExposure As Currency
Dim RTotal As Long
If ConnectADODB(HouseAcct, vDb) = False Then
   Exit Sub
End If
If GetConfigInfo("Functions", "Filter By Location", "", "FUNCTIONS.INI") = "ENABLED" Then
    If ExpiredOnly Then
      Sql = "Select * from [HouseAccountData] Where [LocationName] = '" & LocationName & "' and [Account Expires On] IS NOT NULL Order by [Account Type],[Scan Name] asc"   '&  & "' and [Account Expires On] IS NOT NULL"
    Else
     If AcctType.Text = "All" Or AcctType.Text = "" Then
      Sql = "Select * from [HouseAccountData] Where [LocationName] = '" & LocationName & "' Order by [Account Type],[Scan Name] asc" 'where [Account Type] = '" & AccountType & "'"
     Else
      Sql = "Select * from [HouseAccountData] where [LocationName] = '" & LocationName & "' and [Account Type] = '" & AcctType.Text & "' Order By [Scan Name] asc"
     End If
    End If
Else
    If ExpiredOnly Then
      Sql = "Select * from [HouseAccountData] Where [Account Expires On] IS NOT NULL Order by [Account Type],[Scan Name] asc"   '&  & "' and [Account Expires On] IS NOT NULL"
    Else
     If AcctType.Text = "All" Or AcctType.Text = "" Then
      Sql = "Select * from [HouseAccountData] Order by [Account Type],[Scan Name] asc" 'where [Account Type] = '" & AccountType & "'"
     Else
      Sql = "Select * from [HouseAccountData] where [Account Type] = '" & AcctType.Text & "' Order By [Scan Name] asc"
     End If
    End If
End If

If ConnectADORS(Sql, vDb, vRc) = False Then Exit Sub
If vRc.EOF Then
  Call See_Msg("Nothing to Print for this account type", 5, AcctType.Text)
  SetNothing vRc
  Exit Sub
End If
Heading = AccountType
If StartDate > "" Then
  Heading = AccountType & " Between " & StartDate & " And " & EndDate
Else
  Heading = AccountType
End If
Dim TCharge As Currency, TTip As Currency, TVoid As Currency
   SetPage ReportView.View, Heading
   With ReportView.View
    .FontBold = False
    .FontSize = 8
    .TextAlign = taLeftBottom
   End With
TotalExposure = 0
PR = -1
RTotal = RecordsCount(vRc)
Cr = 0
DrawLine
PrintLine "Name", 3500
PrintLine "Acct", 5500
PrintLine "Type", 7500
PrintLine "Expires ", 10500
PrintLine "Balance ", -1
DrawLine
TotalExposure = 0
Do While Not vRc.EOF
Cr = Cr + 1
NoAcct.Caption = "Calculating " & vRc![scan Name] & "" & " Record " & CStr(Cr) & " Of " & CStr(RTotal)
NoAcct.Refresh
 If ExpiredOnly Then
   If IsDate(vRc![Account Expires On] & "") Then
       AccountExposure = ReturnAllAccountCharges(vRc![scan Name])
       
        With ReportView.View
          PrintLine vRc![scan Name], 3500
          PrintLine vRc![Acct Number] & "", 5500
          PrintLine vRc![Account Type] & "", 7500
          PrintLine vRc![Account Expires On] & "", 10500
          PrintLine Format(AccountExposure, "0.00") & "", -1
          DrawLine
        End With
   End If
  Else
    If IsDate(StartDate) And IsDate(EndDate) Then
     AccountExposure = ReturnAccountCharges(vRc![scan Name], StartDate, EndDate)
    Else
     AccountExposure = ReturnAllAccountCharges(vRc![scan Name])
    End If
    If PurgeAll = True And AccountExposure = 0 And vRc![Account Type] = "Gift Card" Then
       PR = PR + 1
       ReDim Preserve PurgeName(PR)
       PurgeName(PR) = vRc![scan Name]
       GoTo nextR
    End If
    If AccountExposure <> 0 Then
        With ReportView.View
          PrintLine vRc![scan Name], 3500
          PrintLine vRc![Acct Number] & "", 5500
          PrintLine vRc![Account Type] & "", 7500
          PrintLine vRc![Account Expires On] & "", 10500
          PrintLine Format(AccountExposure, "0.00") & "", -1
          DrawLine
        End With
    End If
End If
TotalExposure = TotalExposure + AccountExposure
nextR:
vRc.MoveNext
Loop
DrawLine
For RidRcd = 0 To PR
   Sql = "Delete from [HouseAccountData] where [Scan Name] = '" & PurgeName(RidRcd) & "'"
   vDb.Execute Sql
Next RidRcd
DbClose vDb, vRc
ReportView.View.Text = "Account Total      " & Format(TotalExposure, "0.00") & "" & vbCrLf
ReportView.View.EndDoc
ReportView.Show 1
If ExpiredOnly Then
  If See_Msg("Ok to Delete Records", 4, "") = 6 Then
    Sql = "Select * from [HouseAccountData] "
   If ConnectADODB(HouseAcct, vDb) = False Then
     Exit Sub
   End If
   If ConnectADORS(Sql, vDb, vRc) = False Then Exit Sub
   Do While Not vRc.EOF
     If IsDate(vRc![Account Expires On] & "") Then
       If CDate(vRc![Account Expires On] & "") < Date Then
         vRc.delete
       End If
     End If
     vRc.MoveNext
   Loop
   See_Msg "Completed....", 5, ""
  End If
End If
DbClose vDb, vRc
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Get Exp Date ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub SSCommand18_Click()
End Sub
Private Sub SSCommand19_Click()
End Sub
Private Function is_Required(FD As Integer) As String
   is_Required = ""
   Select Case FD
      Case Is = 0
         If Check_System_Access("Address") > 1 Then
           is_Required = "Address"
         End If
      Case Is = 1
         If Check_System_Access("City") > 1 Then
           is_Required = "City"
         End If
      Case Is = 2
         If Check_System_Access("State") > 1 Then
           is_Required = "State"
         End If
      Case Is = 3
         If Check_System_Access("Zip") > 1 Then
           is_Required = "Zip"
         End If
      Case Is = 4
         If Check_System_Access("Phone") > 1 Then
           is_Required = "Phone"
         End If
      Case Is = 5
         If Check_System_Access("Cell") > 1 Then
           is_Required = "Cell"
         End If
      Case Is = 6
         If Check_System_Access("EMail") > 1 Then
           is_Required = "EMail"
         End If
  End Select
End Function
Private Function GetLocationId() As Long
GetLocationId = 0
'If HouseAcct <> "House Accounts" Then Exit Function
Dim RC As ADODB.Recordset
Dim Db As ADODB.Connection
If ConnectADODB(HouseAcct, Db) Then
  Sql = "Select * from [Locations] where [Location Name] ='" & GetSetting("Micro$ale", "Register Setup", "Site") & "'"
 ConnectADORS Sql, Db, RC
 If Not RC.EOF Then
   GetLocationId = Val(RC.fields("Location Number").Value & "")
   MyLoc.Caption = "Location Name " & GetSetting("Micro$ale", "Register Setup", "Site")
 Else
   MyLoc.Caption = ""
 End If
End If
DbClose Db, RC
End Function
Private Sub SSCommand1_Click()
If Can_Continue(SSCommand1.Caption) = False Then Exit Sub
On Error GoTo ErrorHandler
'If AccountBackUp.Text <> "None" Then
 ' See_Msg "Print Statements Only", 5, ""
 ' Exit Sub
'End If
Dim SearchName As String
Dim ScanName   As String
Dim SD         As Integer
SearchName = Trim(Scan_Data(0).Text)
For SD = 0 To 6
    If Address(SD).Text = "" And is_Required(SD) > "" Then
       See_Msg is_Required(SD), 5, "Required Field"
       Exit Sub
    End If
Next SD
If Trim(SearchName) = "" Then Exit Sub
If Trim(Scan_Data(3).Text) = "" Then
   See_Msg "Must Have Account Number", 5, ""
   Exit Sub
End If
If BDay.Text <> "None" Then
   If IsDate(BDay.Text) = False Then
      See_Msg "Not A Valid Birthday", 5, ""
      Exit Sub
   End If
End If
If AnvDate.Text <> "None" Then
   If IsDate(AnvDate.Text) = False Then
      See_Msg "Not A Valid Anniversary Date", 5, ""
      Exit Sub
   End If
End If

Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If

Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & Trim(Scan_Data(3).Text) & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
  If Trim$(Tb![scan Name]) <> SearchName Then
        See_Msg "Duplicate Account Number", 5, "Can not save data"
        DbClose Db, Tb
        Exit Sub
  End If
End If
SearchName = Trim$(Scan_Data(3).Text)
Sql = "Select * from [HouseAccountData] Where [Scan Name] = '" & Trim$(Scan_Data(0).Text) & "'"
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
  If Trim$(Tb![Acct Number]) <> SearchName Then
        See_Msg "Duplicate Account ", 5, "Can not save data"
        DbClose Db, Tb
        Exit Sub
  End If
End If

DbClose Db, Tb
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If

Sql = "Select * from [HouseAccountData] Where [Scan Name] = '" & Scan_Data(0).Text & "'"
'If Input_Option(50) = "Account #" Then
'  Sql = Sql & " and [Acct Number] = '" & SearchName & "'"
'End If
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Tb.EOF Then
 Tb.AddNew
     'Added for RDS prompt for Account type
    Load_What = "Load Account Groups"
    Find_What = "Select Type"
    Load_What = "Load Account Groups"
    Select_What.Show 1
    If GetSelected = "" Then
     See_Msg "Account Type Required", 5, "Cannot Save...."
     DbClose Db, Tb
     Exit Sub
    End If
    AcctType.Text = GetSelected
    Scan_Data(0).AddItem Scan_Data(0).Text
End If
    Tb![scan Name] = Trim$(Scan_Data(0).Text)
    
    Tb![Max Daily Amount] = Val(Scan_Data(1).Text)
    Tb![Max Weekly Amount] = Val(Scan_Data(2).Text)
    
    Tb![Acct Number] = Trim$(Scan_Data(3).Text)
    Tb![Daily Charge] = 0
    Tb![Suspend] = Suspend.Value
    Tb![Delete Record] = delete.Value
    Tb![Credit Balance] = CreditBalance.Value
    Tb![Address] = Address(0).Text & ""
    Tb![City] = Address(1).Text & ""
    Tb![State] = Address(2).Text & ""
    Tb![Zip] = Address(3).Text & ""
    Tb![Phone] = Address(4).Text & ""
    Tb![Account Type] = AcctType.Text
    Tb![Cell] = Address(5).Text & ""
    Tb![Email] = Address(6).Text & ""
    Tb![Fax] = Address(7).Text & ""
    Tb![Birthday] = BDay.Text & ""
    Tb![Anniversary] = AnvDate.Text & ""
    Tb![Account Expires On] = AExpiration.Text & ""
    Tb![Spc Interest 1] = SPCINT(0).Text
    Tb![Spc Interest 2] = SPCINT(1).Text
    Tb![Spc Interest 3] = SPCINT(2).Text
    Tb![Spc Interest 4] = SPCINT(3).Text
    Tb![Memo] = Memo.Text & ""
    Tb![Meal Limit] = Val(MealLimit.Text & "")
    Tb![Meal Plan] = GetLocationId
    Tb.fields("LocationName").Value = LocationName
    If AttachDiscount.Text = "" Then AttachDiscount.Text = "None"
    Tb![Discount] = AttachDiscount.Text
    Tb![Card Number] = "" 'Encrypt(CardNo.Text, Scan_Data(0).Text & Address(0).Text)
    Tb![Card Expires] = "" 'Encrypt(CardExp.Text, Scan_Data(0).Text & Address(0).Text)
    Tb.Update
    If InmatePic.Text <> "None" And InmatePic.Text > "" Then
      Dim lOffset As Long
      Dim lSize As Long
      Dim sPath As String
      Dim nHandle As Integer
      Dim lSubChunks As Long
      Dim Chunk() As Byte
      Dim nFragmentOffset As Integer
      Dim i As Integer
      Dim lChunks As Long
      SetNothing Tb
      Sql = "Select * from [Photos] Where [Account Name] = '" & Scan_Data(0).Text & "'"
      Sql = Sql & " and [Account Id] = ' " & Scan_Data(3).Text & "'"
      If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
       nHandle = FreeFile
       Open InmatePic.Text For Binary Access Read As nHandle
        lSize = LOF(nHandle)
        If nHandle = 0 Then
                Close nHandle
                See_Msg "Invalid Picture", 5, ""
        Else
         If Tb.EOF Then Tb.AddNew
         Tb![Account Name] = Scan_Data(0).Text
         Tb![Account Id] = Scan_Data(3).Text
         lChunks = lSize \ conChunkSize
         nFragmentOffset = lSize Mod conChunkSize
         ReDim Chunk(nFragmentOffset)
         Get nHandle, , Chunk()
         Tb.fields("Photo").AppendChunk Chunk()
         ReDim Chunk(conChunkSize)
         lOffset = nFragmentOffset
         For i = 1 To lChunks
          Get nHandle, , Chunk()
          Tb.fields("Photo").AppendChunk Chunk()
          lOffset = lOffset + conChunkSize
         Next
         Tb.Update
        End If
       InmateId.Picture = Nothing
    End If
    DbClose Db, Tb
    Call See_Msg("Saved.....", 5, "")
    TrackHouseEvents LoggedOnAs(0).EMPName, " Account Add or Edit " & Scan_Data(0).Text & " " & Scan_Data(3).Text
    Clear_data
    'If OkToDo = True Then Exit_Click
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Assign Name SS1 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Public Sub Clear_data()
Dim SD As Integer
Scan_Data(0).Text = ""
For SD = 1 To 2
Scan_Data(SD).Text = "0.00"
Next SD
Scan_Data(4).Text = "0.00"
Scan_Data(3).Text = ""
Scan_Data(5).Text = "0.00"
Scan_Data(6).Text = "0.00"
Label1(6).Caption = "Current Balance"
For SD = 0 To 7
  Address(SD).Text = ""
Next SD
For SD = 0 To 3
   SPCINT(SD).Text = ""
Next SD
AttachDiscount.Text = "None"
Memo.Text = ""
BDay.Text = "None"
AnvDate.Text = "None"
MealLimit.Text = 0
'Member.Caption = "None"
AExpiration.Text = "Never"
CreditBalance.Value = 1
Suspend.Value = 0
delete.Value = 0
End Sub

Private Sub SSCommand20_Click()
End Sub
Private Sub SSCommand21_Click()
End Sub
Private Sub CleanData()
See_Msg "Checking File.....", 5, ""
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If
   Sql = "select * from [HouseCharges]"
ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
  'If InStr(Tb![Charged To], "'") > 0 Then Stop
  If InStr(Tb![Charged To], "'") > 0 Then
    Tb![Charged To] = Replace(Tb![Charged To], "'", "`")
    Tb.Update
  End If
  If InStr(Tb![Charged To], ",") > 0 Then
    Tb![Charged To] = Replace(Tb![Charged To], ",", "-")
    Tb.Update
  End If
  
  Tb.MoveNext
Loop
SetNothing Tb
   Sql = "select * from [HouseAccountData]"
ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
  If InStr(Tb![scan Name], "'") > 0 Then
    Tb![scan Name] = Replace(Tb![scan Name], "'", "`")
    Tb.Update
  End If
  If InStr(Tb![scan Name], ",") > 0 Then
    Tb![scan Name] = Replace(Tb![scan Name], ",", "-")
    Tb.Update
  End If
  
  Tb.MoveNext
Loop
SetNothing Tb
Sql = "select * from HouseChargeDetail"
ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
  If InStr(Tb![Account Name], "'") > 0 Then
    Tb![Account Name] = Replace(Tb![Account Name], "'", "`")
    Tb.Update
  End If
  If InStr(Tb![Account Name], ",") > 0 Then
    Tb![Account Name] = Replace(Tb![Account Name], ",", "-")
    Tb.Update
  End If
  
  Tb.MoveNext
Loop
SetNothing Tb
   Sql = "Select * from [HousePayments]"
ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
  If InStr(Tb![Account Name], "'") > 0 Then
    Tb![Account Name] = Replace(Tb![Account Name], "'", "`")
    Tb.Update
  End If
  If InStr(Tb![Account Name], ",") > 0 Then
    Tb![Account Name] = Replace(Tb![Account Name], ",", "-")
    Tb.Update
  End If
  
  Tb.MoveNext
Loop
SetNothing Tb
   Sql = "select * from [Starting Balance]"
ConnectADORS Sql, Db, Tb
Do While Not Tb.EOF
  If InStr(Tb![Account Name], "'") > 0 Then
    Tb![Account Name] = Replace(Tb![Account Name], "'", "`")
    Tb.Update
  End If
  If InStr(Tb![Account Name], ",") > 0 Then
    Tb![Account Name] = Replace(Tb![Account Name], ",", "-")
    Tb.Update
  End If
  
  Tb.MoveNext
Loop
See_Msg "Finished......", 5, "Clear up Complete"
DbClose Db, Tb
End Sub
Private Sub SSCommand3_Click()
End Sub

Private Sub SSCommand4_Click()
On Error GoTo ErrorHandler
If Can_Continue(SSCommand4.Caption) = False Then Exit Sub
If AccountBackUp.Text <> "None" Then
  See_Msg "Print Statements Only", 5, ""
  Exit Sub
End If
If See_Msg("Reset Account Activity for " & Scan_Data(0).Text, 4, "Have You Printed Statement?") = 6 Then
   If See_Msg("Ok To Delete Account Activity", 4, "All Data will Be Deleted.") = 6 Then
      ResetIndAccount Scan_Data(0).Text
   End If
End If
If See_Msg("Are you sure ?", 4, "Deleting Name") <> 6 Then Exit Sub
If See_Msg("Account will be Removed....", 4, "OK To Delete ?") <> 6 Then Exit Sub
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If

Sql = "select * from [HouseCharges] where [Scan Number] = '" & Scan_Data(3).Text & "'"
'If HouseAcct = "House Accounts" Then
'  Sql = Sql & " and [location Name] = '" & LocationName & "'"
'End If
If ConnectADORS(Sql, Db, Tb) = False Then Exit Sub
If Not Tb.EOF Then
   See_Msg "Active Account", 5, "Cannot Remove"
   DbClose Db, Tb
   Exit Sub
End If
Sql = "Delete from [HouseAccountData] where [Scan Name] = '" & Scan_Data(0).Text & "'"
Db.Execute Sql
Sql = "Delete from [HousePayments] where [Acccount Name] = '" & Scan_Data(0).Text & "'"
Db.Execute Sql
DbClose Db, Tb
TrackHouseEvents LoggedOnAs(0).EMPName, "Account Removed " & Scan_Data(0).Text
Clear_data
ErrorHandler:
End Sub
Private Sub SSCommand5_Click()
With MemberPage
   .Visible = False
   .Left = -4000
End With
'PaidIn.Show 1
'GetAccount
End Sub
Private Sub ResetIndAccount(AccountName As String)
    Dim DDate As String
    DDate = Format(Date, "MM/DD/YYYY")
    If See_Msg("Reset Account Activity ?", 4, "Thru " & DDate) <> 6 Then Exit Sub
    Dim vDb As ADODB.Connection
    If ConnectADODB(HouseAcct, vDb) = False Then
       Exit Sub
    End If
    Sql = "Delete from [HouseCharges]"
    Sql = Sql & " Where [Charged To] = '" & AccountName & "'"
    Sql = Sql & " and TRY_CAST([Check date] AS datetime) Between '01/01/2000' and '" & CDate(DDate) & "'"
    vDb.Execute Sql
    '**********************
      Sql = "Delete from HouseChargeDetail"
      Sql = Sql & " Where [Account Name] = '" & AccountName & "'"
      Sql = Sql & " and TRY_CAST([Check date] AS datetime) Between '01/01/2000' and '" & CDate(DDate) & "'"
      vDb.Execute Sql
    '**********************
      Sql = "Delete from [HousePayments] where [Account Name] = '" & AccountName & "'"
      Sql = Sql & " and TRY_CAST([date] AS datetime) Between '01/01/2000' and '" & CDate(DDate) & "'"
      vDb.Execute Sql
      Sql = "Delete from [Starting Balance] where [Account Name] = '" & AccountName & "'"
      Sql = Sql & " and TRY_CAST([date] AS datetime) Between '01/01/2000' and '" & CDate(DDate) & "'"
      vDb.Execute Sql
      See_Msg "Completed....", 5, "File Purge Completed."
      DbClose vDb, Nothing
End Sub

Private Sub TargetCard_Click()
Dim Db As ADODB.Connection
Dim RC As ADODB.Recordset
If ConnectADODB(HouseAcct, Db) = False Then
   Exit Sub
End If

Sql = "Select * from [HouseAccountData] Where [Acct Number] = '" & TargetCard.Text & "'"
If ConnectADORS(Sql, Db, RC) = False Then Exit Sub
If RC.EOF Then
  See_Msg "Account Number Not Found", 5, ""
  TargetCard.Text = ""
  TargetName.Caption = ""
  DbClose Db, RC
  Exit Sub
Else
  TargetName.Caption = RC![scan Name] & ""
End If
DbClose Db, RC
End Sub
Private Sub TargetCard_KeyPress(KeyAscii As Integer)
    KeyAscii = ReturnSentinal(KeyAscii)
    Select Case KeyAscii
        Case Is = 39
          KeyAscii = 96
        Case Is = 33, 34, 44
          KeyAscii = 0
          Exit Sub
    End Select
End Sub

Private Sub TFBAL_Click()
If Can_Continue(TFBAL.Caption) = False Then Exit Sub
'LoadAccounts
LoadAccounts
CenterObject Me, TransferBalances
SourceCard.Text = ""
TargetCard.Text = ""
TargetName.Caption = ""
SourceName.Caption = ""
TransferBalances.Visible = True
TransferBalances.ZOr