VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form Time_Sheet 
   Appearance      =   0  'Flat
   BackColor       =   &H002E2620&
   BorderStyle     =   0  'None
   Caption         =   "Time Sheet"
   ClientHeight    =   11460
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   18990
   FillColor       =   &H002E2620&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H002E2620&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11460
   ScaleWidth      =   18990
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame ShiftPunch 
      Appearance      =   0  'Flat
      BackColor       =   &H002E2620&
      Caption         =   "Shift Punch In / Out"
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
      Height          =   7815
      Left            =   -3000
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   11775
      Begin ComctlLib.ListView PunchList 
         Height          =   5055
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8916
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin BTNENHLib4.BtnEnh Button_1 
         Height          =   975
         Left            =   9960
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1720
         _StockProps     =   66
         Caption         =   "Punch In"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
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
         textCaption     =   "Time_Sheet.frx":0000
         textLT          =   "Time_Sheet.frx":0070
         textCT          =   "Time_Sheet.frx":0088
         textRT          =   "Time_Sheet.frx":00A0
         textLM          =   "Time_Sheet.frx":00B8
         textRM          =   "Time_Sheet.frx":00D0
         textLB          =   "Time_Sheet.frx":00E8
         textCB          =   "Time_Sheet.frx":0100
         textRB          =   "Time_Sheet.frx":0118
         colorBack       =   "Time_Sheet.frx":0130
         colorIntern     =   "Time_Sheet.frx":015A
         colorMO         =   "Time_Sheet.frx":0184
         colorFocus      =   "Time_Sheet.frx":01AE
         colorDisabled   =   "Time_Sheet.frx":01D8
         colorPressed    =   "Time_Sheet.frx":0202
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
      Begin BTNENHLib4.BtnEnh Button_2 
         Height          =   1095
         Left            =   9960
         TabIndex        =   8
         Top             =   6120
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1931
         _StockProps     =   66
         Caption         =   "Break"
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
         textCaption     =   "Time_Sheet.frx":022C
         textLT          =   "Time_Sheet.frx":0296
         textCT          =   "Time_Sheet.frx":02AE
         textRT          =   "Time_Sheet.frx":02C6
         textLM          =   "Time_Sheet.frx":02DE
         textRM          =   "Time_Sheet.frx":02F6
         textLB          =   "Time_Sheet.frx":030E
         textCB          =   "Time_Sheet.frx":0326
         textRB          =   "Time_Sheet.frx":033E
         colorBack       =   "Time_Sheet.frx":0356
         colorIntern     =   "Time_Sheet.frx":0380
         colorMO         =   "Time_Sheet.frx":03AA
         colorFocus      =   "Time_Sheet.frx":03D4
         colorDisabled   =   "Time_Sheet.frx":03FE
         colorPressed    =   "Time_Sheet.frx":0428
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
      Begin BTNENHLib4.BtnEnh BtnEnh1 
         Height          =   975
         Left            =   9960
         TabIndex        =   12
         Top             =   1560
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1720
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
         Shape           =   1
         Surface         =   15
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Time_Sheet.frx":0452
         textLT          =   "Time_Sheet.frx":04BA
         textCT          =   "Time_Sheet.frx":04D2
         textRT          =   "Time_Sheet.frx":04EA
         textLM          =   "Time_Sheet.frx":0502
         textRM          =   "Time_Sheet.frx":051A
         textLB          =   "Time_Sheet.frx":0532
         textCB          =   "Time_Sheet.frx":054A
         textRB          =   "Time_Sheet.frx":0562
         colorBack       =   "Time_Sheet.frx":057A
         colorIntern     =   "Time_Sheet.frx":05A4
         colorMO         =   "Time_Sheet.frx":05CE
         colorFocus      =   "Time_Sheet.frx":05F8
         colorDisabled   =   "Time_Sheet.frx":0622
         colorPressed    =   "Time_Sheet.frx":064C
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
      Begin BTNENHLib4.BtnEnh BtnEnh2 
         Height          =   855
         Left            =   9960
         TabIndex        =   14
         Top             =   4560
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Print Time"
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
         textCaption     =   "Time_Sheet.frx":0676
         textLT          =   "Time_Sheet.frx":06EA
         textCT          =   "Time_Sheet.frx":0702
         textRT          =   "Time_Sheet.frx":071A
         textLM          =   "Time_Sheet.frx":0732
         textRM          =   "Time_Sheet.frx":074A
         textLB          =   "Time_Sheet.frx":0762
         textCB          =   "Time_Sheet.frx":077A
         textRB          =   "Time_Sheet.frx":0792
         colorBack       =   "Time_Sheet.frx":07AA
         colorIntern     =   "Time_Sheet.frx":07D4
         colorMO         =   "Time_Sheet.frx":07FE
         colorFocus      =   "Time_Sheet.frx":0828
         colorDisabled   =   "Time_Sheet.frx":0852
         colorPressed    =   "Time_Sheet.frx":087C
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
         Office2003ColorNorth=   12615680
         Office2003ColorSouth=   12615680
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   12615680
         Office2003ColorSouthMouseOver=   12615680
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   5395026
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   16777215
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh MealBreak 
         Height          =   615
         Left            =   240
         TabIndex        =   52
         Top             =   6120
         Visible         =   0   'False
         Width           =   2895
         _Version        =   589828
         _ExtentX        =   5106
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Meal Break"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
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
         textCaption     =   "Time_Sheet.frx":08A6
         textLT          =   "Time_Sheet.frx":091A
         textCT          =   "Time_Sheet.frx":0932
         textRT          =   "Time_Sheet.frx":094A
         textLM          =   "Time_Sheet.frx":0962
         textRM          =   "Time_Sheet.frx":097A
         textLB          =   "Time_Sheet.frx":0992
         textCB          =   "Time_Sheet.frx":09AA
         textRB          =   "Time_Sheet.frx":09C2
         colorBack       =   "Time_Sheet.frx":09DA
         colorIntern     =   "Time_Sheet.frx":0A04
         colorMO         =   "Time_Sheet.frx":0A2E
         colorFocus      =   "Time_Sheet.frx":0A58
         colorDisabled   =   "Time_Sheet.frx":0A82
         colorPressed    =   "Time_Sheet.frx":0AAC
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
      Begin BTNENHLib4.BtnEnh RegularBreak 
         Height          =   615
         Left            =   240
         TabIndex        =   53
         Top             =   6840
         Visible         =   0   'False
         Width           =   2895
         _Version        =   589828
         _ExtentX        =   5106
         _ExtentY        =   1085
         _StockProps     =   66
         Caption         =   "Regular Break"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
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
         textCaption     =   "Time_Sheet.frx":0AD6
         textLT          =   "Time_Sheet.frx":0B50
         textCT          =   "Time_Sheet.frx":0B68
         textRT          =   "Time_Sheet.frx":0B80
         textLM          =   "Time_Sheet.frx":0B98
         textRM          =   "Time_Sheet.frx":0BB0
         textLB          =   "Time_Sheet.frx":0BC8
         textCB          =   "Time_Sheet.frx":0BE0
         textRB          =   "Time_Sheet.frx":0BF8
         colorBack       =   "Time_Sheet.frx":0C10
         colorIntern     =   "Time_Sheet.frx":0C3A
         colorMO         =   "Time_Sheet.frx":0C64
         colorFocus      =   "Time_Sheet.frx":0C8E
         colorDisabled   =   "Time_Sheet.frx":0CB8
         colorPressed    =   "Time_Sheet.frx":0CE2
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
      Begin ComctlLib.ListView BreakView 
         Height          =   1335
         Left            =   3240
         TabIndex        =   54
         Top             =   6120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin BTNENHLib4.BtnEnh BtnEnh8 
         Height          =   855
         Left            =   9960
         TabIndex        =   55
         Top             =   3600
         Width           =   1575
         _Version        =   589828
         _ExtentX        =   2778
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "View Pay Period"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
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
         textCaption     =   "Time_Sheet.frx":0D0C
         textLT          =   "Time_Sheet.frx":0D8A
         textCT          =   "Time_Sheet.frx":0DA2
         textRT          =   "Time_Sheet.frx":0DBA
         textLM          =   "Time_Sheet.frx":0DD2
         textRM          =   "Time_Sheet.frx":0DEA
         textLB          =   "Time_Sheet.frx":0E02
         textCB          =   "Time_Sheet.frx":0E1A
         textRB          =   "Time_Sheet.frx":0E32
         colorBack       =   "Time_Sheet.frx":0E4A
         colorIntern     =   "Time_Sheet.frx":0E74
         colorMO         =   "Time_Sheet.frx":0E9E
         colorFocus      =   "Time_Sheet.frx":0EC8
         colorDisabled   =   "Time_Sheet.frx":0EF2
         colorPressed    =   "Time_Sheet.frx":0F1C
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
         Office2003ColorNorth=   12615680
         Office2003ColorSouth=   12615680
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   12615680
         Office2003ColorSouthMouseOver=   12615680
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   3302400
         Office2003ColorSouthPressed=   2642176
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin VB.Label TTime 
         BackStyle       =   0  'Transparent
         Caption         =   "EmpName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5640
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes are calculated in hundredths. ** indicates altered shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   5640
         Width           =   6495
      End
      Begin VB.Label EMPName 
         BackStyle       =   0  'Transparent
         Caption         =   "EmpName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame SelectAShift 
      Appearance      =   0  'Flat
      BackColor       =   &H002E2620&
      Caption         =   "Select a shift to add a break in"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   15840
      TabIndex        =   49
      Top             =   10200
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ListBox ShiftList 
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
         Height          =   4980
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   3375
      End
      Begin BTNENHLib4.BtnEnh BtnEnh6 
         Height          =   975
         Left            =   120
         TabIndex        =   50
         Top             =   5760
         Width           =   3375
         _Version        =   589828
         _ExtentX        =   5953
         _ExtentY        =   1720
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
         Shape           =   1
         Surface         =   15
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Time_Sheet.frx":0F46
         textLT          =   "Time_Sheet.frx":0FAE
         textCT          =   "Time_Sheet.frx":0FC6
         textRT          =   "Time_Sheet.frx":0FDE
         textLM          =   "Time_Sheet.frx":0FF6
         textRM          =   "Time_Sheet.frx":100E
         textLB          =   "Time_Sheet.frx":1026
         textCB          =   "Time_Sheet.frx":103E
         textRB          =   "Time_Sheet.frx":1056
         colorBack       =   "Time_Sheet.frx":106E
         colorIntern     =   "Time_Sheet.frx":1098
         colorMO         =   "Time_Sheet.frx":10C2
         colorFocus      =   "Time_Sheet.frx":10EC
         colorDisabled   =   "Time_Sheet.frx":1116
         colorPressed    =   "Time_Sheet.frx":1140
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
   End
   Begin VB.Frame BreakSheet 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Break Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   13560
      TabIndex        =   44
      Top             =   -720
      Visible         =   0   'False
      Width           =   9615
      Begin VSFlex8DAOCtl.VSFlexGrid BreakEdit 
         Height          =   4815
         Left            =   2640
         TabIndex        =   47
         Top             =   240
         Width           =   6855
         _cx             =   2003251003
         _cy             =   2003247405
         Appearance      =   0
         BorderStyle     =   1
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
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483634
         BackColorBkg    =   0
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1000
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
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
      Begin BTNENHLib4.BtnEnh BtnEnh7 
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   4440
         Width           =   2295
         _Version        =   589828
         _ExtentX        =   4048
         _ExtentY        =   1085
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
         Shape           =   1
         Surface         =   15
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Time_Sheet.frx":116A
         textLT          =   "Time_Sheet.frx":11D2
         textCT          =   "Time_Sheet.frx":11EA
         textRT          =   "Time_Sheet.frx":1202
         textLM          =   "Time_Sheet.frx":121A
         textRM          =   "Time_Sheet.frx":1232
         textLB          =   "Time_Sheet.frx":124A
         textCB          =   "Time_Sheet.frx":1262
         textRB          =   "Time_Sheet.frx":127A
         colorBack       =   "Time_Sheet.frx":1292
         colorIntern     =   "Time_Sheet.frx":12BC
         colorMO         =   "Time_Sheet.frx":12E6
         colorFocus      =   "Time_Sheet.frx":1310
         colorDisabled   =   "Time_Sheet.frx":133A
         colorPressed    =   "Time_Sheet.frx":1364
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
      Begin BTNENHLib4.BtnEnh SaveKey 
         Height          =   615
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
         _Version        =   589828
         _ExtentX        =   4048
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
         Shape           =   1
         Surface         =   15
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Time_Sheet.frx":138E
         textLT          =   "Time_Sheet.frx":13F6
         textCT          =   "Time_Sheet.frx":140E
         textRT          =   "Time_Sheet.frx":1426
         textLM          =   "Time_Sheet.frx":143E
         textRM          =   "Time_Sheet.frx":1456
         textLB          =   "Time_Sheet.frx":146E
         textCB          =   "Time_Sheet.frx":1486
         textRB          =   "Time_Sheet.frx":149E
         colorBack       =   "Time_Sheet.frx":14B6
         colorIntern     =   "Time_Sheet.frx":14E0
         colorMO         =   "Time_Sheet.frx":150A
         colorFocus      =   "Time_Sheet.frx":1534
         colorDisabled   =   "Time_Sheet.frx":155E
         colorPressed    =   "Time_Sheet.frx":1588
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
         Office2003ColorNorth=   32768
         Office2003ColorSouth=   32768
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   32768
         Office2003ColorSouthMouseOver=   5395026
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   32768
         Office2003ColorSouthPressed=   2302755
         Office2003ColorBorderPressed=   16777215
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
   End
   Begin VB.Frame TimeEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H002E2620&
      Caption         =   "Time Clock Editor"
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
      Height          =   8895
      Left            =   1800
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   13335
      Begin VB.CheckBox LoadOldToNew 
         BackColor       =   &H002E2620&
         Caption         =   "Load in record edit order"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   8160
         Width           =   3015
      End
      Begin VB.ComboBox Employees 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   19
         Text            =   "Employees"
         Top             =   480
         Width           =   4695
      End
      Begin VSFlex8DAOCtl.VSFlexGrid TimeSheet 
         Height          =   5535
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   12735
         _cx             =   1977440191
         _cy             =   1977427491
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483634
         BackColorBkg    =   0
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1000
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   0
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
      Begin BTNENHLib4.BtnEnh DeleteTime 
         Height          =   855
         Left            =   6000
         TabIndex        =   23
         Top             =   6960
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Delete Time"
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
         textCaption     =   "Time_Sheet.frx":15B2
         textLT          =   "Time_Sheet.frx":1628
         textCT          =   "Time_Sheet.frx":1640
         textRT          =   "Time_Sheet.frx":1658
         textLM          =   "Time_Sheet.frx":1670
         textRM          =   "Time_Sheet.frx":1688
         textLB          =   "Time_Sheet.frx":16A0
         textCB          =   "Time_Sheet.frx":16B8
         textRB          =   "Time_Sheet.frx":16D0
         colorBack       =   "Time_Sheet.frx":16E8
         colorIntern     =   "Time_Sheet.frx":1712
         colorMO         =   "Time_Sheet.frx":173C
         colorFocus      =   "Time_Sheet.frx":1766
         colorDisabled   =   "Time_Sheet.frx":1790
         colorPressed    =   "Time_Sheet.frx":17BA
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
      Begin BTNENHLib4.BtnEnh ChangeTime 
         Height          =   855
         Left            =   9600
         TabIndex        =   24
         Top             =   6960
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Save Changes"
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
         textCaption     =   "Time_Sheet.frx":17E4
         textLT          =   "Time_Sheet.frx":185C
         textCT          =   "Time_Sheet.frx":1874
         textRT          =   "Time_Sheet.frx":188C
         textLM          =   "Time_Sheet.frx":18A4
         textRM          =   "Time_Sheet.frx":18BC
         textLB          =   "Time_Sheet.frx":18D4
         textCB          =   "Time_Sheet.frx":18EC
         textRB          =   "Time_Sheet.frx":1904
         colorBack       =   "Time_Sheet.frx":191C
         colorIntern     =   "Time_Sheet.frx":1946
         colorMO         =   "Time_Sheet.frx":1970
         colorFocus      =   "Time_Sheet.frx":199A
         colorDisabled   =   "Time_Sheet.frx":19C4
         colorPressed    =   "Time_Sheet.frx":19EE
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
         Office2003ColorNorth=   32768
         Office2003ColorSouth=   32768
         Office2003ColorBorder=   16777215
         Office2003ColorNorthMouseOver=   32768
         Office2003ColorSouthMouseOver=   32768
         Office2003ColorBorderMouseOver=   65408
         Office2003ColorNorthPressed=   32768
         Office2003ColorSouthPressed=   32768
         Office2003ColorBorderPressed=   65408
         Office2007OuterBorderWidth=   1
         Office2007InnerBorderWidth=   1
         Office2003BorderWidth=   3
      End
      Begin BTNENHLib4.BtnEnh Exit 
         Height          =   855
         Left            =   11400
         TabIndex        =   25
         Top             =   6960
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
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
         Shape           =   1
         Surface         =   15
         LogPixels       =   96
         TextureMode     =   0
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "Time_Sheet.frx":1A18
         textLT          =   "Time_Sheet.frx":1A80
         textCT          =   "Time_Sheet.frx":1A98
         textRT          =   "Time_Sheet.frx":1AB0
         textLM          =   "Time_Sheet.frx":1AC8
         textRM          =   "Time_Sheet.frx":1AE0
         textLB          =   "Time_Sheet.frx":1AF8
         textCB          =   "Time_Sheet.frx":1B10
         textRB          =   "Time_Sheet.frx":1B28
         colorBack       =   "Time_Sheet.frx":1B40
         colorIntern     =   "Time_Sheet.frx":1B6A
         colorMO         =   "Time_Sheet.frx":1B94
         colorFocus      =   "Time_Sheet.frx":1BBE
         colorDisabled   =   "Time_Sheet.frx":1BE8
         colorPressed    =   "Time_Sheet.frx":1C12
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
         Height          =   855
         Left            =   7800
         TabIndex        =   41
         Top             =   6960
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Edit Tips"
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
         textCaption     =   "Time_Sheet.frx":1C3C
         textLT          =   "Time_Sheet.frx":1CAE
         textCT          =   "Time_Sheet.frx":1CC6
         textRT          =   "Time_Sheet.frx":1CDE
         textLM          =   "Time_Sheet.frx":1CF6
         textRM          =   "Time_Sheet.frx":1D0E
         textLB          =   "Time_Sheet.frx":1D26
         textCB          =   "Time_Sheet.frx":1D3E
         textRB          =   "Time_Sheet.frx":1D56
         colorBack       =   "Time_Sheet.frx":1D6E
         colorIntern     =   "Time_Sheet.frx":1D98
         colorMO         =   "Time_Sheet.frx":1DC2
         colorFocus      =   "Time_Sheet.frx":1DEC
         colorDisabled   =   "Time_Sheet.frx":1E16
         colorPressed    =   "Time_Sheet.frx":1E40
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
      Begin BTNENHLib4.BtnEnh BtnEnh5 
         Height          =   855
         Left            =   4200
         TabIndex        =   46
         Top             =   6960
         Width           =   1695
         _Version        =   589828
         _ExtentX        =   2990
         _ExtentY        =   1508
         _StockProps     =   66
         Caption         =   "Show Breaks"
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
         textCaption     =   "Time_Sheet.frx":1E6A
         textLT          =   "Time_Sheet.frx":1EE0
         textCT          =   "Time_Sheet.frx":1EF8
         textRT          =   "Time_Sheet.frx":1F10
         textLM          =   "Time_Sheet.frx":1F28
         textRM          =   "Time_Sheet.frx":1F40
         textLB          =   "Time_Sheet.frx":1F58
         textCB          =   "Time_Sheet.frx":1F70
         textRB          =   "Time_Sheet.frx":1F88
         colorBack       =   "Time_Sheet.frx":1FA0
         colorIntern     =   "Time_Sheet.frx":1FCA
         colorMO         =   "Time_Sheet.frx":1FF4
         colorFocus      =   "Time_Sheet.frx":201E
         colorDisabled   =   "Time_Sheet.frx":2048
         colorPressed    =   "Time_Sheet.frx":2072
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
      Begin VB.Label WeekOf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   5700
         TabIndex        =   43
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Red Indicates Altered Shifts."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1080
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Total_Hours 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Hours :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   7560
         Width           =   3375
      End
      Begin VB.Label TM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "time"
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
         Left            =   5520
         TabIndex        =   21
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label DT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "date"
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
         Left            =   10080
         TabIndex        =   20
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Emp_Name 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   7080
         Width           =   3615
      End
   End
   Begin VB.PictureBox Work_Area 
      Appearance      =   0  'Flat
      BackColor       =   &H002E2620&
      FillColor       =   &H002E2620&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   17400
      ScaleHeight     =   7305
      ScaleWidth      =   5220
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   5250
      Begin VB.ListBox Work_Dept 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4695
      End
      Begin VB.ListBox Work_Dept 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   3015
      End
      Begin VB.ListBox Work_Dept 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin BTNENHLib4.BtnEnh SSCommand1 
         Height          =   780
         Left            =   1320
         TabIndex        =   5
         Top             =   6360
         Width           =   2355
         _Version        =   589828
         _ExtentX        =   4154
         _ExtentY        =   1376
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
         textCaption     =   "Time_Sheet.frx":209C
         textLT          =   "Time_Sheet.frx":2104
         textCT          =   "Time_Sheet.frx":211C
         textRT          =   "Time_Sheet.frx":2134
         textLM          =   "Time_Sheet.frx":214C
         textRM          =   "Time_Sheet.frx":2164
         textLB          =   "Time_Sheet.frx":217C
         textCB          =   "Time_Sheet.frx":2194
         textRB          =   "Time_Sheet.frx":21AC
         colorBack       =   "Time_Sheet.frx":21C4
         colorIntern     =   "Time_Sheet.frx":21EE
         colorMO         =   "Time_Sheet.frx":2218
         colorFocus      =   "Time_Sheet.frx":2242
         colorDisabled   =   "Time_Sheet.frx":226C
         colorPressed    =   "Time_Sheet.frx":2296
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Job Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   4695
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   5895
         Left            =   120
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   14040
      Top             =   7560
   End
   Begin VB.Frame Emp_Validate 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Validation Required"
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
      Height          =   6135
      Left            =   16080
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
      Begin BTNENHLib4.BtnEnh BtnEnh3 
         Height          =   1470
         Left            =   2400
         TabIndex        =   34
         Top             =   4200
         Width           =   2055
         _Version        =   589828
         _ExtentX        =   3625
         _ExtentY        =   2593
         _StockProps     =   66
         Caption         =   "I Agree with this time"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
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
         textCaption     =   "Time_Sheet.frx":22C0
         textLT          =   "Time_Sheet.frx":234C
         textCT          =   "Time_Sheet.frx":2364
         textRT          =   "Time_Sheet.frx":237C
         textLM          =   "Time_Sheet.frx":2394
         textRM          =   "Time_Sheet.frx":23AC
         textLB          =   "Time_Sheet.frx":23C4
         textCB          =   "Time_Sheet.frx":23DC
         textRB          =   "Time_Sheet.frx":23F4
         colorBack       =   "Time_Sheet.frx":240C
         colorIntern     =   "Time_Sheet.frx":2436
         colorMO         =   "Time_Sheet.frx":2460
         colorFocus      =   "Time_Sheet.frx":248A
         colorDisabled   =   "Time_Sheet.frx":24B4
         colorPressed    =   "Time_Sheet.frx":24DE
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
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   1515
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Width           =   2055
         _Version        =   589828
         _ExtentX        =   3625
         _ExtentY        =   2672
         _StockProps     =   66
         Caption         =   "I Do Not Agree with this time"
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
         textCaption     =   "Time_Sheet.frx":2508
         textLT          =   "Time_Sheet.frx":25A2
         textCT          =   "Time_Sheet.frx":25BA
         textRT          =   "Time_Sheet.frx":25D2
         textLM          =   "Time_Sheet.frx":25EA
         textRM          =   "Time_Sheet.frx":2602
         textLB          =   "Time_Sheet.frx":261A
         textCB          =   "Time_Sheet.frx":2632
         textRB          =   "Time_Sheet.frx":264A
         colorBack       =   "Time_Sheet.frx":2662
         colorIntern     =   "Time_Sheet.frx":268C
         colorMO         =   "Time_Sheet.frx":26B6
         colorFocus      =   "Time_Sheet.frx":26E0
         colorDisabled   =   "Time_Sheet.frx":270A
         colorPressed    =   "Time_Sheet.frx":2734
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
      Begin VB.Label TWT 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2760
         TabIndex        =   40
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label ETO 
         BackStyle       =   0  'Transparent
         Caption         =   "TimeO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label BTI 
         BackStyle       =   0  'Transparent
         Caption         =   "BreakI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label BTO 
         BackStyle       =   0  'Transparent
         Caption         =   "BreakO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   37
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label STM 
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   36
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Worked"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Break In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Break Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label TimeON 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label ShiftInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Time"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Label NoJobCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Job Code Assigned. Job Code Required!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   13815
   End
End
Attribute VB_Name = "Time_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NxR          As Integer
Public CloseForm    As Integer
Public DeleteRecord As Integer
Public ForGot       As Integer
Public MRGOK        As String
Public FormIsDirty  As Boolean
Public IApprove     As String
Public RecordId     As String
Public EditDate     As Date

Private Type EmployeeData
      FirstName As String
      LastName As String
      JobCode  As String
      ScheduleId As String
      DepartmentId As String
      JobRoleId As String
End Type

Private Function GetCurrentPayrollWeek() As String
Dim DOW  As Integer
DOW = DatePart("w", Date, WeekStartsOn, vbFirstFullWeek)
If DOW = 1 Then
   GetCurrentPayrollWeek = Date
Else
   GetCurrentPayrollWeek = DateAdd("d", -(DOW - 1), Date)
End If
End Function
Private Function GetEmployee(ByVal JobWorking As String) As EmployeeData
  Dim EMPName()     As String
  Dim Db            As ADODB.Connection
  Dim Tb            As ADODB.Recordset
  Dim SS            As [7Shifts].SevenShift
  Set SS = New [7Shifts].SevenShift
  EMPName = Split(Employees.Text, " ")
  If JobWorking = "" Then
     GetEmployee.FirstName = EMPName(0)
     GetEmployee.LastName = EMPName(1)
     GetEmployee.JobCode = "NotLoaded"
     Exit Function
  End If
  ConnectADODB "MSaleData", Db
     Sql = "Select top 1 * from [Job Codes] where "
     Sql = Sql & " [First Name] = '" & EMPName(0) & "'"
     Sql = Sql & " and [Last Name] = '" & EMPName(1) & "'"
     Sql = Sql & " and [Job Code] = '" & JobWorking & "'"
  ConnectADORS Sql, Db, Tb
  If Not Tb.EOF Then
     GetEmployee.FirstName = EMPName(0)
     GetEmployee.LastName = EMPName(1)
     GetEmployee.JobCode = Tb.fields("Job Code").Value & ""
     GetEmployee.ScheduleId = Tb.fields("Company Id").Value & ""
     
     GetEmployee.JobRoleId = Nz(SS.ReturnRoleId(Tb.fields("Job Code").Value & ""), 0)
     GetEmployee.DepartmentId = Nz(SS.ReturnDepartmentId(Tb.fields("Job Code").Value & ""), 0)
  Else
     GetEmployee.FirstName = "NotFound"
  End If
  Set SS = Nothing
  DbClose Db, Tb
End Function
Private Function PunchCheck() As Boolean
On Error GoTo ERH
Dim Mins   As Long
Dim HMins  As String
PunchCheck = True
MRGOK = ""
If Val(Input_Option(22)) = 0 Then Exit Function
'If Check_Option(41) = 1 Then
'    HMins = QuarterRound(Mins, "")
'Else
    Mins = Minute(Time)
'End If

If LoggedOnAs(0).EmpAccessLevel >= Access_Required("Managers Menu") Then
  PunchCheck = True
  MRGOK = "Mgr Ok"
  Exit Function
End If

If Mins > 60 - Val(Input_Option(22)) Then
   PunchCheck = True
   Exit Function
ElseIf Mins < Val(Input_Option(22)) Then
   PunchCheck = True
   Exit Function
ElseIf Mins > 30 - Val(Input_Option(22)) And Mins < 30 + Val(Input_Option(22)) Then
   PunchCheck = True
   Exit Function
Else
   PunchCheck = False
   Call See_Msg("Please See a Manager", 5, "")
   Exit Function
End If

Exit Function
ERH:
PunchCheck = True
End Function
Private Sub BreakEdit_Click()
  Dim MyDateString  As String
  Dim InDate        As String
  Dim R             As Integer
  Dim Db            As ADODB.Connection
  Dim Tb            As ADODB.Recordset
  
  With BreakEdit
  Select Case .Col
     Case Is = 0
     
     Case Is = 1
        If .TextMatrix(.Row, 1) > "" Then
           If See_Msg("Remove Break?", 4, "") = 6 Then
             ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
               Sql = "delete From [Breaks] where Name = '" & Employees.Text & "'"
               Sql = Sql & " and [RecordId] = '" & .TextMatrix(.Row, 0) & "' and TRY_CAST([TimeOut] AS datetime) = '" & .TextMatrix(.Row, 2) & "'"
               Sql = Sql & " and TRY_CAST([TimeIn] AS datetime) = '" & .TextMatrix(.Row, 3) & "'"
             Db.Execute Sql
             Db.Close
             BtnEnh5_Click
           End If
        End If
     Case Is = 2, 3
       If .TextMatrix(.Row, 0) = "" Then
        ShiftList.Clear
        For R = 1 To TimeSheet.rows - 1
          If TimeSheet.TextMatrix(R, 2) = "" Then Exit For
          ShiftList.AddItem TimeSheet.TextMatrix(R, 1) & " " & TimeSheet.TextMatrix(R, 2)
        Next R
        SelectAShift.Visible = True
        SelectAShift.ZOrder 0
        CenterObject Me, SelectAShift
        Exit Sub
       End If
       Dim PKT As DateOrTime
       Set PKT = New DateOrTime
       InDate = PKT.ReturnDate
       MyDateString = Format(InDate & " " & GetTime(""), "MM/DD/YYYY HH:MM ampm")
       If IsDate(MyDateString) Then
       .TextMatrix(.Row, .Col) = MyDateString
       Else
        See_Msg "Invalid Selection", 5, ""
       End If
     End Select
     .Col = 0
   End With
End Sub

Private Sub BtnEnh1_Click()
If Emp_Validate.Visible Then Exit Sub
Call MemoryCheck(Time_Sheet)
End Sub

Private Sub PrintPunches()
 Dim ff          As Integer
 Dim PS          As Integer
 Dim PrinterPort As String
 Dim AllData$
 ff = FreeFile
 PrinterPort = GetSpoolPrinter(AllPrinter.Item("Receipt").PrinterName)
 CenterText = EC(8)
 LeftText = EC(9)
 EC(7) = vbCrLf
 AllData$ = EC(0)  'Int printer
 AllData$ = AllData$ & EC(1) & EC(3) & EC(8) & EMPName.Caption & EC(2) & EC(7)
 AllData$ = AllData$ & EC(1) & EC(3) & EC(8) & String(40, 45) & EC(2) & EC(7)
 For PS = 1 To PunchList.ListItems.count
  AllData$ = AllData$ & EC(9) & EC(2) & "Date " & PunchList.ListItems(PS).SubItems(1) & EC(2) & EC(7)
  If PunchList.ListItems(PS).SubItems(3) > "" Then
  AllData$ = AllData$ & EC(9) & EC(2) & "Break Out (" & PunchList.ListItems(PS).SubItems(3) & ") Break In (" & PunchList.ListItems(PS).SubItems(5) & ")" & EC(2) & EC(7)
  End If
  AllData$ = AllData$ & EC(9) & EC(2) & "In (" & PunchList.ListItems(PS).SubItems(2) & ") Out (" & PunchList.ListItems(PS).SubItems(5) & ") Hours " & PunchList.ListItems(PS).SubItems(6) & EC(2) & EC(7)
  AllData$ = AllData$ & EC(1) & EC(3) & String(40, 45) & EC(2) & EC(7)
 Next PS
 AllData$ = AllData$ & EC(9) & EC(1) & "Time Total " & TTime.Caption & EC(2) & EC(7)
 AllData$ = AllData$ & EC(6) & EC(5)
 ReceiptChit AllData$, PrinterPort
End Sub

Private Sub BtnEnh2_Click()
If Emp_Validate.Visible Then Exit Sub
PrintPunches
End Sub

Private Sub BtnEnh3_Click()
IApprove = "Approved"
Emp_Validate.Visible = False
End Sub
Private Sub BtnEnh4_Click()
IApprove = "Not Approved"
Emp_Validate.Visible = False
See_Msg "An exception will be added to this records.", 0, "Please See a Manager."
End Sub
Private Sub BtnEnh5_Click()

MealBreak.Visible = False
RegularBreak.Visible = False
SaveKey.Visible = True
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Hoursworked As Single
Dim TimeWorked  As Single
With BreakEdit
     .Clear
     .cols = 5
     .FixedRows = 1
     .ColWidth(0) = 0
     .ColWidth(1) = 1200 * WidScale1024
     .ColWidth(2) = 2000 * WidScale1024
     .ColWidth(3) = 2000 * WidScale1024
     .ColWidth(4) = 1200 * WidScale1024
     .width = 6750 * WidScale1024
     .Row = 0
     ' recordid is stored in this col
     .Col = 1
     .Text = "Break Name"
     .Col = 2
     .Text = "Time Out"
     .Col = 3
     .Text = "Time In"
     .Col = 4
     .Text = "Break Time"
End With
BreakView.Visible = False
BreakEdit.Visible = True
BreakEdit.ZOrder 0
BreakSheet.Visible = True
BreakSheet.ZOrder 0
CenterObject Me, BreakSheet
Dim R As Integer
Dim T As Integer
ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
R = 0
For T = 1 To TimeSheet.rows
   If TimeSheet.TextMatrix(T, 1) = "" Then Exit For
   Sql = "Select * From [Breaks] where [Name] = '" & Employees.Text & "' and [RecordId] = '" & TimeSheet.TextMatrix(T, 24) & "'"
   Sql = Sql & " Order By TRY_CAST([TimeOut] AS datetime) desc"
  ConnectADORS Sql, Db, Tb
   With BreakEdit
    
    Do While Not Tb.EOF
      R = R + 1
      Hoursworked = CalculateBreakTime(Tb.fields("TimeOut").Value, Tb.fields("TimeIn").Value, BreakMins)
      TimeWorked = TimeWorked + Val(Hoursworked)
     .TextMatrix(R, 0) = Tb.fields("RecordId").Value & ""
     .TextMatrix(R, 1) = Tb![BreakType] & ""
     .TextMatrix(R, 2) = Format(Tb![TimeOut] & "", "MM/DD/YY HH:MM ampm")
     .TextMatrix(R, 3) = Format(Tb![TimeIn] & "", "MM/DD/YY HH:MM ampm")
     .TextMatrix(R, 4) = Hoursworked
     Tb.MoveNext
 Loop
 End With
 Tb.Close
Next T
DbClose Db, Tb

End Sub

Private Sub BtnEnh6_Click()
SelectAShift.Visible = False
SelectAShift.Left = -5000
End Sub

Private Sub BtnEnh7_Click()
BreakSheet.Visible = False
BreakSheet.Left = -5000
End Sub

Private Sub BtnEnh8_Click()
LoadData True
End Sub

Private Sub Button_1_Click()
On Error GoTo ErrorHandler
Dim CanClockIN              As Boolean
Dim GetEmpMsg               As String
Dim Job                     As String
Dim SECRETWORD              As String
Dim TipsPaidToOthers        As Currency
Dim ShiftDate               As Date
Dim mPrice                  As String
Dim CanClockOUT             As Boolean
Dim GetEmpMsgOUT            As String
Dim Audit                   As AuditData
Dim PDate                   As String '// this value is used for quarter hour rounding
Dim QtrMins                 As Long

If ForGot = 1 Then
   See_Msg "You Forgot to Punch Out", 5, "Have a manager correct your shift."
   Exit_Click
   Exit Sub
End If

Dim PrinterPort     As String
Dim PF              As Integer
Dim Wwtd            As String
Dim HRS             As String
Dim Total_Time      As Single
Dim TTLWorked       As Single
Dim Hours_Worked    As Single
Dim TTlMins         As Single
Dim Mins_Worked     As Single
Dim StartShiftDate  As String
Dim ShiftNo         As Long
Dim Hoursworked     As String
Dim Response        As String
Dim WhosHere        As EmployeeData
Dim Empdb           As ADODB.Connection
Dim EmpTb           As ADODB.Recordset
Dim currentEmployee As String

currentEmployee = Emp_Name.Caption

If Button_1.Caption = "Punch Out" Then
  
  If Check_Option(46) = 1 Then
    '// requires an audit to be closed to punch out
    '** added 1-7-99 ledgends steak house, Must Close Audit to punch out
    Dim PO       As ADODB.Connection
    Dim POrs     As ADODB.Recordset
    
    ConnectADODB SvrPath.TerminalPath & "MSaleData", PO
      Sql = "Select Top 1 [Servers Name] from [Check Numbers] Where [Servers Name] = '" & currentEmployee & "'"
      Sql = Sql & " and [audit status] = 0 "
    ConnectADORS Sql, PO, POrs
        If Not POrs.EOF Then
            DbClose PO, POrs
            See_Msg "Must Close Audit", 5, "audit close required to punch out"
            Exit Sub
        End If
    DbClose PO, POrs
  End If
  
    '//option to require audit close for punch out
  If GetConfigInfo("Time Clock", "Time Window Enabled for Punch Out", "No", "Functions.INI") = "ENABLED" Then
   If PunchCheck = False Then
     If Can_Continue("Managers Menu") = True Then

     Else
        Exit Sub
     End If
   End If
  End If
  
  '// check for messagers
  Check_Mail currentEmployee
  
  Dim Db   As ADODB.Connection
  Dim OTB  As ADODB.Recordset
  
  ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
   
  If Check_Option(9) = 0 Then  '// option to disable breaks
    Sql = "Select * from [Employee Records] where [employee Name] = '" & currentEmployee & "' and [Time in] > '' and [Break Out] > '' and [Break In] = ''"
    Sql = Sql & " and TRY_CAST([Shift Date] AS date) >= '" & GetStartOfWeek & "'"
    ConnectADORS Sql, Db, OTB
    If Not OTB.EOF Then
      See_Msg "Please finish your break", 5, "cannot punch out until your break is complete."
      DbClose Db, OTB
      Exit Sub
    End If
    
    OTB.Close
    
    If BreakMins > 0 Then
     Sql = "Select * from [Breaks] where [Name] = '" & currentEmployee & "' and [TimeOut] > '' and [TimeIn] = ''"
     ConnectADORS Sql, Db, OTB
      If Not OTB.EOF Then
       See_Msg "Please complete your break", 5, ""
       DbClose Db, OTB
       Exit Sub
      End If
    End If
  End If
   
   
   Sql = "Select * from [Employee Records] where [employee Name] = '" & currentEmployee & "' and [Time in] > '' and [Time Out] = ''"
  ConnectADORS Sql, Db, OTB
   If OTB.EOF Then
     See_Msg "You have no open shift", 5, "Open shift not found."
     DbClose Db, OTB
     Exit Sub
   End If
   Dim CurrentJobCode As String
   CurrentJobCode = OTB.fields("Department Name").Value
   WhosHere = GetEmployee(CurrentJobCode)
   
   If SevenShiftsIsOn Then
     Dim Reply() As String
     Dim RPunch As SevenShift
     Set RPunch = New SevenShift
       Response = RPunch.PunchOut(Nz(OTB.fields("SSN").Value), Now)   'department is jobcode
     Set RPunch = Nothing
     Reply = Split(Response, "|")
     If Reply(0) = "success" Then
       
     Else
       See_Msg "Error Clocking Out " & Reply(1), 0, ""
       Trans_Log "Error Clocking Out with 7 shifts " & Reply(1) & " " & Reply(2)
     End If  '// punch out for 7 shifts
   End If
    
    If Check_Tipped(CurrentJobCode, WhosHere.FirstName, WhosHere.LastName) = 1 Then
      
      Dim TMsg As String
      StartShiftDate = CStr(OTB![Shift date])
      
      If Not IsDate(StartShiftDate) Then StartShiftDate = Format(Date, "MM/DD/YYYY")
      
      Audit.ChargeTips = GetServerTips(currentEmployee, Format(StartShiftDate, "mm/dd/yyyy"), Val(OTB![Shift] & ""), OTB.fields("Time In").Value & "")
      
      If (Audit.ChargeTips) > 0 Then
        TMsg = "Cash Tips Only "
      Else
        TMsg = "Declare All Tips "
      End If
      
      If GetConfigInfo("Time Clock", "Tip Pool for Tip Declaration", "No", "Functions.INI") = "ENABLED" Then
         TMsg = "Tip Pool. Declare All Tips! "
         Audit.ChargeTips = 0
      End If
      
      If Format(Audit.ChargeTips, "0.00") > 0 Then
         See_Msg "Total Credit Card Tips " & Format(Format(Audit.ChargeTips, "0.00"), "0.00"), 0, "Declare Cash Tips Only"
      Else
         See_Msg "No Credit Card Tips Found", 0, "Please Declare All Tips"
      End If
      
      Audit.CashTips = Val(ShowNumberPad(TMsg, False, 0, 0))
      If Audit.CashTips = 0 Then
        If Can_Continue("Tip Required") = False Then
           Exit Sub
        End If
      End If
      
      If GetConfigInfo("Time Clock", "Prompt for Server Tips Paid Out(Yes/No)", "No", "Functions.INI") = "ENABLED" Then
         TipsPaidToOthers = Val(ShowNumberPad("Enter Tips Paid Out", False, 0, 0))
      Else
         TipsPaidToOthers = -1 'GetTipSharePaidOut(Punch_Name, Format(StartShiftDate, "mm/dd/yyyy"))
      End If
    End If
    '// end checking tips
    DbClose Empdb, EmpTb
    
    Hoursworked = CalculateTimeRecord(OTB![Shift date] & "", Date, OTB![Time In] & "", OTB![Break Out] & "", OTB![Break In] & "", Format(Time, "HH:MM AmPm"), Total_Time)
    If (Total_Time / 60) > 16 Then
       Call See_Msg("You forgot to punch out!", 5, "Have a manager correct your time record.")
       DbClose Db, OTB
       Exit Sub
    End If
    
    Dim PunchTimeValue As String
    'If Check_Option(41) = 1 Then  '// round to the quarter hour
    '   PunchTimeValue = QuarterRound(QtrMins, PDate)
    'Else
       PunchTimeValue = CStr(Format(Time, "HH:MM AmPm"))
    'End If
    
    If Trim$(PunchTimeValue) = "" Then PunchTimeValue = CStr(Format(Time, "HH:MM AmPm"))
    
    StartShiftDate = CStr(OTB![Shift date])
    
    OTB![Time Out] = PunchTimeValue & ""
    'If Check_Option(41) = 1 Then
    '   OTB![Punch out Date] = CStr(Format(PDate, "mm/dd/yyyy"))
    'Else
       OTB![Punch out Date] = CStr(Format(Date, "mm/dd/yyyy"))
    'End If
    If IsNull(OTB![Altered By]) Then OTB![Altered By] = ""
    If IsNull(OTB![Altered In]) Then OTB![Altered In] = 0
    If IsNull(OTB![Altered Out]) Then OTB![Altered Out] = 0
    If IsNull(OTB![Altered Brk Out]) Then OTB![Altered Brk Out] = 0
    If IsNull(OTB![Altered Brk In]) Then OTB![Altered Brk In] = 0
    
    OTB![Tip Amount] = Format(Audit.CashTips, "0.00")
    'If SevenShiftsIsOn Then
    
    'Else
     'OTB![SSN] = GetSSN(Punch_Name)
    'End If
    
    ShiftNo = Val(OTB![Shift] & "")
    
    If IsNull(OTB![Tip Share]) Then OTB![Tip Share] = 0
    
    If ShiftNo > 0 Then
       OTB![Sales] = Audit.TipSales
       OTB![Credit Card Tips] = Format(Audit.ChargeTips, "0.00")
       RecordsTips Punch_Name, CStr(OTB![Shift date]), TipsPaidToOthers, -1, ShiftNo, True, Audit
    Else
       OTB![Sales] = 0
       OTB![Credit Card Tips] = 0
    End If
    If TipsPaidToOthers > -1 Then
       OTB![Tips Paid Out] = TipsPaidToOthers
    Else
       OTB![Tips Paid Out] = 0
    End If
    
    CenterObject Me, Emp_Validate
    
    Emp_Validate.Visible = True
    If OTB![Altered In] = 1 Then
       STM.ForeColor = vbRed
    Else
       STM.ForeColor = vbYellow
    End If
    If OTB![Altered Out] = 1 Then
       ETO.ForeColor = vbRed
    Else
       ETO.ForeColor = vbYellow
    End If
    If OTB![Altered Brk Out] = 1 Then
       BTO.ForeColor = vbRed
    Else
       BTO.ForeColor = vbYellow
    End If
    If OTB![Altered Brk In] = 1 Then
       BTI.ForeColor = vbRed
    Else
       BTI.ForeColor = vbYellow
    End If
    STM.Caption = OTB.fields("Time In").Value
    If Check_Option(9) = 0 Then
       If OTB.fields("Break Out").Value > "" Then
          BTO.Caption = Format(OTB.fields("Break Out").Value, "HH:MM ampm")
          BTI.Caption = Format(OTB.fields("Break In").Value, "HH:MM ampm")
       Else
          BTO.Caption = ""
          BTI.Caption = ""
       End If
    Else
       Label3.Caption = "": Label5.Caption = "": BTO.Caption = "": BTI.Caption = ""
    End If
    ETO.Caption = PunchTimeValue
    TWT.Caption = Hoursworked
    If OTB![Altered By] > "" Then
       ShiftInfo.ForeColor = vbRed
       ShiftInfo.Caption = "Altered Shift"
    Else
       ShiftInfo.ForeColor = vbYellow
       ShiftInfo.Caption = "Shift Record"
    End If
    Emp_Validate.ZOrder 0
    Do While Emp_Validate.Visible
       DoEvents
       Sleep 1000
    Loop
    OTB![Altered By] = IApprove & ""
    OTB.Update
    
    If Check_Option(18) = 1 Then
       Printer_Drivers "Receipt"
       PF = FreeFile
       PrinterPort = "Time Sheet"
OneMoreTime:
       Open PrinterPort For Output Shared As #PF
         Print #PF, EC(0) 'Int printer
         Print #PF, EC(8); EC(1); EC(3); "Time Slip"; EC(2); EC(7)
         Print #PF, EC(7)
         Print #PF, EC(1); EC(3); EC(8); Punch_Name; EC(2); EC(7)
         Print #PF, EC(9); String(33, 45); EC(2); ; EC(7)
         Print #PF, EC(9); "Department Name : "; OTB(1); EC(2); EC(7)
         Print #PF, String(33, 45); EC(2); ; EC(7)
         Print #PF, EC(9); EC(1); EC(3); "Shift Date "; OTB(2); EC(7)
         Print #PF, EC(9); EC(1); EC(3); "Time On    "; Format(OTB(3), "HH:MM:AMPM"); EC(7)
         Print #PF, EC(9); EC(1); EC(3); "Time Off   "; Format(OTB(6), "HH:MM:AMPM"); EC(7)
         Hoursworked = CalculateTimeRecord(OTB![Shift date] & "", _
                       OTB![Punch out Date] & "", OTB![Time In] & "", OTB![Break Out] & "", OTB![Break In] & "", OTB![Time Out] & "", Total_Time)
         Print #PF, EC(9); EC(1); EC(3); "Shift : "; Hoursworked
         Total_Time = 0: Hoursworked = 0
         
         SetNothing OTB
         
         Sql = "Select * from [Employee Records] where [employee Name] = '" & Punch_Name & "'"
         Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) between '" & CDate(GetCurrentPayrollWeek) & "' and '" & Date & "'"
         ConnectADORS Sql, Db, OTB
         
         TTLWorked = 0
           Do While Not OTB.EOF
            Hoursworked = CalculateTimeRecord(OTB![Shift date], _
                          OTB![Punch out Date], OTB![Time In] & "", OTB![Break Out] & "", OTB![Break In] & "", OTB![Time Out] & "", Total_Time)
            TTLWorked = TTLWorked + Total_Time
            OTB.MoveNext
         Loop
         If TTLWorked > 0 Then
            Wwtd = "Total Time : " & Format(TTLWorked / 60, "0.00")
         Else
            Wwtd = "Total Time : 0.00"
         End If
         Print #PF, EC(9); EC(1); EC(3); Wwtd
         
         If Audit.CashTips > 0 Then Print #PF, "Cash Tips Reported " & Format(Audit.CashTips, "0.00")
         
         Print #PF, String(33, 45); EC(7)
         If Audit.ChargeTips > 0 Then
            Print #PF, "Non Cash Tips " & Format(Audit.ChargeTips, "0.00")
            Print #PF, "Total Tips    " & Format(Audit.CashTips + Audit.ChargeTips, "0.00")
         End If
         If TipsPaidToOthers > 0 Then
            Print #PF, "Tips Paid Out      " & Format(TipsPaidToOthers, "0.00")
            Print #PF, "Net Tips           " & Format((Audit.CashTips + Audit.ChargeTips) - TipsPaidToOthers, "0.00")
         End If
         Print #PF, "I Have " & IApprove & " This Shift"
         Print #PF, EC(6) & EC(5)
         Close #PF
         SendJobToPrinter PrinterPort, "Receipt", False
    End If
    
    
    RemoteVideoSendItem "  | " & Punch_Name & " Has Punched Out"
    Trans_Log Punch_Name & " Shift Out " & Now
    
    
    See_Msg "Recorded....", 5, PunchTimeValue
    
    DbClose Db, OTB
    
    If Check_Option(74) = 1 Then
      If GetShiftNo(Punch_Name) > 0 Then
        If CloseServerAudit(Punch_Name) Then
           See_Msg "Audit Closed...", 5, ""
        End If
      End If
    End If
    Call MemoryCheck(Time_Sheet)
    Exit Sub
Else
    If Punch_Name = "" Then Exit Sub
    
    If PunchCheck = False Then
        If Can_Continue("Managers Menu") = True Then

        Else
          Exit Sub
        End If
    End If
    
    
    If Work_Dept(0).ListCount > 1 Then
       CenterObject Me, Work_Area
       Work_Area.Visible = True
       Work_Area.ZOrder 0
       Work_Area.Refresh
    Else
       If Work_Dept(0).List(0) = "" Then
          See_Msg "No Job Code", 5, ""
          Exit Sub
       End If
       Emp_Punch_In Work_Dept(0).List(0), Work_Dept(1).List(0), Work_Dept(2).List(0)
       Call MemoryCheck(Time_Sheet)
       Exit Sub
    End If
End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Sheet Button 1 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub Button_2_Click()
On Error GoTo ErrorHandler
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Hoursworked As Single
Dim TimeWorked  As Single

If Button_1.Caption = "Punch In" Then
 See_Msg "Cannot take a break without a valid punch in records", 5, ""
 Exit Sub
End If

Dim RR As ReturnTimeValues
RR = GetTimeRecordId(Punch_Name)
If RR.Results = "Success" Then
  RecordId = RR.RecordId
Else
  See_Msg RR.Results, 0, ""
  RecordId = ""
  Exit Sub
End If
BreakEdit.Visible = False
BreakView.Visible = True
BreakView.ZOrder 0
LoadBreaks
                   
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Sheet BTn2 ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Sub ChangeTime_Click()
If OkToSave = False Then Exit Sub
On Error GoTo ErrorHandler
If Can_Continue(ChangeTime.Caption) = False Then Exit Sub
Dim H    As String
Dim F    As String
Dim Data As String
Dim NxR As Integer
Dim ShiftTime As String
Dim currentEmployee As String
currentEmployee = Employees.Text
If currentEmployee = "" Then
   Call See_Msg("Must Have A Name", 5, "No Employee Name")
   Exit Sub
End If
Dim SvrSales As Currency
Dim ExpSales As Currency
FormIsDirty = False
H = "Department|Date;"
F = "+<2500|+>1000;"

If See_Msg("Are you sure ?", 4, "The Changes Will Become Permanent") <> 6 Then Exit Sub
    With TimeSheet
     For NxR = 1 To .rows
       If .TextMatrix(NxR, 1) = "" Then Exit For
       If .TextMatrix(NxR, 5) > "" Then
        If TestTimeFailed(NxR, .TextMatrix(NxR, 1) & " " & .TextMatrix(NxR, 2), .TextMatrix(NxR, 3), .TextMatrix(NxR, 4), .TextMatrix(NxR, 7) & " " & .TextMatrix(NxR, 5)) Then
           See_Msg "Invalid Time On Line " & NxR, 0, "Cannot Save Work"
           Data = .TextMatrix(NxR, 1) & "|"
           Data = Data & Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY") & ";"
           ReportView.View.TableBorder = tbBoxColumns
           ReportView.View.addTable F, H, Data, vbBlack, vbWhite
           FinishJob 0
           Exit Sub
        End If
       End If
     Next NxR
    End With
    TimeSheet.HighLight = flexHighlightNever
    Dim Db As ADODB.Connection
    Dim Tb As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
    With TimeSheet
      For NxR = 1 To .rows - 1
          If Trim(.TextMatrix(NxR, 0)) = "" Then Exit For
          If Trim(.TextMatrix(NxR, 0)) > "" And Trim(.TextMatrix(NxR, 1)) > "" And Trim(.TextMatrix(NxR, 2)) > "" Then
             Sql = "Select * From [Employee Records] where [Employee Name] = '" & currentEmployee & "'"
             Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) = '" & Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY") & "'"
             Sql = Sql & " and [Department Name] = '" & .TextMatrix(NxR, 0) & "'"
             Sql = Sql & " and [Time In] = '" & .TextMatrix(NxR, 23) & "'"
             
          ConnectADORS Sql, Db, Tb
                 If Tb.EOF Then
                  Tb.AddNew
                 Else
                  If SevenShiftsIsOn Then
                       Dim MyShift           As EmployeeData
                       Dim Response          As String
                       Dim Reply()           As String
                       Dim ShiftId           As String
                       Dim result            As String
                       Dim sTimeIn            As String
                       Dim sTimeOut           As String
                       ShiftId = Tb.fields("SSN").Value
                    If Len(ShiftId) > 0 Then
                       '// grab the shift record
                       MyShift = GetEmployee(.TextMatrix(NxR, 0))    '// looks up the job code
                    
                       sTimeIn = Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY") & " " & .TextMatrix(NxR, 2) & ""
                       sTimeOut = Format(.TextMatrix(NxR, 7) & "", "MM/DD/YYYY") & " " & .TextMatrix(NxR, 5) & ""
                       Dim RPunch As SevenShift
                       Set RPunch = New SevenShift
                       result = RPunch.DeleteTimePunch(ShiftId)

                       If Left$(result, 7) = "success" Then
                        'MsgBox ("Punch deleted")
                       Else
                        MsgBox (result)
                        Trans_Log "error removing shift " & result
                       End If
                       Response = RPunch.PunchIn(MyShift.ScheduleId, MyShift.JobRoleId, MyShift.DepartmentId, CDate(sTimeIn))
                    
                       Reply = Split(Response, "|")
                    
                       If Reply(0) = "success" Then
                          Trans_Log "Writing In Complete " & Reply(1)
                          ShiftId = Reply(1)
                       End If
                    
                       If IsDate(.TextMatrix(NxR, 5) & "") Then
                          Response = RPunch.PunchOut(ShiftId, sTimeOut)
                     
                          Reply = Split(Response, "|")
                     
                          If Reply(0) = "success" Then
                            Trans_Log "Writing Out Complete " & Reply(1)
                          End If
                       End If
                       Set RPunch = Nothing
                    End If
                 End If
              End If
              Tb![employee name] = currentEmployee
              Tb![Department Name] = .TextMatrix(NxR, 0) & ""
              Tb![Shift date] = Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY")
              Tb![Time In] = .TextMatrix(NxR, 2) & ""
              Tb![Break Out] = .TextMatrix(NxR, 3) & ""
              Tb![Break In] = .TextMatrix(NxR, 4) & ""
              Tb![Time Out] = .TextMatrix(NxR, 5) & ""
              Tb![Punch out Date] = .TextMatrix(NxR, 7) & ""
              Tb![Tip Amount] = Val(.TextMatrix(NxR, 13) & "")
              If .TextMatrix(NxR, 12) & "" > "" Then
                  Trans_Log "Time Record Edit for  " & currentEmployee & " Recorded Time " & Now
                  Trans_Log "Edit By               " & .TextMatrix(NxR, 12)
                  Trans_Log "Time In               " & Val(.TextMatrix(NxR, 8) & "")
                  Trans_Log "Time Out              " & Val(.TextMatrix(NxR, 11) & "")
                  Trans_Log "Break In              " & Val(.TextMatrix(NxR, 9) & "")
                  Trans_Log "Break Out             " & Val(.TextMatrix(NxR, 10) & "")
              End If
              Tb![Altered By] = .TextMatrix(NxR, 12) & ""
              Tb![Altered In] = Val(.TextMatrix(NxR, 8) & "")
              Tb![Altered Out] = Val(.TextMatrix(NxR, 11) & "")
              Tb![Altered Brk Out] = Val(.TextMatrix(NxR, 9) & "")
              Tb![Altered Brk In] = Val(.TextMatrix(NxR, 10) & "")
              If Val(.TextMatrix(NxR, 15) & "") = 0 Then
               Tb![Rate of Pay] = GetRateofPay(currentEmployee, .TextMatrix(NxR, 0) & "")
              Else
               Tb![Rate of Pay] = Val(.TextMatrix(NxR, 15) & "")
              End If
              Tb![Paid By] = Val(.TextMatrix(NxR, 16) & "")
              If Val(.TextMatrix(NxR, 17) & "") = 0 And StartAudit(.TextMatrix(NxR, 0) & "") = 1 Then
               Tb.fields("Sales").Value = ServerTipSales(currentEmployee, Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY"), Val(.TextMatrix(NxR, 19) & ""))
              Else
               Tb![Sales] = Val(.TextMatrix(NxR, 17) & "")
              End If
              If Val(.TextMatrix(NxR, 19) & "") = 0 And StartAudit(.TextMatrix(NxR, 0) & "") = 1 Then
                 .TextMatrix(NxR, 19) = GetAuditShift(currentEmployee, Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY"))
                  Tb![Shift] = Val(.TextMatrix(NxR, 19) & "")
              Else
                  Tb![Shift] = Val(.TextMatrix(NxR, 19) & "")
              End If
              If Val(.TextMatrix(NxR, 18) & "") = 0 And Val(.TextMatrix(NxR, 19) & "") > 0 Then
                 Tb![Credit Card Tips] = GetServerTips(currentEmployee, Format(.TextMatrix(NxR, 1) & "", "MM/DD/YYYY"), Val(.TextMatrix(NxR, 19) & ""), CStr(.TextMatrix(NxR, 2) & ""))
              Else
                 Tb![Credit Card Tips] = Val(.TextMatrix(NxR, 18) & "")
              End If
              Tb![Tip Share] = Val(.TextMatrix(NxR, 20) & "")
              Tb![Tips Paid Out] = Val(.TextMatrix(NxR, 21) & "")
              If SevenShiftsIsOn Then
               Tb![SSN] = ShiftId
              Else
               Tb![SSN] = GetSSN(currentEmployee)
              End If
             Tb.Update
          End If
       Next NxR
   End With
    'Tb.Save "C:\time.xml", adPersistXML
    Sql = "Delete From [Employee Records] where [Employee Name] = '" & currentEmployee & "'"
    Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) Between '" & GetBeginPayrollWeek(GetStartOfWeek) & "' and '" & Date & "'"
    Sql = Sql & " and [Time In] = '12:00 AM' and [Time Out] = '12:00 AM'"
    Db.Execute Sql  '// remove all cancelled time records
   DbClose Db, Tb
   LoadData False
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "TS Change Time ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Function GetRateofPay(EWName As String, Department As String) As Currency
On Error GoTo ERH
Dim FN As Integer
Dim EmpFName As String
Dim EmpLName As String
FN = InStr(EWName, " ")
EmpFName = Trim(Left(EWName, FN - 1))
EmpLName = Trim(Mid(EWName, FN + 1, Len(EWName) - FN))
Dim Empdb As ADODB.Connection
Dim EmpTb As ADODB.Recordset
   ConnectADODB "MSaleData", Empdb
      Sql = "Select * from [Job Codes] where "
      Sql = Sql & " [First Name] = '" & EmpFName & "'"
      Sql = Sql & " and [Last Name] = '" & EmpLName & "'"
      Sql = Sql & " and [Job Code] = '" & Department & "'"
   ConnectADORS Sql, Empdb, EmpTb
     If Not EmpTb.EOF Then
       GetRateofPay = Val(EmpTb![rate])
     Else
       GetRateofPay = 0
     End If
  DbClose Empdb, EmpTb
Exit Function
ERH:
GetRateofPay = 0
End Function
Private Sub DeleteTime_Click()
If Can_Continue(DeleteTime.Caption) = False Then Exit Sub
DeleteRecord = 1
End Sub
Private Sub Employees_Click()
Punch_Name = Employees.List(Employees.ListIndex)
Employees.Text = Punch_Name
LoadData False
End Sub
Private Sub Exit_Click()
If FormIsDirty Then
  If See_Msg("Save Changes ? ", 4, "") = 6 Then
    ChangeTime_Click
  End If
End If
Edit_Time = False
Call MemoryCheck(Time_Sheet)
End Sub
Private Sub LoadData(ByVal ViewPeriod As Boolean)
Dim Db           As ADODB.Connection
Dim OTB          As ADODB.Recordset
Dim Total_Time   As Single
Dim TimeWorked   As Single
Dim FN           As Integer
Dim First_Name   As String
Dim Last_Name    As String
Dim Tb           As ADODB.Recordset
Dim CA           As Integer
Dim CTRW         As Integer
NoJobCode.Visible = False
On Error GoTo ErrorHandler
CloseForm = 0: ForGot = 0
If Edit_Time = False Then
    TimeEditor.Visible = False
    ShiftPunch.Visible = True
    CenterObject Me, ShiftPunch
    If ViewPeriod Then
      EditDate = GetBeginPayrollWeek(GetStartOfWeek)
    Else
      EditDate = GetStartOfWeek
    End If
    EMPName.Caption = Punch_Name & "  ** Week Starting " & Format(EditDate, "MM/DD/YYYY")
    OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
    Sql = "Select * from [Employee Records] where [Employee Name] = '" & Punch_Name & "'"
    Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) >= '" & EditDate & "'"
    
    Sql = Sql & " Or [Employee Name] = '" & Punch_Name & "' and [Shift Date] > '' and [Time In] > '' and [Time Out] = ''"
    If LoadOldToNew.Value = 1 Then
     Sql = Sql & " Order By TRY_CAST([Shift Date] AS datetime), TRY_CAST([Time In] AS datetime) desc"
    Else
     Sql = Sql & " Order By TRY_CAST([Shift Date] AS datetime), TRY_CAST([Time In] AS datetime) asc"
    End If
    ConnectADORS Sql, Db, OTB
    Button_1.Caption = "Punch In"
    Button_1.Enabled = True
    Button_2.Visible = False
    Button_1.ZOrder 0
    Button_2.Caption = "Break"
    Button_2.Enabled = True
    Do While Not OTB.EOF
       If OTB![Time In] & "" > "" And OTB![Time Out] & "" = "" Then
          MealBreak.Visible = True
          If BreakMins > 0 Then
            RegularBreak.Visible = True
          Else
            RegularBreak.Visible = False
          End If
          Button_1.Caption = "Punch Out"
       End If
    OTB.MoveNext
    Loop
    DbClose Db, OTB
    Dim i As Integer
    For i = PunchList.ListItems.count To 1 Step -1
     PunchList.ListItems.remove i
    Next i

    PunchList.ColumnHeaders.add 1, , "Department", 1000 * WidScale1024
    PunchList.ColumnHeaders.add 2, , "Date", 1000 * WidScale1024
    PunchList.ColumnHeaders.add 3, , "Time On", 1000 * WidScale1024
    If Check_Option(9) = 1 Then
      PunchList.ColumnHeaders.add 4, , "", 0
      PunchList.ColumnHeaders.add 5, , "", 0
    Else
      PunchList.ColumnHeaders.add 4, , "Break Out", 1000 * WidScale1024
      PunchList.ColumnHeaders.add 5, , "Break In", 1000 * WidScale1024
    End If
    PunchList.ColumnHeaders.add 6, , "Time Out", 1000 * WidScale1024
    PunchList.ColumnHeaders.add 7, , "Hours", 1000 * WidScale1024
    PunchList.ColumnHeaders.add 8, , "RecordId", 0
    If Check_Option(9) = 1 Then
      PunchList.width = 6450 * WidScale1024
    Else
      PunchList.width = 9050 * WidScale1024
    End If
Else
    EditDate = GetBeginPayrollWeek(GetStartOfWeek)
    WeekOf.Caption = "Week Starting " & Format(EditDate, "MM/DD/YYYY") & " and Ending " & Format(Date, "MM/DD/YYYY")
    ShiftPunch.Visible = False
    TimeEditor.Visible = True
    TimeSheet.Clear
    With TimeSheet
        .cols = 25
        .ColWidth(0) = 3000 * WidScale1024
        .ColWidth(1) = 1500 * WidScale1024
        .ColWidth(2) = 1500 * WidScale1024
        .ColWidth(3) = 1500 * WidScale1024
        .ColWidth(4) = 1500 * WidScale1024
        .ColWidth(5) = 1500 * WidScale1024
        .ColWidth(6) = 2000 * WidScale1024
        .width = 12550 * WidScale1024
        For CA = 7 To 24
        .ColWidth(CA) = 0
        Next CA
        
        For CA = 0 To 6
        .ColAlignment(CA) = 1
        Next CA
        .Row = 0
        .Col = 0
        .Text = "Job Name"
        .Col = 1
        .Text = "Shift Date"
        .Col = 2
        .Text = "Time On"
        .Col = 3
        .Text = "Break Out"
        .Col = 4
        .Text = "Break In"
        .Col = 5
        .Text = "Time Out"
        .Col = 6
        .Text = "Worked"
    End With
End If
ForGot = 0: TimeWorked = 0
Emp_Name.Caption = Punch_Name
Dim Hoursworked As String
Dim TTl_Hrs     As Single
Dim PlistItem As ListItem
If Punch_Name = "" Then Exit Sub
OpenDb SvrPath.TerminalPath & TimeRecordsDb, Db
 Sql = "Select * From [Employee Records] where [Employee Name] = '" & Punch_Name & "'"
 Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) >= '" & EditDate & "'"
 Sql = Sql & " Or [Employee Name] = '" & Punch_Name & "' and [Shift Date] > '' and [Time In] > '' and [Time Out] = ''"
 If LoadOldToNew.Value = 1 Then
  ' Sql = Sql & " Order By TRY_CAST([Shift Date] AS datetime), TRY_CAST([Time In] AS datetime) desc"
 Else
   Sql = Sql & " Order By TRY_CAST([Shift Date] AS datetime), TRY_CAST([Time In] AS datetime) asc"
 End If
ConnectADORS Sql, Db, Tb
     With TimeSheet
            Do While Not Tb.EOF
              Hoursworked = CalculateTimeRecord(Tb![Shift date] & "", Tb![Punch out Date] & "", _
                             Tb![Time In] & "", Tb![Break Out] & "", Tb![Break In] & "", Tb![Time Out] & "", Total_Time)
              TimeWorked = TimeWorked + Val(Hoursworked)
             If Edit_Time = False Then
                Set PlistItem = PunchList.ListItems.add(, , Tb![Department Name] & "")  ' (" & Tb![Altered By] & ")")
                   PlistItem.SubItems(1) = Tb![Shift date] & ""
                   If IsDate(Tb![Time In] & "") Then
                     PlistItem.SubItems(2) = Format(Tb![Time In] & "", "hh:mm Ampm")
                   Else
                     PlistItem.SubItems(2) = ""
                   End If
                   If IsDate(Tb![Break Out] & "") Then
                     PlistItem.SubItems(3) = Format(Tb![Break Out] & "", "hh:mm Ampm")
                   Else
                     PlistItem.SubItems(3) = ""
                   End If
                   If IsDate(Tb![Break In] & "") Then
                     PlistItem.SubItems(4) = Format(Tb![Break In] & "", "hh:mm Ampm")
                   Else
                     PlistItem.SubItems(4) = ""
                   End If
                   If IsDate(Tb![Time Out] & "") Then
                     PlistItem.SubItems(5) = Format(Tb![Time Out] & "", "hh:mm Ampm")
                   Else
                     PlistItem.SubItems(5) = ""
                     'Button_1.Caption = "Punch Out"
                   End If
                   If Tb![Altered By] > "" And Tb![Altered By] <> "Approved" Then
                    PlistItem.SubItems(6) = Hoursworked & " **"
                   Else
                    PlistItem.SubItems(6) = Hoursworked
                   End If
                   PlistItem.SubItems(7) = Tb.fields("RecordId").Value & ""
             Else
                .Row = .Row + 1
                .TextMatrix(.Row, 0) = Tb![Department Name] & ""
                .TextMatrix(.Row, 1) = Tb![Shift date] & ""
               If IsDate(Tb![Time In] & "") Then
                .TextMatrix(.Row, 2) = Format(Tb![Time In] & "", "hh:mm Ampm")
               Else
                .TextMatrix(.Row, 2) = ""
               End If
               If IsDate(Tb![Break Out] & "") Then
                .TextMatrix(.Row, 3) = Format(Tb![Break Out] & "", "hh:mm Ampm")
               Else
                .TextMatrix(.Row, 3) = ""
               End If
               If IsDate(Tb![Break In] & "") Then
                .TextMatrix(.Row, 4) = Format(Tb![Break In] & "", "hh:mm Ampm")
               Else
                .TextMatrix(.Row, 4) = ""
               End If
               If IsDate(Tb![Time Out] & "") Then
                .TextMatrix(.Row, 5) = Format(Tb![Time Out] & "", "hh:mm Ampm")
               Else
                .TextMatrix(.Row, 5) = ""
               End If
               .TextMatrix(.Row, 7) = Tb![Punch out Date] & ""
               
               If (Total_Time / 60) >= 24 And Trim(Tb![Time Out]) = "" Then ForGot = 1
                If Trim(Tb![Altered By]) > "" And Trim(Tb![Altered By]) <> "Approved" Then
                 .CellForeColor = vbRed
                Else
                 .CellForeColor = vbBlue
                End If
               
               If (TTl_Hrs / 60) > 40 Then
                .CellForeColor = vbGreen
               Else
                 .CellForeColor = vbBlue
               End If
               .TextMatrix(.Row, 6) = Hoursworked
               
               .TextMatrix(.Row, 8) = Val(Tb![Altered In] & "")
               If Val(.TextMatrix(.Row, 8)) = 1 Then .Cell(flexcpForeColor, .Row, 2, .Row, 2) = vbRed
               
               .TextMatrix(.Row, 9) = Val(Tb![Altered Out] & "")
               If Val(.TextMatrix(.Row, 9)) = 1 Then .Cell(flexcpForeColor, .Row, 5, .Row, 5) = vbRed
               
               .TextMatrix(.Row, 10) = Val(Tb![Altered Brk Out] & "")
               If Val(.TextMatrix(.Row, 10)) = 1 Then .Cell(flexcpForeColor, .Row, 3, .Row, 3) = vbRed
               
               .TextMatrix(.Row, 11) = Val(Tb![Altered Brk In] & "")
               If Val(.TextMatrix(.Row, 11)) = 1 Then .Cell(flexcpForeColor, .Row, 4, .Row, 4) = vbRed
               
               .TextMatrix(.Row, 12) = Tb![Altered By] & ""
               If .TextMatrix(.Row, 12) > "" And .TextMatrix(.Row, 12) <> "Approved" Then .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbRed
               .TextMatrix(.Row, 13) = Format(Val(Tb![Tip Amount] & ""), "0.00")
               .TextMatrix(.Row, 15) = Val(Tb![Rate of Pay] & "")
               .TextMatrix(.Row, 16) = Val(Tb![Paid By] & "")
               .TextMatrix(.Row, 17) = Val(Tb![Sales] & "")
               .TextMatrix(.Row, 18) = Val(Tb![Credit Card Tips] & "")
               .TextMatrix(.Row, 19) = Val(Tb![Shift] & "")
               .TextMatrix(.Row, 20) = Val(Tb![Tip Share] & "")
               .TextMatrix(.Row, 21) = Val(Tb![Tips Paid Out] & "")
               .TextMatrix(.Row, 22) = Tb![SSN] & ""
               .TextMatrix(.Row, 23) = Format(Tb![Time In] & "", "hh:mm Ampm")
               .TextMatrix(.Row, 24) = Tb![RecordId] & ""
           End If
          Tb.MoveNext
         Loop
   End With
   DbClose Db, Tb
   If TimeWorked > 40 Then
       Total_Hours.ForeColor = vbRed
   Else
       Total_Hours.ForeColor = vbYellow
   End If
   If TimeWorked > 0 Then
     Total_Hours.Caption = "Total Hours : " & Format(TimeWorked, "0.00")
   Else
     Total_Hours.Caption = "Total Hours : 0.00"
   End If
   TTime.Caption = Total_Hours.Caption
   TimeSheet.HighLight = flexHighlightNever
   FN = InStr(Punch_Name, " ")
   First_Name = Trim(Left(Punch_Name, FN - 1))
   Last_Name = Trim(Mid(Punch_Name, FN + 1, Len(Punch_Name) - FN))
   Work_Dept(0).Clear
   Work_Dept(1).Clear
   Work_Dept(2).Clear
   OpenDb "MSaleData", Db
     Sql = "Select * from [job Codes] where "
     Sql = Sql & " [First Name] = """ & First_Name & """"
     Sql = Sql & " and [Last Name] = """ & Last_Name & """"
     Sql = Sql & " Order by [Pos] asc"
   ConnectADORS Sql, Db, Tb
    If Tb.EOF Then
      See_Msg "No Job Code Assigned", 5, "Please assign Job Code."
      NoJobCode.Visible = True
      DbClose Db, Tb
      TimeSheet.Clear
      PunchList.ListItems.Clear
      Exit Sub
    Else
     Do While Not Tb.EOF
      If UCase(Tb![Job Code] & "") <> "NONE" Then
       If IsThere(Work_Dept(0), Tb![Job Code] & "") = False Then
        Work_Dept(0).AddItem Tb![Job Code] & ""
        Work_Dept(1).AddItem Format(Tb![rate] & "", "0.00"), Work_Dept(0).NewIndex
        Work_Dept(2).AddItem Val(Tb![Hourly] & ""), Work_Dept(0).NewIndex
       End If
      End If
      Tb.MoveNext
     Loop
    End If
   DbClose Db, Tb
   Dim RR As ReturnTimeValues
   RR = GetTimeRecordId(Punch_Name)
   If RR.Results = "Success" Then
      RecordId = RR.RecordId
   Else
      RecordId = ""
      'Exit Sub
   End If
   LoadBreaks
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Sheet Load Form ")
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
FormIsDirty = False
If Check_Option(9) = 1 Then
  ShiftPunch.Height = 6300
  CenterObject Me, ShiftPunch
Else
   ShiftPunch.Height = 8000
  CenterObject Me, ShiftPunch
End If
LoadData False
Dim EditSecurity    As String
EditSecurity = UCase(GetConfigInfo("Functions", "Time Edit Security", "DISABLED", "FUNCTIONS.INI"))
On Error GoTo ErrorHandler
If Edit_Time = True Then
  Dim Db           As ADODB.Connection
  Dim OTB          As ADODB.Recordset
  ForGot = 0
  Employees.Clear
  Dim FieldName As String
  If GetConfigInfo("Functions", "Time Records Edit by First Name", "", "FUNCTIONS.INI") = "ENABLED" Then
   FieldName = "[First Name]"
  Else
   FieldName = "[Last Name]"
  End If
  ConnectADODB "MSaleData", Db
   Sql = "Select * from [General Data] Where ISNULL([Terminated], 0) = 0 "
   Sql = Sql & " order by " & FieldName & " asc"
  ConnectADORS Sql, Db, OTB
    Do While Not OTB.EOF
      '// Security
      If EditSecurity = "ENABLED" Then
       If Access_Required(OTB.fields("Access Level").Value) & "" < LoggedOnAs(0).EmpAccessLevel Then
         Employees.AddItem Trim$(OTB![First Name] & "") & " " & Trim$(OTB![Last Name] & "")
       End If
      ElseIf EditSecurity = "DISABLED" Then
        Employees.AddItem Trim$(OTB![First Name] & "") & " " & Trim$(OTB![Last Name] & "")
      Else
        If Access_Required(OTB.fields("Access Level").Value) & "" < LoggedOnAs(0).EmpAccessLevel Then
         Employees.AddItem Trim$(OTB![First Name] & "") & " " & Trim$(OTB![Last Name] & "")
        ElseIf Trim$(OTB![First Name] & "") & " " & Trim$(OTB![Last Name] & "") = LoggedOnAs(0).EMPName Then
         Employees.AddItem Trim$(OTB![First Name] & "") & " " & Trim$(OTB![Last Name] & "")
        End If
      End If
      '// end security
     OTB.MoveNext
    Loop
    DbClose Db, OTB
    CenterObject Me, TimeEditor
End If
If CPath.TerminalFailed Then
  See_Msg "Common Terminal Not Present", 5, "Cannot Connect"
End If
Employees.Text = Punch_Name
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Time Sheet Load Form ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Function TestTimeFailed(EditRw As Integer, TimeIn As String, BkInTime As String, BkOutTime As String, TimeOut As String) As Boolean
Dim LT As Integer
TestTimeFailed = False
For LT = 1 To TimeSheet.rows
 If TimeSheet.TextMatrix(LT, 1) = "" Then Exit Function
 If EditRw <> LT Then
    With TimeSheet
       If IsDate(.TextMatrix(LT, 1)) And IsDate(.TextMatrix(LT, 2)) And IsDate(.TextMatrix(LT, 7)) And IsDate(.TextMatrix(LT, 5)) Then
         If IsDate(TimeIn) And IsDate(TimeOut) Then
           If CDate(Format(TimeIn, "MM/DD/YYYY")) = CDate(.TextMatrix(LT, 1)) Then
              'if between
              If IsBetween(CDate(TimeIn), CDate(TimeOut), CDate(.TextMatrix(LT, 1) & " " & .TextMatrix(LT, 2)), CDate(.TextMatrix(LT, 7) & " " & .TextMatrix(LT, 5))) Then
                  TestTimeFailed = True
                  Exit Function
              End If
           End If
        End If
       End If
     End With
 End If
Next LT
End Function
Private Sub TakeABreak(BreakType As String, RecordId As String, RequiredTimeOut As Integer)
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Br As ADODB.Recordset
Dim Hoursworked As Single
Dim TimeWorked  As Single
Dim SqlString   As String
If RecordId = "" Then Exit Sub

ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
  If BreakType = "Meal Break" Then
     Sql = "Select * from [Breaks] where [RecordId] = '" & RecordId & "'"
     Sql = Sql & " Order by TRY_CAST([TimeOut] AS datetime) asc"
     ConnectADORS Sql, Db, Br
     If Not Br.EOF Then
       If Br.fields("TimeOut").Value & "" > "" And Br.fields("TimeIn").Value & "" = "" Then
        See_Msg "Must Complete Break before you can leave for a meal", 5, "Complete break required"
        DbClose Db, Br
        Exit Sub
       End If
     End If
    SqlString = "Select * from [Employee Records] Where [RecordId] = '" & RecordId & "'"
    SqlString = SqlString & " and [Time Out] = ''"
    ConnectADORS SqlString, Db, Tb
    If Not Tb.EOF Then
      If IsDate(Tb.fields("Break Out").Value) And IsDate(Tb.fields("Break IN").Value) Then
       See_Msg "You already have a meal break for this shift", 5, "Only 1 Meal Break Per Shift"
       DbClose Db, Tb
       Exit Sub
      End If

      If Not IsDate(Tb.fields("Break Out").Value) Then
       Tb.fields("Break Out").Value = Format(Now, "MM/DD/YYYY HH:MM ampm")
       Tb.Update
      ElseIf Not IsDate(Tb.fields("Break In").Value) Then
       Tb.fields("Break In").Value = Format(Now, "MM/DD/YYYY HH:MM ampm")
       Tb.Update
      End If
    End If
    DbClose Db, Tb
    Exit Sub
  End If
 '// end meal break
 
  Sql = "Select * from [Breaks] where [RecordId] = '" & RecordId & "'"
  Sql = Sql & " Order by TRY_CAST([TimeOut] AS datetime) asc"
ConnectADORS Sql, Db, Tb
   If Tb.EOF Then
       Tb.AddNew
       Tb![BreakType] = BreakType & ""
       Tb.fields("RecordId").Value = RecordId
       Tb![TimeOut] = Now
       Tb![TimeIn] = ""
       Tb.fields("Required").Value = RequiredTimeOut
       Tb.fields("Name").Value = Punch_Name
       Trans_Log Punch_Name & " On Break " & Now
      Tb.Update
      DbClose Db, Tb
      See_Msg "Recorded....", 5, Format(Time, "HH:MM ampm")
      LoadBreaks
      Exit Sub
   End If
   Do While Not Tb.EOF
      If IsDate(Tb![TimeOut]) And IsDate(Tb![TimeIn]) = False Then
        Hoursworked = CalculateBreakTime(Tb![TimeOut], Now, BreakMins)
        If Tb.fields("Required").Value > 0 And Hoursworked < Tb.fields("Required").Value Then
          See_Msg "Break is short you have ", 5, Format(Tb.fields("Required").Value - Hoursworked, "0.00") & " To Go"
        Else
         Tb![TimeIn] = Now
         Tb.Update
         DbClose Db, Tb
         See_Msg "Recorded....", 5, Format(Time, "HH:MM ampm")
        End If
        LoadBreaks
        Exit Sub
      End If
     Tb.MoveNext
   Loop
   Tb.AddNew
   Tb![BreakType] = BreakType & ""
   Tb.fields("RecordId").Value = RecordId
   Tb![TimeOut] = Now
   Tb![TimeIn] = ""
   Tb.fields("Required").Value = RequiredTimeOut
   Tb.fields("Name").Value = Punch_Name
   Trans_Log Punch_Name & " On Break " & Now
   Tb.Update
   See_Msg "Recorded....", 5, Format(Time, "HH:MM ampm")
   DbClose Db, Tb
LoadBreaks
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Punch In ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select

End Sub
Private Sub LoadBreaks()
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
Dim Hoursworked As Single
Dim TimeWorked  As Single
Dim BlistItem As ListItem
Dim i As Integer
BreakView.ListItems.Clear
For i = BreakView.ListItems.count To 1 Step -1
    BreakView.ListItems.remove i
Next i
If BreakView.ColumnHeaders.count = 0 Then
 BreakView.ColumnHeaders.add 1, , "Break Type", 1000 * WidScale1024
 BreakView.ColumnHeaders.add 2, , "Break Out", 1500 * WidScale1024
 BreakView.ColumnHeaders.add 3, , "Break In", 1500 * WidScale1024
 BreakView.ColumnHeaders.add 4, , "Time Out Min", 1100 * WidScale1024
End If
If BreakMins > 0 Then '// Input_Option(4) paid Regular break minutes
 RegularBreak.Caption = BreakMins & " Mins Break"
 RegularBreak.Visible = True
Else
 RegularBreak.Visible = False
 BreakView.Visible = False
End If
If MealBreakMins > 0 Then   '// Input_Option(7) paid meal break minutes
 MealBreak.Caption = MealBreakMins & " Mins Break"
 'MealBreak.Visible = True
Else
 'MealBreak.Visible = False
End If
BreakView.width = 6035 * WidScale1024
If RecordId = "" Then Exit Sub
   ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
   Sql = "Select * From [Breaks] where [RecordId] = '" & RecordId & "'"
   Sql = Sql & " Order By TRY_CAST([TimeOut] AS datetime) desc"
   ConnectADORS Sql, Db, Tb
   With BreakView
    Do While Not Tb.EOF
      Hoursworked = CalculateBreakTime(Tb.fields("TimeOut").Value, Tb.fields("TimeIn").Value, BreakMins)
      TimeWorked = TimeWorked + Val(Hoursworked)
                Set BlistItem = BreakView.ListItems.add(, , Tb![BreakType] & "")
                If IsDate(Tb![TimeOut] & "") Then
                    BlistItem.SubItems(1) = Format(Tb![TimeOut] & "", "MM/DD hh:mm Ampm")
                Else
                    BlistItem.SubItems(1) = ""
                End If
                If IsDate(Tb![TimeIn] & "") Then
                   BlistItem.SubItems(2) = Format(Tb![TimeIn] & "", "MM/DD hh:mm Ampm")
                Else
                   BlistItem.SubItems(2) = ""
                End If
                BlistItem.SubItems(3) = Format(Hoursworked, "0.00")
   Tb.MoveNext
 Loop
 End With
 DbClose Db, Tb
End Sub
Private Sub LoadOldToNew_Click()
LoadData False
End Sub
Private Sub MealBreak_Click()
TakeABreak "Meal Break", RecordId, MealBreakMins
'BtnEnh7_Click
LoadData False
End Sub

Private Sub PunchList_Click()
'Dim S As String
'With PunchList
'S = .ListItems.Item(.Index).Text
'End With
End Sub

Private Sub RegularBreak_Click()
TakeABreak "10 Min", RecordId, BreakMins
'BtnEnh7_Click
End Sub
Private Sub SaveKey_Click()
Dim R As Integer
Dim Db As ADODB.Connection
Dim Tb As ADODB.Recordset
   ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
    Sql = "Delete From [Breaks] where [Name] = '" & Employees.Text & "'"
    Sql = Sql & " and  TRY_CAST([TimeOut] AS date) >= '" & CDate(EditDate) & "'"
    Db.Execute Sql
   With BreakEdit
    For R = 1 To .rows - 1
    If .TextMatrix(R, 1) = "" Then Exit For
    Sql = "Select * From [Breaks] where [Name] = '" & Employees.Text & "'"
    Sql = Sql & " and [RecordId] = '" & .TextMatrix(R, 0) & "' and TRY_CAST([TimeOut] AS datetime) = '" & CDate(.TextMatrix(R, 2)) & "'"
   ConnectADORS Sql, Db, Tb
     If Tb.EOF Then
      Tb.AddNew
      Tb.fields("RecordId").Value = .TextMatrix(R, 0)
      Tb.fields("Required").Value = 0
      Tb.fields("Name").Value = Punch_Name
      Tb.fields("BreakType").Value = .TextMatrix(R, 1) & ""
     End If
      Tb![TimeOut] = Format(.TextMatrix(R, 2) & "", "MM/DD/YYYY HH:MM ampm")
      Tb![TimeIn] = Format(.TextMatrix(R, 3) & "", "MM/DD/YYYY HH:MM ampm")
     Tb.Update
    Next R
 End With
 DbClose Db, Tb
 See_Msg "Saved...", 5, ""
 BtnEnh7_Click
End Sub

Private Sub ShiftList_Click()
 Dim R As Integer
 Dim RecordId As String
 RecordId = TimeSheet.TextMatrix(ShiftList.ListIndex, 24)
 With BreakEdit
 For R = 1 To .rows
    If .TextMatrix(R, 0) = "" Then
       .TextMatrix(R, 0) = RecordId
       .TextMatrix(R, 1) = "Write In"
       .TextMatrix(R, 2) = ""
       .TextMatrix(R, 3) = ""
       .TextMatrix(R, 4) = ""
       Exit For
    End If
  Next R
 End With
  BtnEnh6_Click
End Sub

Private Sub SSCommand1_Click()
Work_Area.Visible = False
Sleep 0
End Sub
Private Sub Qualifytime(Cr As Long)
On Error GoTo ErrorHandler
Dim Break_Ok         As Boolean
Dim DateOut          As String
Dim BrkOut           As String
Dim BrkIn            As String
Dim PlaceMentPostion As Long
Dim StartCol         As Long
Dim DateIn           As String
Dim TimeIn           As String
Dim TimeOut          As String
Dim Hoursworked      As String
Dim Total_Time       As Single
Dim NewDate          As String
PlaceMentPostion = TimeSheet.Col
With TimeSheet
   If .TextMatrix(Cr, 2) = "" Then
      Call See_Msg("Invalid In Time", 5, "Can not Save Invalid Time")
      'ChangeTime.Enabled = False
      Exit Sub
   End If
   If .TextMatrix(Cr, 3) = "" And .TextMatrix(Cr, 4) > "" Then
      Call See_Msg("Invalid Break Time", 5, "Can not Save Invalid Time")
      'ChangeTime.Enabled = False
      Exit Sub
   End If
   If .TextMatrix(Cr, 3) > "" And .TextMatrix(Cr, 4) > "" Then
     If CDate(.TextMatrix(Cr, 3)) > CDate(.TextMatrix(Cr, 4)) And CDate(.TextMatrix(Cr, 2)) = CDate(.TextMatrix(Cr, 7)) Then
      Call See_Msg("Invalid Break Time", 5, "Can not Save Invalid Time")
      'ChangeTime.Enabled = False
      Exit Sub
     End If
   End If
   If IsDate(.TextMatrix(Cr, 1)) = False Then
      Call See_Msg("Invalid Date", 5, "Can not Save Invalid Date")
      'ChangeTime.Enabled = False
      Exit Sub
   End If
   If Trim(.TextMatrix(Cr, 0)) = "" Then
      Call See_Msg("No Job Title", 5, "Can not Save With Out Job Title")
      'ChangeTime.Enabled = False
      Exit Sub
   End If
   If Trim(.TextMatrix(Cr, 5)) > "" And Trim(.TextMatrix(Cr, 2)) > "" Then
    If CDate(.TextMatrix(Cr, 2)) > CDate(.TextMatrix(Cr, 5)) Then
      'If CDate(.TextMatrix(Cr, 2) & " " & .TextMatrix(Cr, 1)) >= CDate(.TextMatrix(Cr, 7) & " " & .TextMatrix(Cr, 5)) Then
        If DateDiff("h", CDate(.TextMatrix(Cr, 2) & " " & .TextMatrix(Cr, 1)), CDate(.TextMatrix(Cr, 7) & " " & .TextMatrix(Cr, 5))) > 17 Then
          See_Msg "Invalid shift time", 0, "Number of Shift Hours " & DateDiff("h", CDate(.TextMatrix(Cr, 2) & " " & .TextMatrix(Cr, 1)), CDate(.TextMatrix(Cr, 7) & " " & .TextMatrix(Cr, 5)))
          'ChangeTime.Enabled = False
          Exit Sub
        Else
           See_Msg "Setting time out date to " & Str(CDate(.TextMatrix(Cr, 1)) + 1), 5, "Auto correcting date"
           .TextMatrix(Cr, 7) = CDate(.TextMatrix(Cr, 1)) + 1
        End If
      'End If
    End If
   End If
Select Case .Col
  Case Is = 2
      .TextMatrix(Cr, 8) = 1
   Case Is = 3
      .TextMatrix(Cr, 9) = 1
    Case Is = 4
      .TextMatrix(Cr, 10) = 1
   Case Is = 5
      .TextMatrix(Cr, 11) = 1
End Select
   
     .TextMatrix(Cr, 12) = WhoIsManager
tryagain:
With TimeSheet
    DateIn = .TextMatrix(Cr, 1)
    TimeIn = .TextMatrix(Cr, 2)
    BrkOut = .TextMatrix(Cr, 3)
    BrkIn = .TextMatrix(Cr, 4)
    TimeOut = .TextMatrix(Cr, 5)
    DateOut = .TextMatrix(Cr, 7)
End With
Hoursworked = CalculateTimeRecord(DateIn, DateOut, TimeIn, BrkOut, BrkIn, TimeOut, Total_Time)
If Total_Time < 0 Then  'if less that 0 then use indate
      Call See_Msg("Invalid Work Time", 5, "Can not Save")
      'ChangeTime.Enabled = False
      Exit Sub
ElseIf Total_Time > 1440 Then  'if over 24 hours use in date
      Call See_Msg("Invalid Work Time", 5, "Employee has worked more than 24 hours")
      'ChangeTime.Enabled = False
      Exit Sub
End If
    .TextMatrix(Cr, 6) = Hoursworked
    ChangeTime.Enabled = True
End With
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " QualifyTime")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Function GetTime(OldTime As String) As String
GetTime = OldTime
If See_Msg("Enter Time", 4, "Enter or Edit Time") <> 6 Then Exit Function
Dim InDate As String
Dim PKT As DateOrTime
Set PKT = New DateOrTime
InDate = PKT.ReturnTime
If IsDate(InDate) Then
    GetTime = Format(InDate, "hh:mm ampm")
Else
    GetTime = OldTime
End If
Set PKT = Nothing
End Function
Private Sub SSCommand2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub

Private Sub SSCommand3_Click()
DeleteRecord = 0
'EraseRcd = True
End Sub
Private Sub SSCommand4_Click()
End Sub

Private Sub SSCommand7_Click()

End Sub

Private Sub SSCommand6_Click()

End Sub

Private Sub SSCommand5_Click()
If Can_Continue(SSCommand5.Caption) = False Then Exit Sub
Dim DaysToEdit As Integer
DaysToEdit = Val(Input_Option(25))
Dim AMM As MaintenanceMenus
Set AMM = New MaintenanceMenus
    AMM.TipEdit LoggedOnAs(0).EMPName, LoggedOnAs(0).EmpAccessLevel, DaysToEdit
Set AMM = Nothing
End Sub

Private Sub Timer1_Timer()
If Edit_Time = True Then Employees.Text = Punch_Name
TM.Caption = "Current Time " & Format(Time, "HH:MM ampm")
DT.Caption = "Current Date " & Format(Date, "mm/dd/yyyy")
End Sub
Private Function OkToSave() As Boolean
   Dim TL As Integer
   OkToSave = True
   With TimeSheet
   For TL = 1 To .rows - -1
      If Trim(.TextMatrix(TL, 0)) = "" Then Exit For
        If .TextMatrix(TL, 0) > "" And .TextMatrix(TL, 1) = "" Then
          'ChangeTime.Enabled = False
          Label6.Caption = "Problem with Date Field !! You Must Correct before Save"
          OkToSave = False
          Exit Function
        End If
        If .TextMatrix(TL, 1) > "" And .TextMatrix(TL, 2) = "" Then
          'ChangeTime.Enabled = False
          Label6.Caption = "Problem with Time Field !! You Must Correct before Save"
          OkToSave = False
          Exit Function
        End If
   Next TL
   End With
   ChangeTime.Enabled = True
   'Label6.Caption = "All Records Are OK!"
End Function
Private Sub TimeSheet_Click()
FormIsDirty = True
Dim InDate As String
   With TimeSheet
     If .TextMatrix(.Row, 1) > "" Then
         If LoggedOnAs(0).EmpAccessLevel < Access_Required("Current Day Time Rec") Then
              If Format(.TextMatrix(.Row, 1), "mm/dd/yyyy") <> Format(Date, "mm/dd/yyyy") Then
                See_Msg "Unable To Edit This Record", 5, "Access To Low"
                Exit Sub
              End If
        End If
     End If
   End With
  If LoggedOnAs(0).EmpAccessLevel < Access_Required("Current Day Time Rec") Then
     See_Msg "Access to Low...", 5, ""
     Exit Sub
  End If
  If Edit_Time = True And DeleteRecord = 0 Then
    If NoJobCode.Visible Then Exit Sub
    With TimeSheet
       Select Case .Col
          Case Is = 0
            If .TextMatrix(.Row, 0) = "" Then
                 If See_Msg("Add new record ?", 4, "") = 6 Then
                    If Work_Dept(0).ListCount > 1 Then
                        CenterObject Me, Work_Area
                        Work_Area.Visible = True
                        Work_Area.ZOrder 0
                        Work_Area.Refresh
                        Exit Sub
                    Else
                        .TextMatrix(.Row, 0) = Work_Dept(0).List(0)
                        .TextMatrix(.Row, 15) = Work_Dept(1).List(0)
                        .TextMatrix(.Row, 16) = Work_Dept(2).List(0)
                        Exit Sub
                    End If
                   End If
                 Exit Sub
            Else
                If Work_Dept(0).ListCount > 1 Then
                    CenterObject Me, Work_Area
                    Work_Area.Visible = True
                    Work_Area.Refresh
                    Exit Sub
                Else
                    See_Msg "Only One Job Code Assigned", 5, ""
                End If
            End If
          Case Is = 1
             If .TextMatrix(.Row, 0) = "" Then
               Call See_Msg("Select Job Title First", 5, "")
               Exit Sub
             End If
            If .TextMatrix(.Row, 2) = "00:00" Then Exit Sub
            If .TextMatrix(.Row, 1) = "" Then
             If LoggedOnAs(0).EmpAccessLevel < Access_Required("Current Day Time Rec") Then
                .TextMatrix(.Row, 1) = Format(Date, "mm/DD/YYYY")
                .TextMatrix(.Row, 7) = Format(Date, "mm/DD/YYYY")
                CheckShiftDate CStr(.TextMatrix(.Row, 1)), Employees.Text
                Exit Sub
             End If
                If See_Msg("Use Todays Date", 4, "") = 6 Then
                 .TextMatrix(.Row, 1) = Format(Date, "mm/DD/YYYY")
                 .TextMatrix(.Row, 7) = Format(Date, "mm/DD/YYYY")
                  CheckShiftDate CStr(.TextMatrix(.Row, 1)), Employees.Text
                 Exit Sub
                End If
            End If
              If LoggedOnAs(0).EmpAccessLevel < Access_Required("Current Day Time Rec") Then Exit Sub
              If IsDate(.TextMatrix(.Row, 1)) Then
                 See_Msg "Cannot change date", 5, "delete time and add new record"
                 Exit Sub
              End If
              Dim PKT As DateOrTime
              Set PKT = New DateOrTime
              InDate = PKT.ReturnDate
              If DateDiff("d", InDate, Date) > Val(Input_Option(25)) Then
                 See_Msg "Invalid Date Selection", 5, "Edit is only for then last " & Val(Input_Option(25)) & " days"
                 Exit Sub
              End If
              If IsDate(InDate) Then
                .TextMatrix(.Row, 1) = Format(InDate, "mm/DD/YYYY")
                .TextMatrix(.Row, 7) = Format(InDate, "mm/DD/YYYY")
                'DoEvents
              End If
              Set PKT = Nothing
              CheckShiftDate CStr(.TextMatrix(.Row, 1)), Employees.Text
          Case Is = 2, 3, 4, 5
             If .TextMatrix(.Row, 0) = "" Then
               Call See_Msg("Select Job Title First", 5, "")
               Exit Sub
             End If
             If .TextMatrix(.Row, .Col) = "00:00" Then
                Call See_Msg("Can not Change Deleted Time", 5, "")
                Exit Sub
             End If
             .TextMatrix(.Row, .Col) = GetTime(.TextMatrix(.Row, .Col))
             If .TextMatrix(.Row, 1) = "" And .TextMatrix(.Row, 2) > "" Then .TextMatrix(.Row, 1) = Format(Date, "MM/DD/YYYY")
             If .Col = 5 Then
               If .TextMatrix(.Row, 5) > "" And Trim(.TextMatrix(.Row, 2)) > "" Then
               If IsDate(CDate(.TextMatrix(.Row, 5))) Then
                If CDate(.TextMatrix(.Row, 2)) > CDate(.TextMatrix(.Row, 5)) Then
                   .TextMatrix(.Row, 7) = Format(CDate(.TextMatrix(.Row, 1)) + 1, "MM/DD/YYYY")
                Else
                   .TextMatrix(.Row, 7) = Format(CDate(.TextMatrix(.Row, 1)), "MM/DD/YYYY")
                End If
                If StartAudit(.TextMatrix(.Row, 0)) = 1 And Val(.TextMatrix(.Row, 13)) = 0 Then
                  .TextMatrix(.Row, 13) = 0  'ShowNumberPad("Enter All Shift Tips", False, 0, 0)
                  See_Msg "Tip Edit Needed", 5, "Please Enter Tips from Tip Edit Menu"
                End If
               End If
              End If
             End If
             Qualifytime .Row
             Exit Sub
       End Select
    End With
End If
If DeleteRecord = 1 And TimeSheet.TextMatrix(TimeSheet.Row, 0) > "" Then
  If See_Msg("Are You Sure ?", 4, "Deletes Record") = 6 Then
  If SevenShiftsIsOn Then Call SevenShiftRemoveShift(TimeSheet.TextMatrix(TimeSheet.Row, 22))
    With TimeSheet
      .TextMatrix(.Row, 2) = "12:00 AM"
      .TextMatrix(.Row, 3) = "12:00 AM"
      .TextMatrix(.Row, 4) = "12:00 AM"
      .TextMatrix(.Row, 5) = "12:00 AM"
      .TextMatrix(.Row, 12) = WhoIsManager
    End With
  End If
  DeleteRecord = 0
End If
End Sub
Public Function SevenShiftRemoveShift(RecordId As String) As String
SevenShiftRemoveShift = ""
If RecordId = "" Then Exit Function
If SevenShiftsIsOn Then
 Dim SSF As SevenShift
 Set SSF = New SevenShift
  SevenShiftRemoveShift = SSF.DeleteTimePunch(RecordId)
 Set SSF = Nothing
End If
End Function
Private Sub CheckShiftDate(InDate As String, EName As String)
    If InDate = "" Then Exit Sub
    Dim Db As ADODB.Connection
    Dim Tb As ADODB.Recordset
    ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
    Sql = "Select * From [Employee Records] where [Employee Name] = '" & EName & "'"
    Sql = Sql & " and TRY_CAST([Shift Date] AS datetime) = '" & CDate(InDate) & "'"
    ConnectADORS Sql, Db, Tb
    If Not Tb.EOF Then
       See_Msg "Shift Found For This Date", 0, "Please Check Shift for Over Lap"
    End If
    DbClose Db, Tb
End Sub

Private Sub Work_Dept_Click(Index As Integer)
Dim FN         As Integer
Dim First_Name As String
Dim Last_Name  As String
Dim Tb         As ADODB.Recordset
Dim Db         As ADODB.Connection
Dim PRate      As Currency
Dim Hrly       As Integer
If Edit_Time = True Then
 TimeSheet.TextMatrix(TimeSheet.Row, 0) = Work_Dept(0).List(Work_Dept(0).ListIndex)
 TimeSheet.TextMatrix(TimeSheet.Row, 16) = Work_Dept(2).List(Work_Dept(0).ListIndex)
Else
If Trim$(Work_Dept(0).List(Work_Dept(0).ListIndex)) = "" Then
    Call See_Msg("Department Name Must be Selected", 5, "")
    Exit Sub
End If
Emp_Punch_In Work_Dept(0).List(Work_Dept(0).ListIndex), Val(Work_Dept(1).List(Work_Dept(0).ListIndex)), Val(Work_Dept(2).List(Work_Dept(0).ListIndex))
Call MemoryCheck(Time_Sheet)
Exit Sub
End If
Work_Area.Visible = False
Work_Area.Refresh
Exit Sub
End Sub
Public Sub Emp_Punch_In(JobWorking As String, RateofPay As Currency, PayBy As Integer)
On Error GoTo ErrorHandler
Dim PunchTimeValue      As String
Dim currentEmployee     As String
Dim ImWorking           As EmployeeData
Dim Response            As String
Dim Reply()             As String
Dim ShiftId             As String
Dim RecordId            As String
Dim GCID As GeneralLib.Utilities
Set GCID = New GeneralLib.Utilities
RecordId = Trim$(GCID.CreateGuid)
Set GCID = Nothing
Dim QtrMins             As Long
If Trim$(JobWorking) = "" Or Trim$(JobWorking) = "None" Then
    Call See_Msg("No Job Title Assigned.", 5, "")
    Exit Sub
End If
'***************
'Punching IN
'Put Scheduler here
'***************
If Check_Option(95) = 1 Then
    ImWorking = GetEmployee(JobWorking)
    If ImWorking.FirstName = "NotFound" Then
      See_Msg "No Job Code Assigned", 5, "Please assign Job Code."
      Exit Sub
    End If
    
    If SevenShiftsIsOn Then
    '  Dim ShiftOK As SevenShift
    '  Set ShiftOK = New SevenShift
      '  Response = ShiftOK.OkToPunch(ImWorking.ScheduleId, Val(Input_Option(22)))
    '  Set ShiftOK = Nothing
    '  Reply = Split(Response, "|")
    '  If Reply(0) = "success" Then
    '    If Reply(1) = "False" Then
    '       See_Msg "Your Not Scheduled", 5, "Please see a manager"
    '       If Can_Continue("Manager Menu") = False Then Exit Sub
    '    End If
    '  Else
    '    See_Msg Reply(1), 5, Reply(0)
    '    Exit Sub
    '  End If
    '  Punch_Name = currentEmployee
    End If
End If
'***************
'End of Employee Scheduler
'***************
'// make sure ownership will not change with group audits
currentEmployee = Punch_Name

If Check_Option(0) = 0 Then  '// option is for full service
  '// startaudit(JobWorking) = 1 will cause the software to auto start an audit , the value of 1 is job code is a tipped positon
  If StartAudit(JobWorking) = 1 Then StartServerAudit currentEmployee, True, JobWorking, 0
End If

Dim Db As ADODB.Connection
Dim OTB As ADODB.Recordset

ConnectADODB SvrPath.TerminalPath & TimeRecordsDb, Db
  Sql = "Select * from [Employee Records] where [employee name] = '" & Punch_Name & "' and [time in] > '' and [Time out] = ''"
ConnectADORS Sql, Db, OTB
    'If Check_Option(41) = 1 Then
    '  PunchTimeValue = QuarterRound(QtrMins, PDate)
    'Else
      PunchTimeValue = CStr(Format(Time, "HH:MM AmPm"))
    'End If
    If Trim(PunchTimeValue) = "" Then PunchTimeValue = CStr(Format(Time, "HH:MM AmPm"))
    
    If OTB.EOF Then
       ShiftId = ""
       
       If SevenShiftsIsOn Then
         Dim punchText As String
         punchText = Format$(Now, "yyyy-mm-dd hh:nn:ss")
         ImWorking = GetEmployee(JobWorking)
         
         Dim PIN As SevenShift
         Set PIN = New SevenShift
         If ImWorking.JobRoleId = "" Then ImWorking.JobRoleId = "0"
         If ImWorking.DepartmentId = "" Then ImWorking.DepartmentId = "0"
         Response = PIN.PunchIn(ImWorking.ScheduleId, ImWorking.JobRoleId, ImWorking.DepartmentId, punchText)
         
         Set PIN = Nothing
         
         Reply = Split(Response, "|")
         
         If Reply(0) = "success" Then
           ShiftId = Reply(1)
         Else
           Call See_Msg(Reply(2), 0, "Error with 7 Shifts")
         End If
       End If
       
       OTB.AddNew
       
       OTB.fields("Employee Name").Value = Punch_Name
       
       OTB![Department Name] = JobWorking & ""
       'If Check_Option(41) = 1 Then
       '  OTB![Shift date] = Format(PDate, "mm/dd/yyyy")
       'Else
         OTB![Shift date] = Format(Date, "mm/dd/yyyy")
       'End If
       OTB![Time In] = PunchTimeValue & ""
       OTB![Break Out] = ""
       OTB![Break In] = ""
       OTB![Time Out] = ""
       OTB![Punch out Date] = Format(Date, "mm/dd/yyyy")
       OTB![Tip Amount] = 0
       OTB![Altered By] = MRGOK
       OTB![Altered In] = 0
       OTB![Altered Out] = 0
       OTB![Altered Brk Out] = 0
       OTB![Altered Brk In] = 0
       OTB![Rate of Pay] = Val(RateofPay & "")
       OTB![Paid By] = Val(PayBy & "")
       OTB![Credit Card Tips] = 0
       If SevenShiftsIsOn Then
        '// add the shift id to the ssn number for reference.
        OTB.fields("SSN").Value = ShiftId
       Else
        OTB.fields("SSN").Value = GetSSN(Punch_Name)
       End If
       OTB![Tip Share] = 0
       OTB![Sales] = 0
       OTB![Tips Paid Out] = 0
       
       If Check_Option(0) = 1 Then
        OTB![Shift] = 0
       Else
        OTB![Shift] = GetShiftNo(Punch_Name)
       End If
       
       OTB.fields("RecordId").Value = RecordId
       OTB.Update
       
       DbClose Db, OTB
       
       See_Msg "Shift Started....", 5, PunchTimeValue
       
       '// send infor for video system.
       RemoteVideoSendItem " | " & Punch_Name & " Has Punched In"
       '// record for tracking
       Trans_Log Punch_Name & " Shift In " & Now
       '//
    Else
      See_Msg "You are already on the clock...", 5, ""
    End If
Exit Sub
ErrorHandler:
ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, "Punch In ")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Function CheckBreak(BrkOut As String, BrkIn As String) As Boolean
On Error Resume Next
CheckBreak = True
If Trim(BrkOut) = "" Then Exit Function
    If Check_Option(20) = 1 Then
        If DateDiff("n", CDate(BrkOut), CDate(BrkIn), vbMonday) < MealBreakMins Then 'option 7 = Maniditory break
            CheckBreak = False
        End If
    End If
End Function

                                                                                                                          