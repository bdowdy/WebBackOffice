VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#9.4#0"; "3dabm9u.ocx"
Begin VB.Form RemoteVideo 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "RearDisplay"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin BTNENHLib4.BtnEnh ScrollBanner 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   10200
      Width           =   15135
      _Version        =   589828
      _ExtentX        =   26696
      _ExtentY        =   1296
      _StockProps     =   66
      Caption         =   "Scroll"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
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
      SpecialEffectFactor=   3
      UserData        =   0.1
      textCaption     =   "RemoteVideo.frx":0000
      textLT          =   "RemoteVideo.frx":006C
      textCT          =   "RemoteVideo.frx":0084
      textRT          =   "RemoteVideo.frx":009C
      textLM          =   "RemoteVideo.frx":00B4
      textRM          =   "RemoteVideo.frx":00CC
      textLB          =   "RemoteVideo.frx":00E4
      textCB          =   "RemoteVideo.frx":00FC
      textRB          =   "RemoteVideo.frx":0114
      colorBack       =   "RemoteVideo.frx":012C
      colorIntern     =   "RemoteVideo.frx":0156
      colorMO         =   "RemoteVideo.frx":0180
      colorFocus      =   "RemoteVideo.frx":01AA
      colorDisabled   =   "RemoteVideo.frx":01D4
      colorPressed    =   "RemoteVideo.frx":01FE
   End
   Begin VB.Timer ScrollBannerTimer 
      Left            =   120
      Top             =   7680
   End
   Begin VSFlex8Ctl.VSFlexGrid VirtualCheck 
      Height          =   9855
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   14895
      _cx             =   1995663009
      _cy             =   1995654119
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
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   200
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      DataMember      =   ""
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
   Begin VB.Frame VideoBox 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   14895
      Begin WMPLibCtl.WindowsMediaPlayer RearVideoPlayer 
         Height          =   3000
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   4500
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   7938
         _cy             =   5292
      End
   End
   Begin VB.Image SlideShow 
      Height          =   9855
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   14895
   End
End
Attribute VB_Name = "RemoteVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DefaultImage    As String
Private DefaultVideo    As String
Private TempEvent       As String
Private Slides()        As String
Private SlidesPause     As Integer
Dim CurrentPos          As Integer
Dim RunShow             As Boolean
Dim RemoteVideoGrid()   As Variant
Dim RemoteRowNo()       As Integer
Dim PlayVideo           As Boolean
Public Sub Load_VirtualCheck()
On Error GoTo ErrorHandler
If Check_Option(17) = 1 Then Exit Sub
Dim Rw     As Integer
Dim Cl     As Integer
Dim CRL    As Integer
Dim SeeIt1 As Integer
Dim SeeIt2 As Integer
Dim RF As Currency
If DualPricing > 0 Then
  RF = GetCashCheck(False)
  Sales_Totals
Else
  RF = GetProcessingFee(0)
End If
ConsolidateVirtualCheck
VirtualCheckIs True
With VirtualCheck
   .Clear
   .TopRow = 0
   
    For Rw = 0 To RemoteRowNo(SeatNumber) - 1
     For Cl = 0 To 2
       If Cl = 1 And Left(RemoteVideoGrid(SeatNumber, Rw, Cl), 1) = "*" Then
         .Cell(flexcpForeColor, Rw, 0, Rw, 2) = &HC0&
       End If
       If Cl = 2 And Val(RemoteVideoGrid(SeatNumber, Rw, Cl) & "") < 0 Then
         .Cell(flexcpForeColor, Rw, 0, Rw, 2) = &H800080
       End If
       If Cl = 0 And Val(RemoteVideoGrid(SeatNumber, Rw, 0) & "") = 0 And Trim(Grid_array(SeatNumber, Rw, 1)) > "" Then
         .Cell(flexcpForeColor, Rw, 0, Rw, 2) = &HC00000
       End If
       .TextMatrix(Rw, Cl) = RemoteVideoGrid(SeatNumber, Rw, Cl) & ""
     Next Cl
    Next Rw
   If Gross_Receipt > 0 Then
    .TextMatrix(RemoteRowNo(SeatNumber), 0) = ""
    .TextMatrix(RemoteRowNo(SeatNumber), 1) = ""
    .TextMatrix(RemoteRowNo(SeatNumber), 2) = ""
    .Cell(flexcpForeColor, RemoteRowNo(SeatNumber) + 1, 0, RemoteRowNo(SeatNumber) + 2, 2) = &HC000C0
    .TextMatrix(RemoteRowNo(SeatNumber) + 1, 1) = "Sales Tax"
    .TextMatrix(RemoteRowNo(SeatNumber) + 1, 2) = Format(Chk_Sales_data(1) + Chk_Sales_data(3) + Chk_Sales_data(2) + ((Chk_Sales_data(12) + Chk_Sales_data(20)) + Abs(Chk_Sales_data(14))), ".00")
        If ProcessingFeesOn Or CashDiscountData.CreditAmount > 0 Then
         .TextMatrix(RemoteRowNo(SeatNumber) + 2, 1) = "Cash Total"
         If ProcessingFeesOn Then
          .TextMatrix(RemoteRowNo(SeatNumber) + 2, 2) = Format(GetProcessingCashPrice, "0.00")
         Else
          .TextMatrix(RemoteRowNo(SeatNumber) + 2, 2) = Format(RF, "0.00")
         End If
         .TextMatrix(RemoteRowNo(SeatNumber) + 3, 1) = "Total"
         .TextMatrix(RemoteRowNo(SeatNumber) + 3, 2) = Format(Chk_Sales_data(5), "0.00")
        Else
         .TextMatrix(RemoteRowNo(SeatNumber) + 2, 1) = "Total"
         .TextMatrix(RemoteRowNo(SeatNumber) + 2, 2) = Format(Chk_Sales_data(5), "0.00")
        End If
    End If
End With
Sleep 0
SeeIt1 = RemoteRowNo(SeatNumber) - 1: SeeIt2 = RemoteRowNo(SeatNumber) + 2
SeeItOnCheck SeeIt1, SeeIt2
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " Load Virtual Grid")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub VirtualCheckIs(IsVisable As Boolean)
On Error Resume Next
If PlayVideo Then
 If IsVisable Then
  SlideShow.width = (15015 * WidScale1024) / 2
  RearVideoPlayer.windowlessVideo = False
  RearVideoPlayer.Controls.pause
  RearVideoPlayer.width = 0 ' (15015 * WidScale1024) / 2
  VideoBox.width = 0 '(RemoteVideo.Width / 2) - 500
  SlideShow.Visible = True
  VirtualCheck.Visible = True
  VirtualCheck.ZOrder 0
  ScrollBannerTimer.Enabled = True
  Sleep 0
 Else
  VirtualCheck.Visible = False
  SlideShow.Visible = False
  RearVideoPlayer.Visible = True
  VideoBox.Visible = True
  RearVideoPlayer.width = (15015 * WidScale1024)
  VideoBox.width = (15015 * WidScale1024)
  RearVideoPlayer.windowlessVideo = False
  RearVideoPlayer.Controls.play
  ScrollBannerTimer.Enabled = False
 End If
Else
 If IsVisable Then
  VirtualCheck.Visible = True
  VirtualCheck.ZOrder 0
  If SlideShow.width <> Int(((15015 * WidScale1024) / 2)) Then SlideShow.width = (15015 * WidScale1024) / 2
 Else
  VirtualCheck.Visible = False
  If SlideShow.width <> (15015 * WidScale1024) Then SlideShow.width = (15015 * WidScale1024)
 End If
End If
Sleep 0
End Sub
Public Sub VirtualCheckChange(Tendered As String, TendAmount As Currency)
If Check_Option(17) = 1 Then Exit Sub
On Error GoTo ErrorHandler
Dim Rw     As Integer
Dim Cl     As Integer
Dim CRL    As Integer
Dim SeeIt1 As Integer
Dim SeeIt2 As Integer
With VirtualCheck
      .Cell(flexcpForeColor, RemoteRowNo(SeatNumber) + 3, 0, RemoteRowNo(SeatNumber) + 3, 2) = &H8000&
      .TextMatrix(RemoteRowNo(SeatNumber) + 3, 1) = Tendered
      If TendAmount = 0 Then
       .TextMatrix(RemoteRowNo(SeatNumber) + 3, 2) = ""
       .TextMatrix(RemoteRowNo(SeatNumber) + 3, 1) = "No Change"
      Else
       .TextMatrix(RemoteRowNo(SeatNumber) + 3, 2) = Format(TendAmount, "0.00")
      End If
End With
SeeIt1 = RemoteRowNo(SeatNumber) - 1: SeeIt2 = RemoteRowNo(SeatNumber) + 3
SeeItOnCheck SeeIt1, SeeIt2
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " Load Virtual Change")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Public Sub VirtualCheckClearChange()
On Error GoTo ErrorHandler
If Check_Option(17) = 1 Then Exit Sub
With VirtualCheck
      .TextMatrix(RemoteRowNo(SeatNumber) + 3, 1) = ""
      .TextMatrix(RemoteRowNo(SeatNumber) + 3, 2) = ""
End With
Exit Sub
ErrorHandler:
  ErrorRespose = Main_Error_Handler(Err.Number, Err.Description, " Load Virtual Clear Change")
   Select Case ErrorRespose
     Case "Abort"
       Exit Sub
     Case Is = "Resume"
       Resume
     Case Is = "Resume Next"
       Resume Next
   End Select
End Sub
Private Function ConsolidateVirtualCheck()
On Error Resume Next
Erase RemoteVideoGrid
Erase RemoteRowNo
Dim RemoteGrid()        As Variant
Dim RemoteNextSeat()    As Integer
Dim CNip                As Integer
Dim CSn                 As Integer
Dim CBal                As Integer
Dim CutAst              As String
Dim CPName1             As String
Dim CPName2             As String
Dim SS                  As Integer   'Search Seat
Dim C                   As Integer
RemoteGrid() = Grid_array()
RemoteNextSeat() = NextRow()
ReDim RemoteVideoGrid(Number_In_Party, MaxRows, TTL_Col) As Variant
ReDim RemoteRowNo(Number_In_Party)
   For CNip = 0 To Number_In_Party - 1
    For CSn = 0 To RemoteNextSeat(CNip) - 1
      If Left(Trim(RemoteGrid(CNip, CSn, 1)), 1) = "v" Or Val(RemoteGrid(CNip, CSn, 13)) = "1" Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 2)) < 0 Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 0)) = 0 Then GoTo NextItem
      If Val(RemoteGrid(CNip, CSn, 0)) > 0 And Val(RemoteGrid(CNip, CSn + 1, 0)) = 0 And Trim(RemoteGrid(CNip, CSn + 1, 1)) > "" Then GoTo NextItem
      If Trim(RemoteGrid(CNip, CSn, 1)) = "<>" Then GoTo NextItem
      For SS = CSn + 1 To RemoteNextSeat(CNip) - 1
      If Trim(RemoteGrid(CNip, SS, 1)) = "<>" Or Val(RemoteGrid(CNip, SS, 13)) = "1" Then GoTo NextLine
      If Val(RemoteGrid(CNip, SS, 0)) = 0 And RemoteGrid(CNip, SS, 1) > "" Then GoTo NextLine
      If Val(RemoteGrid(CNip, CSn, 0)) > 0 And Val(RemoteGrid(CNip, SS, 0)) > 0 Then
        If Val(RemoteGrid(CNip, SS + 1, 0)) = 0 And Trim(RemoteGrid(CNip, SS + 1, 1)) > "" Then GoTo NextLine
          CPName1 = Trim(RemoteGrid(CNip, CSn, 1))
          CPName2 = Trim(RemoteGrid(CNip, SS, 1))
             CutAst = Left(Trim(RemoteGrid(CNip, SS, 1)), 1)
             If CutAst <> "v" Then
             If Trim(CPName1) = Trim(CPName2) Then
               If CutAst = "*" Then
                If RemoteGrid(CNip, CSn, 31) = "None" Then RemoteGrid(CNip, CSn, 31) = ""
                RemoteGrid(CNip, CSn, 31) = CStr(RemoteGrid(CNip, CSn, 31)) & CStr(RemoteGrid(CNip, SS, 14))
               End If
               RemoteGrid(CNip, CSn, 0) = Val(RemoteGrid(CNip, CSn, 0)) + Val(RemoteGrid(CNip, SS, 0))
               RemoteGrid(CNip, CSn, 2) = Format(Val(RemoteGrid(CNip, CSn, 2)) + Val(RemoteGrid(CNip, SS, 2)), "0.00")
               RemoteGrid(CNip, SS, 1) = "<>": RemoteGrid(CNip, SS, 2) = 0: RemoteGrid(CNip, SS, 12) = 0
              End If
             End If
        End If
NextLine:
       Next SS
NextItem:
      Next CSn
   Next CNip
   C = -1
   For CNip = 0 To Number_In_Party - 1
    RemoteRowNo(CNip) = 0
    C = -1
    For CSn = 0 To RemoteNextSeat(CNip) - 1
      If RemoteGrid(CNip, CSn, 1) <> "<>" Then
         C = C + 1
         RemoteRowNo(CNip) = RemoteRowNo(CNip) + 1
         For SS = 0 To TTL_Col
          RemoteVideoGrid(CNip, C, SS) = RemoteGrid(CNip, CSn, SS)
         Next SS
      End If
     Next CSn
  Next CNip
Erase RemoteGrid
Erase RemoteNextSeat
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Enter_Id.IDbox.Text) = 0 Then Enter_Id.IDbox.SetFocus
If KeyCode = 13 Then
  Enter_Id.OK.DoClick
End If
'// need to keep the focus on the enter id box for the pos for enter id number.
End Sub
Private Sub Form_Load()
ResizeScreen1024 Me
Me.Move Screen.width, 0, Screen.width, Screen.Height
'// Moves the screen to the extended desktop
Dim C             As Integer  '// counter for image files
Dim LT            As Integer
Dim Db            As ADODB.Connection
Dim Tb            As ADODB.Recordset
Dim MediaList     As String   '// image file list
Dim VideoList     As String   '// video file
Dim SideBar       As String   '// side bar is for modifier when virtual check is visable
ScrollBannerTimer.Enabled = False
DefaultImage = "OCD Image\": DefaultVideo = "OCD Video\"
RunShow = True
PlayVideo = False
VideoList = Dir(SvrPath.LocalTerminalPath & DefaultVideo & "*.*")
 If VideoList > "" Then
    PlayVideo = True
 End If
 If Dir(SvrPath.LocalTerminalPath & DefaultImage & "*.*") = "" Then
   Check_Option(229) = 1
 End If
 TempEvent = GetCurrentDirectory
 '// TempEvent is for the side show or for day parts for images
 With VirtualCheck
    .HighLight = flexHighlightNever
    .rows = MaxRows
    .cols = 3
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignRightCenter
   If Check_Option(229) = 1 Then
    .ColWidth(0) = ((1700 * WidScale1024) * (Screen.width / (12000 * WidScale1024)))
    .ColWidth(1) = ((8000 * WidScale1024) * (Screen.width / (12000 * WidScale1024)))
    .ColWidth(2) = ((2050 * WidScale1024) * (Screen.width / (12000 * WidScale1024)))
    .width = (.ColWidth(0) + .ColWidth(1) + .ColWidth(2)) + 100
    .Height = Screen.Height - (1000 * WidScale1024)
    .FontSize = 15 * WidScale1024
    .Top = 100
    .Left = 100
   Else
    .ColWidth(0) = ((1700 * WidScale1024) * (Screen.width / (12000 * WidScale1024))) / 2
    .ColWidth(1) = ((8000 * WidScale1024) * (Screen.width / (12000 * WidScale1024))) / 2
    .ColWidth(2) = ((2050 * WidScale1024) * (Screen.width / (12000 * WidScale1024))) / 2
    .width = (.ColWidth(0) + .ColWidth(1) + .ColWidth(2)) + 100
    .Height = Screen.Height - (1000 * WidScale1024)
    .Left = Screen.width - (.width)
    .FontSize = 15 * WidScale1024
    .Top = 100
  End If
End With
SlideShow.Height = Screen.Height - (1000 * WidScale1024)
'// set the SlideShow heigth the same as the virtual checks
 SideBar = ""
 If SideBar = "" Then SideBar = TempEvent
 MediaList = Dir(SvrPath.LocalTerminalPath & SideBar & "*.*")
 If MediaList = "" Then
    MediaList = Dir(SvrPath.LocalTerminalPath & DefaultImage & "*.*")
    If MediaList > "" Then SideBar = DefaultImage
 End If
 ReDim Slides(0)
 C = -1
 Do Until MediaList = vbNullString
  C = C + 1
  ReDim Preserve Slides(C)
  Slides(C) = MediaList
  MediaList = Dir
 Loop

 If PlayVideo Then
   VideoBox.Visible = True
   SlideShow.Visible = False
   RearVideoPlayer.Visible = True
   RearVideoPlayer.URL = SvrPath.LocalTerminalPath & DefaultVideo & VideoList
   RearVideoPlayer.settings.setMode "loop", True
   RearVideoPlayer.Controls.play
   RearVideoPlayer.ZOrder 0
 End If
 If C > -1 Then
   SlidesPause = 5
   SlideShow.Picture = LoadPicture(SvrPath.LocalTerminalPath & SideBar & Slides(0))
   If PlayVideo Then
    ScrollBannerTimer.Enabled = False
   Else
    ScrollBannerTimer.Enabled = True
   End If
   ScrollBannerTimer.Interval = SlidesPause * 1000
   RunShow = True
 End If
 
 If Check_Option(77) = 1 Then
    ScrollBanner.Visible = False
 Else
   If ConnectADODB(SvrPath.TerminalPath & "MSaleData", Db) = False Then Exit Sub
   Sql = "Select * from [Daily Specials] Where [Day] = '" & GetDayOfWeek(DatePart("w", Now, vbMonday)) & "'"
   OpenTb Sql, Tb, Db
   If Not Tb.EOF Then
     ScrollBanner.Caption = Tb.fields("Special 1").Value & " " & Tb.fields("Special 2").Value & " " & Tb.fields("Special 3").Value & " " & Tb.fields("Special 4").Value & " "
   Else
     SetNothing Tb
     Sql = "Select * from [Daily Specials] Where [Day] = 'Default'"
     OpenTb Sql, Tb, Db
     If Not Tb.EOF Then
       ScrollBanner.Caption = Tb.fields("Special 1").Value & " " & Tb.fields("Special 2").Value & " " & Tb.fields("Special 3").Value & " " & Tb.fields("Special 4").Value & " "
     End If
   End If
   DbClose Db, Tb
 End If
 CurrentPos = -1
 ScrollBanner.Top = Screen.Height - ScrollBanner.Height
End Sub
Public Function GetCurrentDirectory() As String
On Error GoTo ERH
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
Dim R  As Integer
GetCurrentDirectory = "OCD Image\"
If ConnectADODB(SvrPath.LocalTerminalPath & RegisterOptionsDB, Db) = False Then Exit Function
 Sql = "Select * from [Rear Display] Where '" & Format(CDate(Time), "yyyy-mm-dd hh:nn:ss") & "' Between TRY_CAST([Starts] AS datetime) and TRY_CAST([Ends] AS datetime)"
 Sql = Sql & " and '" & Format(CDate(Date), "yyyy-mm-dd hh:nn:ss") & "' Between TRY_CAST([Start Date] AS datetime) and TRY_CAST([End date] AS datetime)"
If ConnectADORS(Sql, Db, RS) = False Then Exit Function
If Not RS.EOF Then
  GetCurrentDirectory = "OCD Image\" & RS.fields("Name").Value & "\"
Else
  SetNothing RS
  Sql = "Select * from [Rear Display] Where '" & Format(CDate(Time), "yyyy-mm-dd hh:nn:ss") & "' Between TRY_CAST([Starts] AS datetime) and TRY_CAST([Ends] AS datetime)"
  If ConnectADORS(Sql, Db, RS) = False Then Exit Function
   If Not RS.EOF Then
      Select Case UCase(RS.fields("Name").Value)
       Case Is = "BREAKFAST", "LUNCH", "THRU", "DINNER", "LATE"
         GetCurrentDirectory = "OCD Image\" & RS.fields("Name").Value & "\"
      End Select
   End If
End If
DbClose Db, RS
Exit Function
ERH:
Error_Log Err.Number, Err.Description, "Rear Display", "None"
Err.Clear
End Function
Public Sub ShowModPic(ModPicName As String)
  If PlayVideo Then Exit Sub
  Dim MediaList As String
  MediaList = Dir(SvrPath.LocalTerminalPath & "OCD Image\Modifiers\" & ModPicName)
  If MediaList = "" Then Exit Sub
  VirtualCheckIs False
  SlideShow.Picture = LoadPicture(SvrPath.LocalTerminalPath & "OCD Image\Modifiers\" & ModPicName)
  SlideShow.Visible = True
  ScrollBannerTimer.Enabled = False
End Sub
Public Sub HideModPic()
    ScrollBannerTimer.Enabled = True
    ScrollBannerTimer_Timer
End Sub
Private Sub RearVideoPlayer_PlayStateChange(ByVal NewState As Long)
  If NewState = 8 Then
     '// 8 = video player ended play
  End If
End Sub
Private Sub ScrollBannerTimer_Timer()
On Error Resume Next
Dim SideBar         As String
Dim MediaList       As String
Dim C               As Integer
 If Check_Option(229) = 1 Then
  If VirtualCheck.Visible = True Then Exit Sub
 End If
 TempEvent = GetCurrentDirectory
 SideBar = ""
 If VirtualCheck.Visible = True Then
    MediaList = Dir(SvrPath.LocalTerminalPath & TempEvent & "SideBar\" & "*.*")
    If MediaList > "" Then SideBar = TempEvent & "SideBar\"
    If MediaList = "" Then
       MediaList = Dir(SvrPath.LocalTerminalPath & DefaultImage & "SideBar\" & "*.*")
       If MediaList > "" Then SideBar = DefaultImage & "SideBar\"
    End If
 End If
 If SideBar = "" Then SideBar = TempEvent
 MediaList = Dir(SvrPath.LocalTerminalPath & SideBar & "*.*")
 If MediaList = "" Then
    MediaList = Dir(SvrPath.LocalTerminalPath & DefaultImage & "*.*")
    If MediaList > "" Then SideBar = DefaultImage
 End If
 ReDim Slides(0)
 C = -1
 Do Until MediaList = vbNullString
    C = C + 1
    ReDim Preserve Slides(C)
    Slides(C) = MediaList
    MediaList = Dir
 Loop

CurrentPos = CurrentPos + 1
If CurrentPos > UBound(Slides, 1) Then CurrentPos = 0
SlideShow.Visible = False
SlideShow.Picture = LoadPicture(SvrPath.LocalTerminalPath & SideBar & Slides(CurrentPos))
If PlayVideo Then
  RearVideoPlayer.Visible = False
End If
SlideShow.Visible = True
If Gross_Receipt = 0 Then
  VirtualCheckIs False
Else
  VirtualCheckIs True
  VirtualCheck.ZOrder 0
End If
End Sub
Private Sub SeeItOnCheck(SeeIt1 As Integer, SeeIt2 As Integer)
    Dim SeeIt As Integer
    If Check_Option(229) = 0 Then
      If RemoteRowNo(SeatNumber) < 17 Then VirtualCheck.FontSize = 15 * WidScale1024
      Select Case CInt(VirtualCheck.FontSize)
       Case Is <= 12 * WidScale1024
            If RemoteRowNo(SeatNumber) < 23 Then
               VirtualCheck.TopRow = 0
               Exit Sub
            End If
            With VirtualCheck
             If Val(SeeIt1 + SeeIt2) <= 0 Then
               SeeIt = RemoteRowNo(SeatNumber)
             Else
               If (SeeIt1 - 22) < 0 Then
                .TopRow = 0
               Else
                .TopRow = SeeIt1 - 22
               End If
             End If
             Exit Sub
            End With
       Case Is = 13 * WidScale1024
            If RemoteRowNo(SeatNumber) < 21 Then
               VirtualCheck.TopRow = 0
               Exit Sub
            End If
            VirtualCheck.FontSize = 11 * WidScale1024
            With VirtualCheck
             If Val(SeeIt1 + SeeIt2) <= 0 Then
               SeeIt = RemoteRowNo(SeatNumber)
             Else
               If (SeeIt1 - 21) < 0 Then
                .TopRow = 0
               Else
                .TopRow = SeeIt1 - 21
               End If
             End If
             Exit Sub
            End With
       Case Else
           If RemoteRowNo(SeatNumber) < 17 Then
               VirtualCheck.TopRow = 0
               Exit Sub
            End If
            VirtualCheck.FontSize = 13 * WidScale1024
             If Screen.width > 15000 Then
              VirtualCheck.FontSize = (11.25 * WidScale1024)
             End If
             With VirtualCheck
                 SeeIt = RemoteRowNo(SeatNumber) + 5
                Sleep 0
                Exit Sub
             End With
      End Select
    Else
      If RemoteRowNo(SeatNumber) < 10 Then VirtualCheck.FontSize = 15 * WidScale1024
      Select Case CInt(VirtualCheck.FontSize)
       Case Is = 15 * WidScale1024
            If RemoteRowNo(SeatNumber) < 19 Then
               VirtualCheck.TopRow = 0
               Exit Sub
            End If
            With VirtualCheck
             If Val(SeeIt1 + SeeIt2) <= 0 Then
               SeeIt = RemoteRowNo(SeatNumber)
             Else
               If (SeeIt1 - 16) < 0 Then
                .TopRow = 0
               Else
                .TopRow = SeeIt1 - 16
               End If
             End If
             Exit Sub
            End With
       Case Is = 20 * WidScale1024
            If RemoteRowNo(SeatNumber) < 15 Then
               VirtualCheck.TopRow = 0
               Exit Sub
            End If
            VirtualCheck.FontSize = 15 * WidScale1024
            With VirtualCheck
             If Val(SeeIt1 + SeeIt2) <= 0 Then
               SeeIt = RemoteRowNo(SeatNumber)
             Else
               If (SeeIt1 - 12) < 0 Then
                .TopRow = 0
               Else
                .TopRow = SeeIt1 - 12
               End If
             End If
             Exit Sub
            End With
       
       Case Else
           If RemoteRowNo(SeatNumber) < 11 Then
               VirtualCheck.TopRow = 0
               Exit Sub
            End If
            VirtualCheck.FontSize = 15 * WidScale1024
             With VirtualCheck
               If Val(SeeIt1 + SeeIt2) <= 0 Then
                 SeeIt = RemoteRowNo(SeatNumber)
               Else: .TopRow = 0
                Exit Sub
               End If
             End With
      End Select
    End If
End Sub

