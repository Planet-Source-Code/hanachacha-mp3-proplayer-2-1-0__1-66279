VERSION 5.00
Object = "{0179B2D7-CD62-439D-BE78-CF820F5A4B44}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmPlaylist 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   37
   ShowInTaskbar   =   0   'False
   Begin M3P_Control.ctlSlider sldPl 
      Height          =   3735
      Left            =   4080
      TabIndex        =   10
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   6588
      BackColor       =   -2147483643
      BorderStyle     =   0
      SliderOrien     =   0
      CueWidth        =   1
      CueHeight       =   249
      ForeColor       =   -2147483647
      MouseIcon       =   "frmPlayList.frx":000C
   End
   Begin M3P_Control.DynamicButton btnPlaylist 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.ctlList List 
      Height          =   3735
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6588
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5640
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrTest 
      Interval        =   25
      Left            =   4680
      Top             =   720
   End
   Begin M3P_Control.DynamicButton btnPlaylist 
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   4320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.DynamicButton btnPlaylist 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.DynamicButton btnPlaylist 
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.DynamicButton btnPlaylist 
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   8
      Top             =   4320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.DynamicButton btnPlaylist 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.Label lblCurrentperTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   60
   End
   Begin VB.Label lblTotalTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "        "
      Height          =   240
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Const WM_MOUSEWHEEL = 522
Private Const PM_REMOVE = &H1

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long

Private HoverMe As Boolean

Dim i As Long

Private PlaylistRgnSrc(1) As RECT
Private PlaylistReg(2) As Long

Private Sub ResizeForm()
    On Error Resume Next
    Dim W As Long
    Dim H As Long
    
    W = CLng(ReadINI("Playlist", "Width", tCurrentSkin.strConfig)) * Screen.TwipsPerPixelX
    H = CLng(ReadINI("Playlist", "Height", tCurrentSkin.strConfig)) * Screen.TwipsPerPixelY
    If Me.width < W Then Me.width = W
    If Me.height < H Then Me.height = H
    
    Dim lngW As Long
    Dim lngH As Long
        
    lngW = Me.ScaleWidth
    lngH = Me.ScaleHeight
        
    List.Move PlaylistRad(10).Right, PlaylistRad(0).Bottom, lngW - (PlaylistRad(10).Right + PlaylistRad(11).Right), lngH - (PlaylistRad(0).Bottom + PlaylistRad(5).Bottom)
    sldPl.height = lngH - (PlaylistRad(1).Bottom + PlaylistRad(6).Bottom)
    sldPl.Left = lngW - PlaylistRad(11).Right + (ReadINI("ScrollList", "Left", tCurrentSkin.strConfig))
    
    If List.ListItemCount <= List.ItemPerPage Then
        sldPl.Enabled = False
        tmrTest.Enabled = False
    Else
        sldPl.Enabled = True
        tmrTest.Enabled = True
    End If
    
    Dim lngAdd As Long
    Dim lngSub As Long
    Dim lngSelect As Long
    Dim lngSort As Long
    Dim lngEditT As Long, lngEditL As Long
    Dim lngClose As Long
        
    lngAdd = (ReadINI("Add", "Top", tCurrentSkin.strConfig))
    lngSub = (ReadINI("Sub", "Top", tCurrentSkin.strConfig))
    lngSelect = (ReadINI("Select", "Top", tCurrentSkin.strConfig))
    lngSort = (ReadINI("Sort", "Top", tCurrentSkin.strConfig))
    lngEditT = (ReadINI("EditList", "Top", tCurrentSkin.strConfig))
    lngEditL = (ReadINI("EditList", "Left", tCurrentSkin.strConfig))
    lngClose = (ReadINI("ClosePL", "Left", tCurrentSkin.strConfig))
        
    btnPlaylist(0).Top = lngH - PlaylistRad(5).Bottom + lngAdd
    btnPlaylist(1).Top = lngH - PlaylistRad(5).Bottom + lngSub
    btnPlaylist(2).Top = lngH - PlaylistRad(5).Bottom + lngSelect
    btnPlaylist(3).Top = lngH - PlaylistRad(5).Bottom + lngSort
    btnPlaylist(4).Top = lngH - PlaylistRad(5).Bottom + lngEditT
    '
    btnPlaylist(4).Left = lngW - PlaylistRad(6).Right + lngEditL
    btnPlaylist(5).Left = lngW - PlaylistRad(1).Right + lngClose
        
    lblTotalTime.Left = lngW - PlaylistRad(1).Right - lblTotalTime.width
    
    Call DrawMain
    
    
    'CreateRoundRectRgn
    PlaylistRgnSrc(0).Left = 0
    PlaylistRgnSrc(0).Top = 0
    PlaylistRgnSrc(0).Right = Me.ScaleWidth
    PlaylistRgnSrc(0).Bottom = List.Top
    
    PlaylistRgnSrc(1).Left = 0
    PlaylistRgnSrc(1).Top = List.Top + List.height
    PlaylistRgnSrc(1).Right = Me.ScaleWidth
    PlaylistRgnSrc(1).Bottom = Me.ScaleHeight + PlaylistRgnSrc(1).Top
    
    PlaylistReg(0) = MakeRecRegion(Me, PlaylistRgnSrc(0), RGB(255, 0, 255))
    PlaylistReg(1) = MakeRecRegion(Me, PlaylistRgnSrc(1), RGB(255, 0, 255))
    PlaylistReg(2) = CreateRectRgn(0, List.Top, Me.ScaleWidth, Me.ScaleHeight - List.Top)
    
    CombineRgn PlaylistReg(0), PlaylistReg(0), PlaylistReg(1), RGN_OR
    CombineRgn PlaylistReg(2), PlaylistReg(2), PlaylistReg(0), RGN_OR
    
    SetWindowRgn Me.hwnd, PlaylistReg(2), True
    
    DeleteObject PlaylistReg(0)
    DeleteObject PlaylistReg(1)
    DeleteObject PlaylistReg(2)
    
End Sub

Private Sub btnPlaylist_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrTest.Enabled = False
End Sub

Private Sub btnPlayList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Select Case Index
            Case 0
                PopupMenu frmMenu.mnuAdd, vbPopupMenuLeftAlign
            Case 1
                PopupMenu frmMenu.mnuSub, vbPopupMenuLeftAlign
            Case 2
                PopupMenu frmMenu.mnuSelect, vbPopupMenuLeftAlign
            Case 3
                PopupMenu frmMenu.mnuSort, vbPopupMenuLeftAlign
            Case 4
                PopupMenu frmMenu.mnuManaList, vbPopupMenuLeftAlign
            Case 5
                Me.Hide
                tPlaylistConfig.bolHidePL = True
                frmMedia.btnMedia(8).bolOn = False
                frmMedia.btnMedia(8).Refresh
                frmMedia.btnMiniMedia(8).bolOn = False
                frmMedia.btnMiniMedia(8).Refresh
                frmMenu.mnuShowPL.Checked = False
        End Select
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        If Not bolCDPlay Then
            Call Play(currentIndex)
        Else
            Call PlayCD(tCDplay.CurrentDrive, currentIndex - 1)
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 36 Then Call GoTop  'Home
    If KeyCode = 66 Then Call BackTrack   'B
    If KeyCode = 78 Then Call NextTrack   'N
    If KeyCode = 35 Then Call GoBottom  'End
    If KeyCode = 32 Then Call StopPlayer  'Spacebar
    If KeyCode = vbKeyDelete Then
        frmPlayList.MousePointer = 11
        For i = frmPlayList.List.ListItemCount To 1 Step -1
            If frmPlayList.List.ListItemSelect(i) = True Then
                SubFile (i)
            End If
        Next i
        frmPlayList.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Select Case tPlaylistConfig.intSortKey
        Case 0
            frmMenu.mnuSortC(0).Checked = True
        Case 1
            frmMenu.mnuSortC(1).Checked = True
        Case 2
            frmMenu.mnuSortC(2).Checked = True
        Case 3
            frmMenu.mnuSortC(3).Checked = True
        Case 4
            frmMenu.mnuSortC(4).Checked = True
    End Select
        
    frmMenu.mnuSortStyleC(0).Checked = True
    Call SetSortString(tPlaylistConfig.strSortString)
    
    List.Number = tPlaylistConfig.bolShowNumber
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        If Me.MousePointer = 0 Then
            DragForm Me
            Call SnapForm(frmMedia, Me)
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sizeX As Boolean, sizeY As Boolean
    Dim bolDown As Boolean
    Dim MouseRad As Integer
    
    If Button = vbLeftButton Then
        bolDown = True
    Else
        bolDown = False
    End If
    
    If x > Me.ScaleWidth - 10 Or x < 10 Then
        sizeX = True
        If x < 10 Then
            MouseRad = 0
        Else
            MouseRad = 1
        End If
    Else
        sizeX = False
    End If
    
    If y > Me.ScaleHeight - 10 Or y < 10 Then
        sizeY = True
        If y < 10 Then
            MouseRad = 2
        Else
            MouseRad = 3
        End If
    Else
        sizeY = False
    End If
    
    If sizeX And sizeY Then
        If x < 10 And y < 10 Then
            MouseRad = 4
            Me.MousePointer = 8
        End If
        If x > 10 And y > 10 Then
            MouseRad = 5
            Me.MousePointer = 8
        End If
        If x < 10 And y > 10 Then
            MouseRad = 6
            Me.MousePointer = 6
        End If
        If x > 10 And y < 10 Then
            MouseRad = 7
            Me.MousePointer = 6
        End If
        ElseIf sizeY And (bolDown = False) Then
            Me.MousePointer = 7
        Else
            If sizeX And (bolDown = False) Then
                Me.MousePointer = 9
            Else
                If bolDown = False Then
                    Me.MousePointer = 0
                    MouseRad = 100
                End If
            End If
    End If
    
    If Button = vbLeftButton Then
        Dim rec As RECT
        Dim pPoint As POINTAPI
        GetWindowRect Me.hwnd, rec
        GetCursorPos pPoint
        On Error Resume Next
        Debug.Print MouseRad
        Select Case MouseRad
            Case 7
                Me.Move Me.Left, pPoint.y * Screen.TwipsPerPixelY, (pPoint.x - rec.Left) * Screen.TwipsPerPixelX, (rec.Bottom - pPoint.y) * Screen.TwipsPerPixelY
            Case 6
                Me.Move pPoint.x * Screen.TwipsPerPixelX, Me.Top, (rec.Right - pPoint.x) * Screen.TwipsPerPixelX, (pPoint.y - rec.Top) * Screen.TwipsPerPixelY
            Case 5
                Me.Move Me.Left, Me.Top, (pPoint.x - rec.Left) * Screen.TwipsPerPixelX, (pPoint.y - rec.Top) * Screen.TwipsPerPixelY
            Case 4
                Me.Move pPoint.x * Screen.TwipsPerPixelX, pPoint.y * Screen.TwipsPerPixelY, (rec.Right - pPoint.x) * Screen.TwipsPerPixelX, (rec.Bottom - pPoint.y) * Screen.TwipsPerPixelY
            Case 3
                Me.Move Me.Left, Me.Top, Me.width, (pPoint.y - rec.Top) * Screen.TwipsPerPixelY
            Case 2
                Me.Move Me.Left, pPoint.y * Screen.TwipsPerPixelY, Me.width, (rec.Bottom - pPoint.y) * Screen.TwipsPerPixelY
            Case 1
                Me.Move Me.Left, Me.Top, (pPoint.x - rec.Left) * Screen.TwipsPerPixelX, Me.height
            Case 0
                Me.Move pPoint.x * Screen.TwipsPerPixelX, Me.Top, (rec.Right - pPoint.x) * Screen.TwipsPerPixelX, Me.height
        End Select
    End If
    tmrTest.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMain, vbPopupMenuLeftAlign
    End If
    Call SnapForm(frmMedia, Me)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim strFilePath As String
    Dim intRecord As Integer
    Dim fn As Integer
    
    If Not bolCDPlay Then
        fn = FreeFile
        strFilePath = App.path & "\M3P.m3u"
        If frmPlayList.List.ListItemCount > 0 Then
            Call SaveM3U(strFilePath)
        End If
    End If
    WriteINI "Demension", "PlaylistTop", Me.Top, strFileconfig
    WriteINI "Demension", "PlaylistLeft", Me.Left, strFileconfig
    WriteINI "Demension", "PlaylistWidth", Me.width, strFileconfig
    WriteINI "Demension", "PlaylistHeight", Me.height, strFileconfig
End Sub

Private Sub Form_Resize()
    Call ResizeForm
    'If List.ItemPerPage <> 0 Then
     '   List.height = (List.ItemPerPage * List.itemHeight)
     '   Me.height = (List.height + (PlaylistRad(0).H + PlaylistRad(5).H)) * Screen.TwipsPerPixelY
    'End If
End Sub


Private Sub List_DblClick()
    On Error Resume Next
    If Not bolCDPlay Then
        Call Play(List.CurrentPlayItem)
    Else
        Call PlayCD(tCDplay.CurrentDrive, List.CurrentPlayItem - 1)
    End If
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
        currentIndex = List.CurrentItem
        sldPl.value = currentIndex
        Call sldPl_Change(currentIndex)
    End If
    If KeyCode = 37 Then Call Prev   'Left <--
    If KeyCode = 39 Then Call Forw   'Right -->
End Sub

Private Sub List_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        currentIndex = List.CurrentItem
    End If
    If Button = vbRightButton Then
        currentRIndex = List.CurrentItem
    End If
    sldPl.value = List.CurrentItem
    Call sldPl_Change(List.CurrentItem)
End Sub

Private Sub List_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrTest.Enabled = False
    HoverMe = True
End Sub


Private Sub List_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If List.ListItemCount > 0 Then PopupMenu frmMenu.mnuPlaylist, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub sldPl_Change(lValue As Long)
    On Error GoTo beep
    Dim strText As String
    Dim lngTop As Long
    Dim tmpTrack As track
    Dim x As Long
    If List.ListItemCount > 1 Then
        If lValue <= List.ListItemCount Then
            List.ItemVisible (lValue)
        End If
        If tPlaylistConfig.intLoadID = 1 And (bolCDPlay = False) Then
            For x = lValue To (lValue + List.ItemPerPage)
                If x > List.ListItemCount Then Exit For
                If NowPlaying(List.Key(x)).strText = GetShortName(NowPlaying(List.Key(x)).Infor.FullName) Then
                    If LibOption.bolUse Then
                        NowPlaying(List.Key(x)).Infor = Library(TrackIndex(NowPlaying(List.Key(x)).Infor.FullName)).Infor
                    Else
                        Call OpenTrack(NowPlaying(List.Key(x)).Infor.FullName, tmpTrack)
                        NowPlaying(List.Key(x)).Infor = tmpTrack
                    End If
                    strText = ""
                    strText = tPlaylistConfig.strDisplay
                    strText = Replace(strText, "%1", NowPlaying(List.Key(x)).Infor.Artist)
                    strText = Replace(strText, "%2", NowPlaying(List.Key(x)).Infor.Title)
                    strText = Replace(strText, "%3", NowPlaying(List.Key(x)).Infor.Album)
                    strText = Replace(strText, "%4", NowPlaying(List.Key(x)).Infor.Genre)
                    strText = Replace(strText, "%5", NowPlaying(List.Key(x)).Infor.Year)
                    strText = Replace(strText, "%6", NowPlaying(List.Key(x)).Infor.Filename)
                    strText = Replace(strText, "%7", NowPlaying(List.Key(x)).Infor.FullName)
                    NowPlaying(List.Key(x)).strText = strText
                    
                    List.ListItemText(x) = strText
                    List.ListItemTime(x) = Time2String(NowPlaying(List.Key(x)).Infor.Duration)
                    List.Number = tPlaylistConfig.bolShowNumber
                Else
                    Exit Sub
                End If
           Next x
       End If
   End If
beep:
    If sldPl.value = 0 Then sldPl.value = 1
    Exit Sub
End Sub

Private Sub tmrTest_Timer()
    If Not TestOver Then
        HoverMe = False
        tmrTest.Enabled = False
    Else
        HoverMe = True
    End If
    ProcessMessages
End Sub
Private Sub DrawMain()
    On Error Resume Next
    '-- Do Not Change This !!!
    Me.Cls
    'left top
    Me.PaintPicture picSource, 0, 0, PlaylistRad(0).Right, PlaylistRad(0).Bottom, PlaylistRad(0).Left, PlaylistRad(0).Top, PlaylistRad(0).Right, PlaylistRad(0).Bottom, vbSrcCopy
    'left top rez
    Me.PaintPicture picSource, PlaylistRad(0).Right, 0, Me.ScaleWidth / 2 - PlaylistRad(2).Right / 2 - PlaylistRad(0).Right + 1, PlaylistRad(3).Bottom, PlaylistRad(3).Left, PlaylistRad(3).Top, PlaylistRad(3).Right, PlaylistRad(3).Bottom, vbSrcCopy
    'right top rez
    Me.PaintPicture picSource, Me.ScaleWidth / 2 + PlaylistRad(2).Right / 2 - 1, 0, Me.ScaleWidth / 2 - PlaylistRad(2).Right / 2, PlaylistRad(3).Bottom, PlaylistRad(4).Left, PlaylistRad(4).Top, PlaylistRad(4).Right, PlaylistRad(4).Bottom, vbSrcCopy
    'middle top
    Me.PaintPicture picSource, Me.ScaleWidth / 2 - PlaylistRad(2).Right / 2, 0, PlaylistRad(2).Right, PlaylistRad(2).Bottom, PlaylistRad(2).Left, PlaylistRad(2).Top, PlaylistRad(2).Right, PlaylistRad(2).Bottom, vbSrcCopy
    'right top
    Me.PaintPicture picSource, Me.ScaleWidth - PlaylistRad(1).Right, 0, PlaylistRad(1).Right, PlaylistRad(1).Bottom, PlaylistRad(1).Left, PlaylistRad(1).Top, PlaylistRad(1).Right, PlaylistRad(1).Bottom, vbSrcCopy
    'left bottom
    Me.PaintPicture picSource, 0, Me.ScaleHeight - PlaylistRad(5).Bottom, PlaylistRad(5).Right, PlaylistRad(5).Bottom, PlaylistRad(5).Left, PlaylistRad(5).Top, PlaylistRad(5).Right, PlaylistRad(5).Bottom, vbSrcCopy
    'left bottom rez
    Me.PaintPicture picSource, PlaylistRad(5).Right, Me.ScaleHeight - PlaylistRad(8).Bottom, Me.ScaleWidth / 2 - PlaylistRad(7).Right / 2 - PlaylistRad(5).Right + 1, PlaylistRad(8).Bottom, PlaylistRad(8).Left, PlaylistRad(8).Top, PlaylistRad(8).Right, PlaylistRad(8).Bottom, vbSrcCopy
    'right bottom rez
    Me.PaintPicture picSource, Me.ScaleWidth / 2 + PlaylistRad(7).Right / 2 - 1, Me.ScaleHeight - PlaylistRad(9).Bottom, Me.ScaleWidth / 2 - PlaylistRad(7).Right / 2, PlaylistRad(9).Bottom, PlaylistRad(9).Left, PlaylistRad(9).Top, PlaylistRad(9).Right, PlaylistRad(9).Bottom, vbSrcCopy
    'middle bottom
    Me.PaintPicture picSource, Me.ScaleWidth / 2 - PlaylistRad(7).Right / 2, Me.ScaleHeight - PlaylistRad(7).Bottom, PlaylistRad(7).Right, PlaylistRad(7).Bottom, PlaylistRad(7).Left, PlaylistRad(7).Top, PlaylistRad(7).Right, PlaylistRad(7).Bottom, vbSrcCopy
    'right bottom
    Me.PaintPicture picSource, Me.ScaleWidth - PlaylistRad(6).Right, Me.ScaleHeight - PlaylistRad(6).Bottom, PlaylistRad(6).Right, PlaylistRad(6).Bottom, PlaylistRad(6).Left, PlaylistRad(6).Top, PlaylistRad(6).Right, PlaylistRad(6).Bottom, vbSrcCopy
    'left middle
    Me.PaintPicture picSource, 0, PlaylistRad(0).Bottom, PlaylistRad(10).Right, Me.ScaleHeight - PlaylistRad(0).Bottom - PlaylistRad(5).Bottom, PlaylistRad(10).Left, PlaylistRad(10).Top, PlaylistRad(10).Right, PlaylistRad(10).Bottom, vbSrcCopy
    'right middle
    Me.PaintPicture picSource, Me.ScaleWidth - PlaylistRad(11).Right, PlaylistRad(1).Bottom, PlaylistRad(11).Right, Me.ScaleHeight - (PlaylistRad(1).Bottom + PlaylistRad(6).Bottom), PlaylistRad(11).Left, PlaylistRad(11).Top, PlaylistRad(11).Right, PlaylistRad(11).Bottom, vbSrcCopy
    
End Sub
Private Function TestOver() As Boolean
    Dim rad As POINTAPI
    GetCursorPos rad
    TestOver = (WindowFromPoint(rad.x, rad.y) = Me.hwnd)
End Function
Private Sub ProcessMessages()
    Dim Message As msg
    Dim ItemPage As Integer
    
    
    Do While HoverMe
        WaitMessage 'Wait For message and...
        If PeekMessage(Message, Me.hwnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...
            ItemPage = List.ItemPerPage
            If Message.wParam < 0 Then '...scroll up...
                If sldPl.value + tPlaylistConfig.intRowS < List.ListItemCount Then
                    If tPlaylistConfig.intRowS <= ItemPage Then
                        sldPl.value = sldPl.value + tPlaylistConfig.intRowS
                    Else
                        sldPl.value = sldPl.value + ItemPage
                    End If
                Else
                    sldPl.value = List.ListItemCount
                End If
            Else '... or scroll down
                If sldPl.value > tPlaylistConfig.intRowS Then
                    If tPlaylistConfig.intRowS <= ItemPage Then
                        sldPl.value = sldPl.value - tPlaylistConfig.intRowS
                    Else
                        sldPl.value = sldPl.value - ItemPage
                    End If
                Else
                    sldPl.value = 1
                    List.ItemVisible (1)
                End If
            End If
            Call sldPl_Change(sldPl.value)
        End If
        DoEvents
    Loop
End Sub
