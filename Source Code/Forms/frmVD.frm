VERSION 5.00
Object = "{0179B2D7-CD62-439D-BE78-CF820F5A4B44}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmVD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MP3_proPlayer : Movie"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   Icon            =   "frmVD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   138
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picScr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      Begin VB.Label lblScr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   24
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   569
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   8535
      Begin M3P_Control.Progress prgStatus 
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Value           =   0
         BackColor       =   8388736
         BorderColor     =   12632256
         ForeColor       =   9224224
         MinColor        =   8388608
         MouseIcon       =   "frmVD.frx":23D2
      End
      Begin VB.Label lblButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   360
         Index           =   4
         Left            =   7200
         TabIndex        =   13
         Top             =   165
         Width           =   300
      End
      Begin VB.Label lblButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         Index           =   3
         Left            =   6540
         TabIndex        =   12
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         Index           =   2
         Left            =   6000
         TabIndex        =   11
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         Index           =   1
         Left            =   5460
         TabIndex        =   10
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         Index           =   0
         Left            =   4920
         TabIndex        =   9
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Progress"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1605
      End
   End
   Begin M3P_Control.DynamicButton btnVideo 
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrPos 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2640
      Top             =   3480
   End
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   18
      Left            =   2040
      Top             =   3480
   End
   Begin M3P_Control.DynamicButton btnVideo 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.DynamicButton btnVideo 
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.DynamicButton btnVideo 
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   4
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin M3P_Control.Video Video 
      Height          =   3975
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7011
   End
End
Attribute VB_Name = "frmVD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bolShow As Boolean

Dim i As Integer
Private rad As POINTAPI
Private intStart As Integer
Private VideoRgnSrc(1) As RECT
Private VideoReg(2) As Long


Private Function TestOver() As Boolean
    Dim rad As POINTAPI
    GetCursorPos rad
    TestOver = (WindowFromPoint(rad.x, rad.y) = Me.hwnd)
End Function

Private Sub DrawMain()
    On Error Resume Next
    '-- Do Not Change This !!!
    Me.Cls
    'left top
    Me.PaintPicture picSource, 0, 0, VideoRad(0).Right, VideoRad(0).Bottom, VideoRad(0).Left, VideoRad(0).Top, VideoRad(0).Right, VideoRad(0).Bottom, vbSrcCopy
    'left top rez
    Me.PaintPicture picSource, VideoRad(0).Right, 0, Me.ScaleWidth / 2 - VideoRad(2).Right / 2 - VideoRad(0).Right + 1, VideoRad(3).Bottom, VideoRad(3).Left, VideoRad(3).Top, VideoRad(3).Right, VideoRad(3).Bottom, vbSrcCopy
    'right top rez
    Me.PaintPicture picSource, Me.ScaleWidth / 2 + VideoRad(2).Right / 2 - 1, 0, Me.ScaleWidth / 2 - VideoRad(2).Right / 2, VideoRad(3).Bottom, VideoRad(4).Left, VideoRad(4).Top, VideoRad(4).Right, VideoRad(4).Bottom, vbSrcCopy
    'middle top
    Me.PaintPicture picSource, Me.ScaleWidth / 2 - VideoRad(2).Right / 2, 0, VideoRad(2).Right, VideoRad(2).Bottom, VideoRad(2).Left, VideoRad(2).Top, VideoRad(2).Right, VideoRad(2).Bottom, vbSrcCopy
    'right top
    Me.PaintPicture picSource, Me.ScaleWidth - VideoRad(1).Right, 0, VideoRad(1).Right, VideoRad(1).Bottom, VideoRad(1).Left, VideoRad(1).Top, VideoRad(1).Right, VideoRad(1).Bottom, vbSrcCopy
    'left bottom
    Me.PaintPicture picSource, 0, Me.ScaleHeight - VideoRad(5).Bottom, VideoRad(5).Right, VideoRad(5).Bottom, VideoRad(5).Left, VideoRad(5).Top, VideoRad(5).Right, VideoRad(5).Bottom, vbSrcCopy
    'left bottom rez
    Me.PaintPicture picSource, VideoRad(5).Right, Me.ScaleHeight - VideoRad(8).Bottom, Me.ScaleWidth / 2 - VideoRad(7).Right / 2 - VideoRad(5).Right + 1, VideoRad(8).Bottom, VideoRad(8).Left, VideoRad(8).Top, VideoRad(8).Right, VideoRad(8).Bottom, vbSrcCopy
    'right bottom rez
    Me.PaintPicture picSource, Me.ScaleWidth / 2 + VideoRad(7).Right / 2 - 1, Me.ScaleHeight - VideoRad(9).Bottom, Me.ScaleWidth / 2 - VideoRad(7).Right / 2, VideoRad(9).Bottom, VideoRad(9).Left, VideoRad(9).Top, VideoRad(9).Right, VideoRad(9).Bottom, vbSrcCopy
    'middle bottom
    Me.PaintPicture picSource, Me.ScaleWidth / 2 - VideoRad(7).Right / 2, Me.ScaleHeight - VideoRad(7).Bottom, VideoRad(7).Right, VideoRad(7).Bottom, VideoRad(7).Left, VideoRad(7).Top, VideoRad(7).Right, VideoRad(7).Bottom, vbSrcCopy
    'right bottom
    Me.PaintPicture picSource, Me.ScaleWidth - VideoRad(6).Right, Me.ScaleHeight - VideoRad(6).Bottom, VideoRad(6).Right, VideoRad(6).Bottom, VideoRad(6).Left, VideoRad(6).Top, VideoRad(6).Right, VideoRad(6).Bottom, vbSrcCopy
    'left middle
    Me.PaintPicture picSource, 0, VideoRad(0).Bottom, VideoRad(10).Right, Me.ScaleHeight - VideoRad(0).Bottom - VideoRad(5).Bottom, VideoRad(10).Left, VideoRad(10).Top, VideoRad(10).Right, VideoRad(10).Bottom, vbSrcCopy
    'right middle
    Me.PaintPicture picSource, Me.ScaleWidth - VideoRad(11).Right, VideoRad(1).Bottom, VideoRad(11).Right, Me.ScaleHeight - (VideoRad(1).Bottom + VideoRad(6).Bottom), VideoRad(11).Left, VideoRad(11).Top, VideoRad(11).Right, VideoRad(11).Bottom, vbSrcCopy
    
End Sub

Private Sub btnVideo_TurnOn(Index As Integer)
    Select Case Index
        Case 0
            Video.Display = DefaultSize
            frmMenu.mnuScreenC(4).Enabled = False
        Case 1
            Video.Display = DoubleSize
            frmMenu.mnuScreenC(4).Enabled = False
        Case 2
            Video.Display = Fullscreen
            frmMenu.mnuScreenC(4).Enabled = True
        Case 3
            Call StopPlayer
            Me.Hide
    End Select
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If Video.Display <> CustomizeSize Then Video.Display = CustomizeSize
        Me.WindowState = vbNormal
        frmMenu.mnuScreenC(4).Enabled = False
    End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        If Me.MousePointer = 0 Then
            DragForm Me
            Call SnapForm(frmMedia, Me)
            If tPlaylistConfig.bolHidePL = False Then Call SnapForm(frmPlayList, Me)
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
    
    If Me.MousePointer <> 0 Then Video.Display = CustomizeSize
    
    If Button = vbLeftButton Then
        Dim rec As RECT
        Dim pPoint As POINTAPI
        
        GetWindowRect Me.hwnd, rec
        GetCursorPos pPoint
        
        On Error Resume Next
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
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMain, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call DrawMain
    If Me.WindowState <> vbMinimized Then
        If Me.WindowState = vbNormal Then
            If Video.Display = CustomizeSize Then
                If tDevice.VideoD.bolLockRatio Then
                    Me.width = Me.height * tDevice.VideoD.intRatioWidth / tDevice.VideoD.intRatioHeight
                End If
                Video.Move VideoRad(10).Right, VideoRad(0).Bottom, Me.ScaleWidth - (VideoRad(10).Right + VideoRad(11).Right), Me.ScaleHeight - (VideoRad(0).Bottom + VideoRad(5).Bottom)
            End If
        End If
        If Me.WindowState = vbMaximized And Video.Display <> Fullscreen Then
            Video.Move VideoRad(10).Right, VideoRad(0).Bottom, Me.ScaleWidth - (VideoRad(10).Right + VideoRad(11).Right), Me.ScaleHeight - (VideoRad(0).Bottom + VideoRad(5).Bottom)
            Video.ZOrder 1
            tmrPos.Enabled = False
            picStatus.Visible = False
            picScr.Visible = False
        End If
        Call ResizeVideo
        'Dim reg As Long
        'reg = MakeRegion(Me, RGB(255, 0, 255))
        'SetWindowRgn Me.hwnd, reg, True
    End If
End Sub


Private Sub lblButton_Click(Index As Integer)
    Select Case Index
        Case 0 'prev
            Call BackTrack
        Case 1 'play
            Call frmMedia.btnPlayClick
        Case 2 'pause
            Call frmMedia.btnPauseClick
        Case 3
            Call NextTrack
        Case 4
            Call StopPlayer
    End Select
End Sub

Private Sub lblButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmp As Integer
    lblButton(Index).ForeColor = RGB(200, 200, 200)
    For tmp = 0 To 4
        If tmp <> Index Then
            lblButton(tmp).ForeColor = &H800080
        End If
    Next
    If Video.State = playing Then
        lblButton(1).ForeColor = RGB(0, 127, 255)
    Else
        lblButton(2).ForeColor = RGB(0, 127, 255)
    End If
End Sub

Private Sub lblScr_Click()
    If Video.Display <> CustomizeSize Then Video.Display = CustomizeSize
    Me.WindowState = vbNormal
    frmMenu.mnuScreenC(4).Enabled = False
End Sub

Private Sub prgStatus_Change(lValue As Long)
    If Video.State = playing Then
        Video.Position = lValue
    End If
End Sub

Private Sub tmrPos_Timer()
    On Error Resume Next
    prgStatus.value = Video.Position
End Sub

Private Sub Video_DisplaySizeChange()
    If Video.Display = DefaultSize Or Video.Display = DoubleSize Or Video.Display = HaftSize Then
        Me.WindowState = vbNormal
        Me.width = (Video.width + VideoRad(10).Right + VideoRad(11).Right) * Screen.TwipsPerPixelX
        Me.height = (Video.height + VideoRad(0).Bottom + VideoRad(5).Bottom) * Screen.TwipsPerPixelY
        If Me.width > Screen.width Or Me.height > Screen.height Then Video.Display = DefaultSize
        Video.Move VideoRad(10).Right, VideoRad(0).Bottom
        tmrPos.Enabled = False
        picStatus.Visible = False
        picScr.Visible = False
    End If
    If Video.Display = Fullscreen Then
        Me.WindowState = vbMaximized
        Video.Move 0, 0
        Video.ZOrder 0
        tmrPos.Enabled = True
        VideoReg(0) = CreateRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight)
        SetWindowRgn Me.hwnd, VideoReg(0), True
    End If
    If Video.Display = CustomizeSize Then
        If Me.WindowState = vbMaximized Then
            Video.Move VideoRad(10).Right, VideoRad(0).Bottom, Me.ScaleWidth - (VideoRad(10).Right + VideoRad(11).Right), Me.ScaleHeight - (VideoRad(0).Bottom + VideoRad(5).Bottom)
            Video.ZOrder 1
            tmrPos.Enabled = False
            picStatus.Visible = False
            picScr.Visible = False
        End If
    End If
    For i = 0 To frmMenu.mnuScreenC.Count - 1
        frmMenu.mnuScreenC(i).Checked = False
    Next i
    If Video.Display = HaftSize Then frmMenu.mnuScreenC(0).Checked = True: Call AlwaysOnTop(Me, False)
    If Video.Display = DefaultSize Then frmMenu.mnuScreenC(1).Checked = True: Call AlwaysOnTop(Me, False)
    If Video.Display = DoubleSize Then frmMenu.mnuScreenC(2).Checked = True: Call AlwaysOnTop(Me, False)
    If Video.Display = Fullscreen Then frmMenu.mnuScreenC(3).Checked = True: Call AlwaysOnTop(Me, True)
End Sub

Public Sub ResizeVideo()
    On Error Resume Next
    
    Dim strImgVideo As String
    Dim tmpDefault As String
    strImgVideo = tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Video.bmp"
    
    If FileExists(strImgVideo) = False Then
        tmpDefault = App.path & "\Skins\Default\Default.skn"
    Else
        tmpDefault = tCurrentSkin.strConfig
    End If
        
    If Me.width < 275 * Screen.TwipsPerPixelX Then Me.width = 275 * Screen.TwipsPerPixelX
    If Me.height < 116 * Screen.TwipsPerPixelY Then Me.height = 116 * Screen.TwipsPerPixelY
    
    Call DrawMain
    
    Dim lngW As Long
    Dim lngH As Long
        
    lngW = Me.ScaleWidth
    lngH = Me.ScaleHeight
        
    btnVideo(0).Top = lngH - VideoRad(5).Bottom + (ReadINI("1X", "Top", tmpDefault))
    btnVideo(1).Top = lngH - VideoRad(5).Bottom + (ReadINI("2X", "Top", tmpDefault))
    btnVideo(2).Top = lngH - VideoRad(5).Bottom + (ReadINI("FS", "Top", tmpDefault))
    btnVideo(3).Left = lngW - VideoRad(1).Right + (ReadINI("CloseVideo", "Left", tmpDefault))
    picStatus.Move 0, Me.ScaleHeight - picStatus.height, Me.ScaleWidth
    picScr.Move Me.ScaleWidth - picScr.width, 0
    
    If Video.Display <> Fullscreen Then
        'CreateRoundRectRgn
        VideoRgnSrc(0).Left = 0
        VideoRgnSrc(0).Top = 0
        VideoRgnSrc(0).Right = Me.ScaleWidth
        VideoRgnSrc(0).Bottom = Video.Top
        
        VideoRgnSrc(1).Left = 0
        VideoRgnSrc(1).Top = Video.Top + Video.height
        VideoRgnSrc(1).Right = Me.ScaleWidth
        VideoRgnSrc(1).Bottom = Me.ScaleHeight + VideoRgnSrc(1).Top
        
        VideoReg(0) = MakeRecRegion(Me, VideoRgnSrc(0), RGB(255, 0, 255))
        VideoReg(1) = MakeRecRegion(Me, VideoRgnSrc(1), RGB(255, 0, 255))
        VideoReg(2) = CreateRectRgn(0, Video.Top, Me.ScaleWidth, Me.ScaleHeight - Video.Top)
        
        CombineRgn VideoReg(0), VideoReg(0), VideoReg(1), RGN_OR
        CombineRgn VideoReg(2), VideoReg(2), VideoReg(0), RGN_OR
        
        SetWindowRgn Me.hwnd, VideoReg(2), True
    End If
End Sub
Private Sub tmrMouse_Timer()
    Dim RadT As POINTAPI
    GetCursorPos RadT
    If RadT.x <> rad.x Or RadT.y <> rad.y Then
        Video.HideCursor False
        intStart = 0
        If Video.Display = Fullscreen Then
            picStatus.Visible = True
            picStatus.ZOrder 0
            picScr.Visible = True
            picScr.ZOrder 0
            Video.Move 0, 0, Video.width, Screen.height / Screen.TwipsPerPixelY - picStatus.ScaleHeight
        End If
        tmrMouse.Enabled = False
    Else
        intStart = intStart + 1
        If intStart >= 100 Then
            Video.HideCursor True
            If Video.Display = Fullscreen Then
                picStatus.Visible = False
                picStatus.ZOrder 1
                picScr.Visible = False
                picScr.ZOrder 0
                Video.Move 0, 0, Video.width, Screen.height / Screen.TwipsPerPixelY
            End If
            intStart = 0
        End If
    End If
End Sub


Private Sub Video_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    GetCursorPos rad
    tmrMouse.Enabled = True
End Sub

Private Sub Video_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(frmMenu.mnuScreen, vbPopupMenuLeftAlign)
    End If
End Sub

Private Sub Video_PlayStateChange()
    If Video.State <> Stopped Then
        If frmVD.Video.State = playing Then
            frmMedia.btnMedia(1).bolOn = True
            frmMedia.btnMedia(2).bolOn = False
            frmMedia.btnMedia(2).Enabled = True
            lblButton(1).ForeColor = RGB(0, 127, 255)
            lblButton(2).ForeColor = &H800080
        Else
            frmMedia.btnMedia(2).bolOn = True
            frmMedia.btnMedia(2).Enabled = True
            frmMedia.btnMedia(1).bolOn = False
            lblButton(2).ForeColor = RGB(0, 127, 255)
            lblButton(1).ForeColor = &H800080
        End If
    Else
        frmMedia.btnMedia(1).bolOn = False
        frmMedia.btnMedia(2).bolOn = False
        frmMedia.btnMedia(2).Enabled = False
    End If
    With frmMedia
        .btnMedia(1).Refresh
        .btnMedia(2).Refresh
        .btnMiniMedia(1).bolOn = .btnMedia(1).bolOn
        .btnMiniMedia(2).bolOn = .btnMedia(2).bolOn
        .btnMiniMedia(1).Refresh
        .btnMiniMedia(2).Refresh
        If frmPlayList.List.ListItemCount <= 1 Then
            .btnMedia(0).Enabled = False
            .btnMedia(3).Enabled = False
            .btnMiniMedia(0).Enabled = False
            .btnMiniMedia(3).Enabled = False
        Else
            .btnMedia(0).Enabled = True
            .btnMedia(3).Enabled = True
            .btnMiniMedia(0).Enabled = True
            .btnMiniMedia(3).Enabled = True
        End If
    End With
End Sub


