VERSION 5.00
Begin VB.UserControl Video 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   KeyPreview      =   -1  'True
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ToolboxBitmap   =   "Video.ctx":0000
   Begin VB.Timer tmrEnd 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   3120
   End
End
Attribute VB_Name = "Video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private obj_MediaControl As IMediaControl
Private obj_Audio As IBasicAudio
Private obj_Video As IBasicVideo
Private obj_MediaEvent As IMediaEvent
Private obj_VideoWindow As IVideoWindow
Private obj_MediaPosition As IMediaPosition

Enum vdDisplaySize
    [HaftSize] = 0
    [DefaultSize] = 1
    [DoubleSize] = 2
    [CustomizeSize] = 3
    [FullScreen] = 4
End Enum

Enum vdPlayState
    [Paused] = 2
    [Playing] = 1
    [Stopped] = 0
End Enum

Private Type VideoSize
    Width As Long
    Height As Long
End Type

Event EndOfStream()
Event DisplaySizeChange()
Event PlayStateChange()
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Resize()

Dim PlayState As vdPlayState
Dim DisplayType As vdDisplaySize
Dim HaftS As VideoSize
Dim DefaultS As VideoSize
Dim DoubleS As VideoSize
Dim FullScreenS As VideoSize
Dim CurrentSize As VideoSize

Dim lngDur As Long
Dim lngVol As Long
Dim lngBal As Long
Dim dblRate As Double

Dim Filename As String

Public Property Get ScrW() As Long
    ScrW = DefaultS.Width
End Property
Public Property Get ScrH() As Long
    ScrH = DefaultS.Height
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Get Rate() As Double
    Rate = dblRate
End Property
Public Property Let Rate(New_Val As Double)
    On Error Resume Next
    dblRate = New_Val
    If ObjPtr(obj_MediaPosition) > 0 Then
        obj_MediaPosition.Rate = dblRate
    End If
End Property
Public Property Get Balance() As Long
    Balance = lngBal
End Property
Public Property Let Balance(New_Val As Long)
    lngBal = New_Val
    If ObjPtr(obj_Audio) > 0 Then
        obj_Audio.Balance = New_Val * 100 '(New_val is between -100 <> 100)
    End If
End Property
Public Property Get Volume() As Long
    Volume = lngVol
End Property
Public Property Let Volume(New_Val As Long)
    lngVol = New_Val
    If New_Val > 100 Then New_Val = 100
    If New_Val < 0 Then New_Val = 0
    If ObjPtr(obj_Audio) > 0 Then obj_Audio.Volume = (New_Val * 100) - 10000
End Property
Public Property Get State() As vdPlayState
    State = PlayState
End Property
Public Property Get Display() As vdDisplaySize
    Display = DisplayType
End Property
Public Property Let Display(New_Val As vdDisplaySize)
    Dim tmp As vdDisplaySize
    tmp = DisplayType
    Select Case New_Val
        Case HaftSize
            CurrentSize.Width = DefaultS.Width
            CurrentSize.Height = DefaultS.Height
            DisplayType = New_Val
            UserControl.Height = CurrentSize.Height * Screen.TwipsPerPixelY
            UserControl.Width = CurrentSize.Width * Screen.TwipsPerPixelX
            SetSize (UserControl.ScaleWidth - HaftS.Width) / 2, (UserControl.ScaleHeight - HaftS.Height) / 2, HaftS.Width, HaftS.Height
            RaiseEvent DisplaySizeChange
            Exit Property
        Case DefaultSize
            CurrentSize.Width = DefaultS.Width
            CurrentSize.Height = DefaultS.Height
            If tmp = HaftSize Then
                DisplayType = New_Val
                UserControl.Height = CurrentSize.Height * Screen.TwipsPerPixelY
                UserControl.Width = CurrentSize.Width * Screen.TwipsPerPixelX
                SetSize , , DefaultS.Width, DefaultS.Height
                RaiseEvent DisplaySizeChange
                Exit Property
            End If
        Case DoubleSize
            CurrentSize.Width = DoubleS.Width
            CurrentSize.Height = DoubleS.Height
        Case FullScreen
            CurrentSize.Width = FullScreenS.Width
            CurrentSize.Height = FullScreenS.Height
        Case CustomizeSize
            CurrentSize.Width = UserControl.ScaleWidth
            CurrentSize.Height = UserControl.ScaleHeight
    End Select
    DisplayType = New_Val
    UserControl.Height = CurrentSize.Height * Screen.TwipsPerPixelY
    UserControl.Width = CurrentSize.Width * Screen.TwipsPerPixelX
    RaiseEvent DisplaySizeChange
End Property
Public Property Get Duration() As Long
    Duration = lngDur
End Property
Public Property Get Position() As Long
    If ObjPtr(obj_MediaPosition) > 0 Then Position = obj_MediaPosition.CurrentPosition
End Property
Public Property Let Position(New_Val As Long)
    If ObjPtr(obj_MediaPosition) > 0 Then
        If New_Val < 0 Then New_Val = 0
        If New_Val > lngDur Then New_Val = lngDur
        If obj_MediaPosition.CanSeekBackward Then
            If obj_MediaPosition.CanSeekForward Then
                obj_MediaPosition.CurrentPosition = New_Val
            End If
        End If
    End If
End Property

Public Sub OpenVideo(strMedia As String)
    Call CloseVideo
    Set obj_MediaControl = New FilgraphManager
    Call obj_MediaControl.RenderFile(strMedia)
    Set obj_Audio = obj_MediaControl
    Set obj_Video = obj_MediaControl
    Set obj_VideoWindow = obj_MediaControl
    Set obj_MediaEvent = obj_MediaControl
    Set obj_MediaPosition = obj_MediaControl

    DefaultS.Width = obj_Video.SourceWidth
    DefaultS.Height = obj_Video.SourceHeight
    Call CalSize(DefaultS)
    lngDur = obj_MediaPosition.Duration
    obj_MediaPosition.Rate = dblRate
    With obj_VideoWindow
        .WindowStyle = &H6000000
        .Owner = UserControl.hwnd
        .MessageDrain = UserControl.hwnd
    End With
End Sub

Public Sub PlayVideo(Optional from As Long)
    On Error Resume Next
    If ObjPtr(obj_MediaPosition) > 0 Then
        If from = vbNull Then from = 0
        obj_MediaPosition.CurrentPosition = from
        Call obj_MediaControl.Run
        PlayState = Playing
        RaiseEvent PlayStateChange
        tmrEnd.Enabled = True
    End If
End Sub
Public Sub CloseVideo()
    If ObjPtr(obj_MediaControl) > 0 Then
        obj_MediaControl.Stop
    End If
    'Destroy all objects
    If ObjPtr(obj_Audio) > 0 Then Set obj_Audio = Nothing
    If ObjPtr(obj_Video) > 0 Then Set obj_Video = Nothing
    If ObjPtr(obj_MediaControl) > 0 Then Set obj_MediaControl = Nothing
    If ObjPtr(obj_VideoWindow) > 0 Then Set obj_VideoWindow = Nothing
    If ObjPtr(obj_MediaPosition) > 0 Then Set obj_MediaPosition = Nothing
    If ObjPtr(obj_VideoWindow) > 0 Then
        obj_VideoWindow.Owner = 0
    End If
    PlayState = Stopped
    RaiseEvent PlayStateChange
End Sub

Public Sub PauseVideo()
    If ObjPtr(obj_MediaControl) > 0 Then
        Call obj_MediaControl.Pause
        PlayState = Paused
        RaiseEvent PlayStateChange
        tmrEnd.Enabled = False
    End If
End Sub

Public Sub StopVideo()
    If (ObjPtr(obj_MediaControl) > 0) And (ObjPtr(obj_MediaPosition) > 0) Then
        Call obj_MediaControl.Stop
        obj_MediaPosition.CurrentPosition = 0
        PlayState = Stopped
        RaiseEvent PlayStateChange
    End If
End Sub

Public Sub SetSize(Optional Left As Long, Optional Top As Long, Optional Width As Long, Optional Height As Long)
    On Error Resume Next
    If Left = vbNull Then Left = 0
    If Top = vbNull Then Top = 0
    If Width = vbNull Then Width = UserControl.ScaleWidth
    If Height = vbNull Then Height = UserControl.ScaleHeight
    obj_VideoWindow.SetWindowPosition Left, Top, Width, Height
End Sub
Private Sub tmrEnd_Timer()
    If (ObjPtr(obj_MediaPosition) > 0) Then
        If obj_MediaPosition.CurrentPosition = obj_MediaPosition.Duration Then
            RaiseEvent EndOfStream
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    lngBal = 0
    lngVol = 100
    dblRate = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If DisplayType <> HaftSize Then
        SetSize , , UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
    RaiseEvent Resize
End Sub

Private Sub UserControl_Terminate()
    Call CloseVideo
End Sub
Private Sub CalSize(Default As VideoSize)
    HaftS.Height = Default.Height \ 2
    HaftS.Width = Default.Width \ 2
    
    DoubleS.Height = Default.Height * 2
    DoubleS.Width = Default.Width * 2
    
    FullScreenS.Height = Screen.Height / Screen.TwipsPerPixelY
    FullScreenS.Width = Screen.Width / Screen.TwipsPerPixelX
End Sub

Public Sub HideCursor(Hide As Boolean)
    Select Case Hide
        Case False
            Do Until ShowCursor(True) > 0
                obj_VideoWindow.HideCursor Hide
            Loop
        Case True
            Do Until ShowCursor(False) < 0
                obj_VideoWindow.HideCursor Hide
            Loop
    End Select
End Sub

