VERSION 5.00
Begin VB.UserControl ctlSlider 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ToolboxBitmap   =   "ctlSlider.ctx":0000
   Begin VB.PictureBox Cue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer tmrTestState 
      Interval        =   10
      Left            =   0
      Top             =   3120
   End
   Begin VB.PictureBox picSlider 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "ctlSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum SliderStyle
    [Horizontal] = 0
    [Vertical] = 1
End Enum

Enum HBorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Enum DrawStyle
    [Standard] = 0
    [Graphic] = 1
End Enum

Enum AutoSizeCue
    [Auto] = 0
    [Normal] = 1
End Enum

Dim lngMax As Long
Dim lngMin As Long
Dim lngValue As Long
Dim bolActive As Boolean
Dim bolOver As Boolean
Dim intState As Integer
Dim sMDownX As Single
Dim sMDownY As Single
Dim sMRad As Single
Dim dblOneValue As Double

Dim m_AutoSize As AutoSizeCue
Dim m_Style As DrawStyle
Dim m_Slider As SliderStyle

Dim pic As StdPicture
Dim PicOver  As StdPicture
Dim PicCue As StdPicture
Dim PicCueDown As StdPicture
Dim PicCueOver As StdPicture

Public Event Change(lValue As Long)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Property Get BorderStyle() As HBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(New_Val As HBorderStyle)
    UserControl.BorderStyle = New_Val
    PropertyChanged "BorderStyle"
End Property
Public Property Get CueHeight() As Long
    CueHeight = Cue.Height
End Property
Public Property Let CueHeight(New_Val As Long)
    Cue.Height = New_Val
    PropertyChanged "CueHeight"
End Property
Public Property Get CueWidth() As Long
    CueWidth = Cue.Width
End Property
Public Property Let CueWidth(New_Val As Long)
    Cue.Width = New_Val
    PropertyChanged "CueWidth"
End Property
Public Property Get SliderOrien() As SliderStyle
    SliderOrien = m_Slider
End Property
Public Property Let SliderOrien(New_Val As SliderStyle)
    m_Slider = New_Val
    Draw
    PropertyChanged "SliderOrien"
End Property
Public Property Get Style() As DrawStyle
    Style = m_Style
End Property
Public Property Let Style(New_Val As DrawStyle)
    m_Style = New_Val
    PropertyChanged "Style"
End Property
Public Property Get AutoCue() As AutoSizeCue
    AutoCue = m_AutoSize
End Property
Public Property Let AutoCue(New_Val As AutoSizeCue)
    m_AutoSize = New_Val
    PropertyChanged "AutoCue"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = pic
End Property

Public Property Set Picture(NewPic As StdPicture)
    Set pic = NewPic
    Draw
    PropertyChanged "Picture"
End Property

Public Property Get PictureOver() As StdPicture
    Set PictureOver = PicOver
End Property

Public Property Set PictureOver(NewPic As StdPicture)
    Set PicOver = NewPic
    Draw
    PropertyChanged "PictureOver"
End Property

Public Property Get PictureCue() As StdPicture
    Set PictureCue = PicCue
End Property

Public Property Set PictureCue(NewPic As StdPicture)
    Set PicCue = NewPic
    Draw
    PropertyChanged "PictureCue"
End Property

Public Property Get PictureCueDown() As StdPicture
    Set PictureCueDown = PicCueDown
End Property

Public Property Set PictureCueDown(NewPic As StdPicture)
    Set PicCueDown = NewPic
    Draw
    PropertyChanged "PictureCueDown"
End Property
Public Property Get PictureCueOver() As StdPicture
    Set PictureCueOver = PicCueOver
End Property

Public Property Set PictureCueOver(NewPic As StdPicture)
    Set PicCueOver = NewPic
    Draw
    PropertyChanged "PictureCueOver"
End Property
Private Sub Cue_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        bolActive = True
        If m_Slider = Horizontal Then
            sMDownX = x / Screen.TwipsPerPixelX
        Else
            sMDownY = Y / Screen.TwipsPerPixelY
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub
Private Sub Cue_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If (Button = vbLeftButton) And bolActive Then
        If m_Slider = Horizontal Then
            dblOneValue = (UserControl.ScaleWidth - Cue.Width) / (lngMax - lngMin)
            sMRad = Cue.Left + x / Screen.TwipsPerPixelX
            Cue.Left = sMRad - sMDownX
            If Cue.Left < 0 Then Cue.Left = 0
            If Cue.Left > UserControl.ScaleWidth - Cue.Width Then Cue.Left = (UserControl.ScaleWidth - Cue.Width)
            lngValue = lngMin + (Cue.Left / dblOneValue)
        Else
            dblOneValue = (UserControl.ScaleHeight - Cue.Height) / (lngMax - lngMin)
            sMRad = Cue.top + Y / Screen.TwipsPerPixelY
            Cue.top = sMRad - sMDownY
            If Cue.top < 0 Then Cue.top = 0
            If Cue.top > UserControl.ScaleHeight - Cue.Height Then Cue.top = (UserControl.ScaleHeight - Cue.Height)
            lngValue = lngMin + (Cue.top / dblOneValue)
        End If
        RaiseEvent Change(lngValue)
    End If
End Sub
Private Sub Cue_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bolActive And (Button = vbLeftButton) Then
        bolActive = False
    End If
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub tmrTestState_Timer()
    If Not TestOver Then
        tmrTestState.Enabled = False
        bolOver = False
        intState = 0
        Call Draw
    End If
End Sub

Private Sub UserControl_Initialize()
    lngMin = 0
    lngMax = 100
    lngValue = 0
    picSlider.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If UserControl.ScaleHeight > UserControl.ScaleWidth Then
        m_Slider = Vertical
        PropertyChanged "SliderOrien"
    Else
        m_Slider = Horizontal
        PropertyChanged "SliderOrien"
    End If
End Sub

Private Sub UserControl_InitProperties()
    Set pic = Nothing
    Set PicOver = Nothing
    Set PicCue = Nothing
    Set PicCueDown = Nothing
    Set PicCueOver = Nothing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        bolActive = True
        If m_Slider = Horizontal Then
            If x > Cue.Left And x <= Cue.Width + Cue.Left Then
                intState = 1
            Else
                intState = 2
            End If
        Else
            If Y > Cue.top And Y <= Cue.Height + Cue.top Then
                intState = 1
            Else
                intState = 2
            End If
        End If
        Call Draw
    End If
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        If m_Slider = Horizontal Then
            dblOneValue = (UserControl.ScaleWidth - Cue.Width) / (lngMax - lngMin)
            Cue.Left = x - Cue.Width / 2
            If Cue.Left < 0 Then Cue.Left = 0
            If Cue.Left > UserControl.ScaleWidth - Cue.Width Then Cue.Left = (UserControl.ScaleWidth - Cue.Width)
            lngValue = (Cue.Left / dblOneValue) + lngMin
        Else
            dblOneValue = (UserControl.ScaleHeight - Cue.Height) / (lngMax - lngMin)
            Cue.top = Y - Cue.Height / 2
            If Cue.top < 0 Then Cue.top = 0
            If Cue.top > UserControl.ScaleHeight - Cue.Height Then Cue.top = (UserControl.ScaleHeight - Cue.Height)
            lngValue = (Cue.top / dblOneValue) + lngMin
        End If
        
        RaiseEvent Change(lngValue)
        Draw
    End If
    If Button < 2 Then
        If Not TestOver Then
            bolOver = False
            intState = 0
            Call Draw
        Else
            If Button = 0 And Not bolOver Then
                tmrTestState.Enabled = True
                bolOver = True
                intState = 2
                Call Draw
            ElseIf Button = 1 Then
                bolOver = True
                intState = 1
                Call Draw
                bolOver = False
            End If
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        bolActive = False
        If m_Slider = Horizontal Then
            dblOneValue = (UserControl.ScaleWidth - Cue.Width) / (lngMax - lngMin)
            Cue.Left = x - Cue.Width / 2
            If Cue.Left < 0 Then Cue.Left = 0
            If Cue.Left > UserControl.ScaleWidth - Cue.Width Then Cue.Left = (UserControl.ScaleWidth - Cue.Width)
            lngValue = (Cue.Left / dblOneValue) + lngMin
        Else
            dblOneValue = (UserControl.ScaleHeight - Cue.Height) / (lngMax - lngMin)
            Cue.top = Y - Cue.Height / 2
            If Cue.top < 0 Then Cue.top = 0
            If Cue.top > UserControl.ScaleHeight - Cue.Height Then Cue.top = (UserControl.ScaleHeight - Cue.Height)
            lngValue = (Cue.top / dblOneValue) + lngMin
        End If
        RaiseEvent Change(lngValue)
        If TestOver Then
            intState = 2
        Else
            intState = 0
        End If
        Call Draw
    End If
    
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lngMax = .ReadProperty("Max", 100)
        lngMin = .ReadProperty("Min", 0)
        lngValue = .ReadProperty("Value", 0)
        UserControl.BackColor = .ReadProperty("BackColor", 1)
        UserControl.BorderStyle = .ReadProperty("BorderStyle", 1)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        m_Slider = .ReadProperty("SliderOrien", 0)
        m_Style = .ReadProperty("Style", 0)
        m_AutoSize = .ReadProperty("AutoCue", 0)
        Cue.Height = .ReadProperty("CueHeight", UserControl.ScaleHeight)
        Cue.Width = .ReadProperty("CueWidth", 50)
        Set pic = .ReadProperty("Picture", Nothing)
        Set PicOver = .ReadProperty("PictureOver", Nothing)
        Set PicCue = .ReadProperty("PictureCue", Nothing)
        Set PicCueDown = .ReadProperty("PictureCueDown", Nothing)
        Set PicCueOver = .ReadProperty("PictureCueOver", Nothing)
        Cue.BackColor = .ReadProperty("ForeColor", vbRed)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", UserControl.MouseIcon)
    End With
    Draw
End Sub

Private Sub UserControl_Resize()
    If m_Slider = Horizontal Then
        If (UserControl.ScaleWidth - Cue.Width) = 0 Then Exit Sub
        dblOneValue = (UserControl.ScaleWidth - Cue.Width) / (lngMax - lngMin)
    Else
        If (UserControl.ScaleHeight - Cue.Height) = 0 Then Exit Sub
        dblOneValue = (UserControl.ScaleHeight - Cue.Height) / (lngMax - lngMin)
    End If
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
            Call .WriteProperty("Max", lngMax, 100)
            Call .WriteProperty("Min", lngMin, 0)
            Call .WriteProperty("Value", lngValue, 0)
            Call .WriteProperty("Enabled", UserControl.Enabled, True)
            Call .WriteProperty("BackColor", UserControl.BackColor, vbBlue)
            Call .WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
            Call .WriteProperty("SliderOrien", 0)
            Call .WriteProperty("Style", m_Style, 0)
            Call .WriteProperty("AutoCue", m_AutoSize, 0)
            Call .WriteProperty("CueWidth", Cue.Width)
            Call .WriteProperty("CueHeight", Cue.Height)
            Call .WriteProperty("Picture", pic)
            Call .WriteProperty("PictureOver", PicOver)
            Call .WriteProperty("PictureCue", PicCue)
            Call .WriteProperty("PictureCueDown", PicCueDown)
            Call .WriteProperty("PictureCueOver", PicCueOver)
            Call .WriteProperty("ForeColor", Cue.BackColor, vbGrayed)
            Call .WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
            Call .WriteProperty("MouseIcon", UserControl.MouseIcon)
    End With
End Sub
Public Property Get Max() As Long
    Max = lngMax
End Property

Public Property Let Max(ByVal lngValueMax As Long)
    If (lngValueMax <> lngMin) Then lngMax = lngValueMax
    Draw
    PropertyChanged "Max"
End Property
Public Property Get Min() As Long
    Min = lngMin
End Property
Public Property Let Min(ByVal lngValueMin As Long)
    If (lngValueMin <> lngMax) Then lngMin = lngValueMin
    Draw
    PropertyChanged "Min"
End Property
Public Property Get Value() As Long
    Value = lngValue
End Property
Public Property Let Value(ByVal lngCurrentValue As Long)
    If lngMax > lngMin Then
        If (lngCurrentValue >= lngMin) And (lngCurrentValue <= lngMax) Then lngValue = lngCurrentValue
    End If
    If lngMax < lngMin Then
        If (lngCurrentValue <= lngMin) And (lngCurrentValue >= lngMax) Then lngValue = lngCurrentValue
    End If
    lngValue = lngCurrentValue
    Draw
    PropertyChanged "Value"
End Property
Private Sub CalLeft(lngValueChange As Long)
    If Not bolActive Then
        dblOneValue = (UserControl.ScaleWidth - Cue.Width) / (lngMax - lngMin)
        Cue.Left = (lngValueChange - lngMin) * dblOneValue
    End If
End Sub
Private Sub CalTop(lngValueChange As Long)
    If Not bolActive Then
        dblOneValue = (UserControl.ScaleHeight - Cue.Height) / (lngMax - lngMin)
        Cue.top = (lngValueChange - lngMin) * dblOneValue
    End If
End Sub
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    UserControl.BackColor = NewValue
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Cue.BackColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    Cue.BackColor = NewValue
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
    PropertyChanged "MousePointer"
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal Img As Picture)
    Set UserControl.MouseIcon = Img
    PropertyChanged "MouseIcon"
End Property
Private Sub Draw()
    On Error Resume Next
    Dim i As Long
    If m_Slider = Horizontal Then
        If m_Style = Graphic Then
            Cue.Visible = False
            CalLeft lngValue
            If intState = 2 Then
                picSlider.Picture = PicOver
                UserControl.Cls
                UserControl.Height = picSlider.Height * Screen.TwipsPerPixelY
                For i = 0 To UserControl.ScaleWidth Step picSlider.Width
                    UserControl.PaintPicture PicOver, i, 0
                Next i
                Cue.Picture = PicCueOver
                Cue.top = (UserControl.ScaleHeight - Cue.Height) / 2
                UserControl.PaintPicture Cue, Cue.Left, Cue.top, Cue.Width, Cue.Height, 0, 0, Cue.Width, Cue.Height
            ElseIf intState = 1 Then
                UserControl.Cls
                UserControl.Height = picSlider.Height * Screen.TwipsPerPixelY
                For i = 0 To UserControl.ScaleWidth Step picSlider.Width
                    UserControl.PaintPicture PicOver, i, 0
                Next i
                Cue.Picture = PicCueDown
                Cue.top = (UserControl.ScaleHeight - Cue.Height) / 2
                UserControl.PaintPicture Cue, Cue.Left, Cue.top, Cue.Width, Cue.Height, 0, 0, Cue.Width, Cue.Height
            Else
                picSlider.Picture = pic
                UserControl.Cls
                UserControl.Height = picSlider.Height * Screen.TwipsPerPixelY
                For i = 0 To UserControl.ScaleWidth Step picSlider.Width
                    UserControl.PaintPicture pic, i, 0
                Next i
                Cue.Picture = PicCue
                Cue.top = (UserControl.ScaleHeight - Cue.Height) / 2
                UserControl.PaintPicture Cue, Cue.Left, Cue.top, Cue.Width, Cue.Height, 0, 0, Cue.Width, Cue.Height
            End If
        Else
            Cue.Visible = True
            Cue.top = 0
            If m_AutoSize = Auto Then
            Cue.Height = UserControl.ScaleHeight
                If lngMax > lngMin Then
                    Cue.Width = (UserControl.ScaleWidth) \ (lngMax - lngMin)
                Else
                    Cue.Width = (UserControl.ScaleWidth) \ (lngMin - lngMax)
                End If
            Else
                Cue.Width = Cue.Width
                Cue.Height = Cue.Height
            End If
            CalLeft lngValue
        End If
    Else
        If m_Style = Graphic Then
            Cue.Visible = False
            CalTop lngValue
            If intState = 2 Then
                picSlider.Picture = PicOver
                UserControl.Cls
                UserControl.Width = picSlider.Width * Screen.TwipsPerPixelX
                For i = 0 To UserControl.ScaleHeight Step picSlider.Height
                    UserControl.PaintPicture PicOver, 0, i
                Next i
                Cue.Picture = PicCueOver
                Cue.Left = (UserControl.ScaleWidth - Cue.Width) / 2
                UserControl.PaintPicture Cue, Cue.Left, Cue.top, Cue.Width, Cue.Height, 0, 0, Cue.Width, Cue.Height
            ElseIf intState = 1 Then
                UserControl.Cls
                UserControl.Width = picSlider.Width * Screen.TwipsPerPixelX
                For i = 0 To UserControl.ScaleHeight Step picSlider.Height
                    UserControl.PaintPicture PicOver, 0, i
                Next i
                Cue.Picture = PicCueDown
                Cue.Left = (UserControl.ScaleWidth - Cue.Width) / 2
                UserControl.PaintPicture Cue, Cue.Left, Cue.top, Cue.Width, Cue.Height, 0, 0, Cue.Width, Cue.Height
            Else
                picSlider.Picture = pic
                UserControl.Cls
                UserControl.Width = picSlider.Width * Screen.TwipsPerPixelX
                For i = 0 To UserControl.ScaleHeight Step picSlider.Height
                    UserControl.PaintPicture pic, 0, i
                Next i
                Cue.Picture = PicCue
                Cue.Left = (UserControl.ScaleWidth - Cue.Width) / 2
                UserControl.PaintPicture Cue, Cue.Left, Cue.top, Cue.Width, Cue.Height, 0, 0, Cue.Width, Cue.Height
            End If
        Else
            Cue.Visible = True
            Cue.Left = 0
            If m_AutoSize = Auto Then
            Cue.Width = UserControl.ScaleWidth
                If lngMax > lngMin Then
                    Cue.Height = (UserControl.ScaleHeight) \ (lngMax - lngMin)
                Else
                    Cue.Height = (UserControl.ScaleHeight) \ (lngMin - lngMax)
                End If
            Else
                Cue.Width = Cue.Width
                Cue.Height = Cue.Height
            End If
            CalTop lngValue
        End If
    End If
    
End Sub
Private Function TestOver() As Boolean
    Dim rad As POINTAPI
    GetCursorPos rad
    TestOver = (WindowFromPoint(rad.x, rad.Y) = UserControl.hwnd)
End Function

Public Sub Refresh()
    Call Draw
End Sub


