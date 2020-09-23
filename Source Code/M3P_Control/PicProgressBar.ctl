VERSION 5.00
Begin VB.UserControl Progress 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ToolboxBitmap   =   "PicProgressBar.ctx":0000
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer tmrTestState 
      Interval        =   10
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Shape shpBorder 
      Height          =   1575
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum PBorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Enum ProgressStyle
    [Standard] = 0
    [Gradient] = 1
    [Graphical] = 2
End Enum

Enum Orientation
    [Horizontal] = 0
    [Vertical] = 1
End Enum

Dim lngMax As Long
Dim lngMin As Long
Dim lngValue As Long
Dim dblOneValue As Double
Dim bolActive As Boolean
Dim sngDown As Single

Dim Style As ProgressStyle
Dim Orien As Orientation
Dim bolBorder As PBorderStyle

Dim pic As StdPicture
Dim PicOver  As StdPicture
Dim picStatusN As StdPicture
Dim picStatusOver As StdPicture
Dim picStatusDown As StdPicture

Dim lngMinColor As Long
Dim lngMaxColor As Long

Dim bolOver As Boolean
Dim intState As Integer

Public Event Change(lValue As Long)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Property Get BorderStyle() As PBorderStyle
    BorderStyle = bolBorder 'UserControl.BorderStyle
End Property
Public Property Let BorderStyle(New_Val As PBorderStyle)
    bolBorder = New_Val
    If Style = Graphical Then bolBorder = None
    If bolBorder = [Fixed Single] Then
        shpBorder.Visible = True
    Else
        shpBorder.Visible = False
    End If
    If Orien = Vertical Then
        If bolBorder = None Then
            picStatus.Height = UserControl.ScaleHeight
            picStatus.Width = UserControl.ScaleWidth
            picStatus.Left = 0
        Else
            picStatus.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        End If
    End If
    If Orien = Horizontal Then
        If bolBorder = None Then
            picStatus.Width = UserControl.ScaleWidth
            picStatus.Height = UserControl.ScaleHeight
            picStatus.top = 0
        Else
            picStatus.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        End If
    End If
    PropertyChanged "BorderStyle"
End Property
Public Property Get ProgressStyle() As ProgressStyle
    ProgressStyle = Style
End Property
Public Property Let ProgressStyle(New_Val As ProgressStyle)
    Style = New_Val
    If Style = Graphical Then bolBorder = None
    Call Draw
    PropertyChanged "ProgressStyle"
End Property
Public Property Get Orientation() As Orientation
    Orientation = Orien
End Property
Public Property Let Orientation(New_Val As Orientation)
    Orien = New_Val
    Call Draw
    PropertyChanged "Orientation"
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

Public Property Get PictureStatus() As StdPicture
    Set PictureStatus = picStatusN
End Property

Public Property Set PictureStatus(NewPic As StdPicture)
    Set picStatusN = NewPic
    Draw
    PropertyChanged "PictureStatus"
End Property

Public Property Get PictureStatusDown() As StdPicture
    Set PictureStatusDown = picStatusDown
End Property

Public Property Set PictureStatusDown(NewPic As StdPicture)
    Set picStatusDown = NewPic
    Draw
    PropertyChanged "PictureStatusDown"
End Property
Public Property Get PictureStatusOver() As StdPicture
    Set PictureStatusOver = picStatusOver
End Property

Public Property Set PictureStatusOver(NewPic As StdPicture)
    Set picStatusOver = NewPic
    Draw
    PropertyChanged "PictureStatusOver"
End Property

Private Sub picStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        bolActive = True
        If Orien = Horizontal Then sngDown = x / Screen.TwipsPerPixelX
        If Orien = Vertical Then sngDown = Y / Screen.TwipsPerPixelY
        Call Draw
    End If
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub
Private Sub picStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bolActive And Button = vbLeftButton Then
        If Orien = Horizontal Then
            picStatus.Left = picStatus.Left + (x / Screen.TwipsPerPixelX - sngDown)
            If picStatus.Left > 0 Then picStatus.Left = 0
            If picStatus.Left < -picStatus.Width Then picStatus.Left = -picStatus.Width
                dblOneValue = picStatus.Height / (lngMax - lngMin)
                lngValue = (picStatus.Left + picStatus.Width) / dblOneValue + lngMin
        End If
        If Orien = Vertical Then
            picStatus.top = picStatus.top + (Y / Screen.TwipsPerPixelY - sngDown)
            If picStatus.top < 0 Then picStatus.top = 0
            If picStatus.top > picStatus.Height Then picStatus.top = picStatus.Height
                dblOneValue = picStatus.Height / (lngMax - lngMin)
                lngValue = (picStatus.top + picStatus.Height) / dblOneValue + lngMin
        End If
        Call Draw
        RaiseEvent Change(lngValue)
    End If
End Sub
Private Sub picStatus_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And bolActive Then
        bolActive = False
        Call Draw
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


Private Sub UserControl_InitProperties()
    If UserControl.ScaleHeight < UserControl.ScaleWidth Then
        Orien = Horizontal
    Else
        Orien = Vertical
    End If
    lngMin = 0
    lngMax = 100
    lngValue = 0
    Style = Standard
    Set pic = Nothing
    Set PicOver = Nothing
    Set picStatus = Nothing
    Set picStatusDown = Nothing
    Set picStatusOver = Nothing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        bolActive = True
    End If
    intState = 2
    Call Draw
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bolActive Then
        If Orien = Horizontal Then
            picStatus.Left = x - picStatus.Width
            If picStatus.Left > 0 Then picStatus.Left = 0
            If picStatus.Left < -picStatus.Width Then picStatus.Left = -picStatus.Width
            dblOneValue = picStatus.Width / (lngMax - lngMin)
            lngValue = (picStatus.Left + picStatus.Width) / dblOneValue + lngMin
        End If
        If Orien = Vertical Then
            picStatus.top = Y
            If picStatus.top < 0 Then picStatus.top = 0
            If picStatus.top > picStatus.Height Then picStatus.top = picStatus.Height
            dblOneValue = picStatus.Height / (lngMax - lngMin)
            lngValue = (picStatus.Height - picStatus.top) / dblOneValue + lngMin
        End If
        RaiseEvent Change(lngValue)
        Call Draw
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
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
    If Button = vbLeftButton Then
        If bolActive Then bolActive = False
        If TestOver Then
            intState = 2
        Else
            intState = 0
        End If
        Call Draw
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lngMax = .ReadProperty("Max", 100)
        lngMin = .ReadProperty("Min", 0)
        lngValue = .ReadProperty("Value", 0)
        
        UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        shpBorder.BorderColor = .ReadProperty("BorderColor", vbGreen)
        Orien = .ReadProperty("Orientation", 0)
        Style = .ReadProperty("ProgressStyle", 0)
        bolBorder = .ReadProperty("BorderStyle", 1)
        
        If bolBorder = [Fixed Single] Then
            shpBorder.Visible = True
        Else
            shpBorder.Visible = False
        End If
        
        Set pic = .ReadProperty("Picture", Nothing)
        Set PicOver = .ReadProperty("PictureOver", Nothing)
        Set picStatusN = .ReadProperty("PictureStatus", Nothing)
        Set picStatusDown = .ReadProperty("PictureStatusDown", Nothing)
        Set picStatusOver = .ReadProperty("PictureStatusOver", Nothing)
        
        lngMinColor = .ReadProperty("MinColor", vbBlue)
        lngMaxColor = .ReadProperty("MaxColor", vbWhite)
        
        picStatus.BackColor = .ReadProperty("ForeColor", vbRed)
        picStatus.MousePointer = .ReadProperty("MousePointer", vbDefault)
        picStatus.MouseIcon = .ReadProperty("MouseIcon", picStatus.MouseIcon)
    End With
    If Style = Gradient Then
        If Orien = Horizontal Then
            FillGradient picStatus.hDC, 0, 0, picStatus.Width, picStatus.Height, lngMaxColor, lngMinColor, Fill_Horizontal
        End If
        If Orien = Vertical Then
            FillGradient picStatus.hDC, 0, 0, picStatus.Width, picStatus.Height, lngMinColor, lngMaxColor, Fill_Vertical
        End If
        picStatus.Picture = picStatus.Image
    Else
        If Style = Standard Then
            picStatus.Cls
            Set picStatus.Picture = Nothing
        End If
    End If
    Draw
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If Orien = Vertical Then
        If bolBorder = None Then
            picStatus.Height = UserControl.ScaleHeight
            picStatus.Width = UserControl.ScaleWidth
            picStatus.Left = 0
        Else
            picStatus.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        End If
        dblOneValue = picStatus.Height / (lngMax - lngMin)
    End If
    If Orien = Horizontal Then
        If bolBorder = None Then
            picStatus.Width = UserControl.ScaleWidth
            picStatus.Height = UserControl.ScaleHeight
            picStatus.top = 0
        Else
            picStatus.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        End If
        dblOneValue = picStatus.Width / (lngMax - lngMin)
    End If
    If Style = Gradient Then
        If Orien = Horizontal Then
            FillGradient picStatus.hDC, 0, 0, picStatus.Width, picStatus.Height, lngMaxColor, lngMinColor, Fill_Horizontal
        End If
        If Orien = Vertical Then
            FillGradient picStatus.hDC, 0, 0, picStatus.Width, picStatus.Height, lngMinColor, lngMaxColor, Fill_Vertical
        End If
        picStatus.Picture = picStatus.Image
    Else
        If Style = Standard Then
            picStatus.Cls
            Set picStatus.Picture = Nothing
        End If
    End If
    Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Max", lngMax, 100)
        Call .WriteProperty("Min", lngMin, 0)
        Call .WriteProperty("Value", lngValue, 50)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("BackColor", UserControl.BackColor, vbWhite)
        Call .WriteProperty("BorderColor", shpBorder.BorderColor, vbGreen)
        Call .WriteProperty("ForeColor", picStatus.BackColor, vbGrayed)
        Call .WriteProperty("MaxColor", lngMaxColor, vbWhite)
        Call .WriteProperty("MinColor", lngMinColor, vbBlue)
        Call .WriteProperty("BorderStyle", bolBorder, 1)
        Call .WriteProperty("Orientation", Orien, 0)
        Call .WriteProperty("ProgressStyle", Style, 0)
        Call .WriteProperty("Picture", pic)
        Call .WriteProperty("PictureOver", PicOver)
        Call .WriteProperty("PictureStatus", picStatusN)
        Call .WriteProperty("PictureStatusDown", picStatusDown)
        Call .WriteProperty("PictureStatusOver", picStatusOver)
        Call .WriteProperty("MousePointer", picStatus.MousePointer, vbDefault)
        Call .WriteProperty("MouseIcon", picStatus.MouseIcon)
    End With
End Sub
Public Property Get Max() As Long
    Max = lngMax
End Property
Public Property Let Max(ByVal lngValueMax As Long)
    If lngValueMax <> lngMin Then
        lngMax = lngValueMax
        PropertyChanged "Max"
    End If
End Property
Public Property Get Min() As Long
    Min = lngMin
End Property
Public Property Let Min(ByVal lngValueMin As Long)
    If lngValueMin <> lngMax Then
        lngMin = lngValueMin
        PropertyChanged "Min"
    End If
End Property
Public Property Get Value() As Long
    Value = lngValue
End Property
Public Property Let Value(ByVal lngCurrentValue As Long)
    If lngMax > lngMin Then
        If lngCurrentValue >= lngMin And lngCurrentValue <= lngMax Then lngValue = lngCurrentValue
    Else
        If lngCurrentValue <= lngMin And lngCurrentValue >= lngMax Then lngValue = lngCurrentValue
    End If
    If Not bolActive Then
        CalSize lngValue
        Draw
        PropertyChanged "Value"
    End If
End Property
Private Sub CalSize(lValue As Long)
    If Orien = Vertical Then
        dblOneValue = picStatus.Height / (lngMax - lngMin)
    End If
    If Orien = Horizontal Then
        dblOneValue = picStatus.Width / (lngMax - lngMin)
    End If
    If Not bolActive Then
        If Orien = Vertical Then
            If bolBorder = None Then
                picStatus.top = picStatus.Height - ((lngValue - lngMin) * dblOneValue)
            Else
                picStatus.top = picStatus.Height - ((lngValue - lngMin) * dblOneValue) - 1
            End If
        End If
        If Orien = Horizontal Then
            If bolBorder = None Then
                picStatus.Left = (lngValue - lngMin) * dblOneValue - picStatus.Width
            Else
                picStatus.Left = (lngValue - lngMin) * dblOneValue - picStatus.Width + 1
            End If
        End If
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
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = shpBorder.BorderColor
End Property
Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
    shpBorder.BorderColor = NewValue
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picStatus.BackColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    picStatus.BackColor = NewValue
End Property
Public Property Get MaxColor() As OLE_COLOR
    MaxColor = lngMaxColor
End Property
Public Property Let MaxColor(ByVal NewValue As OLE_COLOR)
    lngMaxColor = NewValue
End Property
Public Property Get MinColor() As OLE_COLOR
    MinColor = lngMinColor
End Property
Public Property Let MinColor(ByVal NewValue As OLE_COLOR)
    lngMinColor = NewValue
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = picStatus.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    picStatus.MousePointer = NewValue
    PropertyChanged "MousePointer"
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = picStatus.MouseIcon
End Property
Public Property Set MouseIcon(ByVal Img As Picture)
    Set picStatus.MouseIcon = Img
    PropertyChanged "MouseIcon"
End Property
Private Sub Draw()
    On Error Resume Next
    Dim i As Long
    
    If Orien = Horizontal Then
        Select Case Style
            Case Is = Graphical
                picStatus.Visible = False
                picStatus.Width = UserControl.ScaleWidth
                Call CalSize(lngValue)
                If intState = 2 Then
                    picBack.Picture = PicOver
                    picStatus.Picture = picStatusOver
                    UserControl.Cls
                    UserControl.Height = picBack.ScaleHeight * Screen.TwipsPerPixelY
                    For i = 0 To UserControl.ScaleWidth Step picBack.Width
                        UserControl.PaintPicture PicOver, i, 0
                    Next i
                    UserControl.PaintPicture picStatus, 0, picStatus.top, (picStatus.Width + picStatus.Left), picStatus.Height, 0, 0, (picStatus.Width + picStatus.Left), picStatus.Height
                ElseIf intState = 1 Then
                    UserControl.Cls
                    UserControl.Height = picBack.ScaleHeight * Screen.TwipsPerPixelY
                    For i = 0 To UserControl.ScaleWidth Step picBack.Width
                        UserControl.PaintPicture PicOver, i, 0
                    Next i
                    picStatus.Picture = picStatusDown
                    UserControl.PaintPicture picStatus, 0, picStatus.top, (picStatus.Width + picStatus.Left), picStatus.Height, 0, 0, (picStatus.Width + picStatus.Left), picStatus.Height
                    Else
                        picBack.Picture = pic
                        picStatus.Picture = picStatusN
                        UserControl.Cls
                        UserControl.Height = picBack.ScaleHeight * Screen.TwipsPerPixelY
                        For i = 0 To UserControl.ScaleWidth Step picBack.Width
                            UserControl.PaintPicture pic, i, 0
                        Next i
                        UserControl.PaintPicture picStatus, 0, picStatus.top, (picStatus.Width + picStatus.Left), picStatus.Height, 0, 0, (picStatus.Width + picStatus.Left), picStatus.Height
                End If
            Case Is = Standard
                Call CalSize(lngValue)
                UserControl.Cls
                If bolBorder = [Fixed Single] Then
                    UserControl.Line (1, 1)-(picStatus.Width + picStatus.Left, UserControl.ScaleHeight - 1), picStatus.BackColor, BF
                Else
                    UserControl.Line (0, 0)-(picStatus.Width + picStatus.Left, UserControl.ScaleHeight), picStatus.BackColor, BF
                End If
            Case Is = Gradient
                Call CalSize(lngValue)
                UserControl.Cls
                If bolBorder = [Fixed Single] Then
                    UserControl.PaintPicture picStatus, 1, 1, (picStatus.Width + picStatus.Left), UserControl.ScaleHeight - 1, 0, 0, (picStatus.Width + picStatus.Left), picStatus.Height
                Else
                    UserControl.PaintPicture picStatus, 0, 0, (picStatus.Width + picStatus.Left), UserControl.ScaleHeight, 0, 0, (picStatus.Width + picStatus.Left), picStatus.Height
                End If
        End Select
    End If
    If Orien = Vertical Then
        Select Case Style
            Case Is = Graphical
                picStatus.Visible = False
                picStatus.Height = UserControl.ScaleHeight
                Call CalSize(lngValue)
                If intState = 2 Then
                    picBack.Picture = PicOver
                    UserControl.Cls
                    UserControl.Width = picBack.Width * Screen.TwipsPerPixelX
                    For i = 0 To UserControl.ScaleHeight Step picBack.Height
                        UserControl.PaintPicture PicOver, 0, i
                    Next i
                    picStatus.Picture = picStatusOver
                    UserControl.PaintPicture picStatus, picStatus.Left, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top, 0, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top
                ElseIf intState = 1 Then
                    UserControl.Cls
                    UserControl.Width = picBack.Width * Screen.TwipsPerPixelX
                    For i = 0 To UserControl.ScaleHeight Step picBack.Height
                        UserControl.PaintPicture PicOver, 0, i
                    Next i
                    picStatus.Picture = picStatusDown
                    UserControl.PaintPicture picStatus, picStatus.Left, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top, 0, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top
                Else
                    picBack.Picture = pic
                    UserControl.Cls
                    UserControl.Width = picBack.Width * Screen.TwipsPerPixelX
                    For i = 0 To UserControl.ScaleHeight Step picBack.Height
                        UserControl.PaintPicture pic, 0, i
                    Next i
                    picStatus.Picture = picStatusN
                    UserControl.PaintPicture picStatus, picStatus.Left, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top, 0, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top
                End If
            Case Is = Standard
                Call CalSize(lngValue)
                UserControl.Cls
                If bolBorder = [Fixed Single] Then
                    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)-(1, picStatus.top + 1), picStatus.BackColor, BF
                Else
                    UserControl.Line (UserControl.ScaleWidth, UserControl.ScaleHeight)-(0, picStatus.top), picStatus.BackColor, BF
                End If
            Case Is = Gradient
                Call CalSize(lngValue)
                UserControl.Cls
                If bolBorder = [Fixed Single] Then
                    UserControl.PaintPicture picStatus, 1, picStatus.top + 1, picStatus.Width, picStatus.Height - picStatus.top, 0, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top
                Else
                    UserControl.PaintPicture picStatus, 0, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top, 0, picStatus.top, picStatus.Width, picStatus.Height - picStatus.top
                End If
        End Select
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

