VERSION 5.00
Begin VB.UserControl DynamicButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   795
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   53
   ToolboxBitmap   =   "GraphButton.ctx":0000
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrTestState 
      Interval        =   10
      Left            =   360
      Top             =   0
   End
End
Attribute VB_Name = "DynamicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim pic As StdPicture
Dim PicOver As StdPicture
Dim PicDown As StdPicture

Dim PicOn As StdPicture
Dim PicOnOver As StdPicture
Dim PicOnDown As StdPicture

Public bolOn As Boolean
Dim bolOver As Boolean
Dim MouseState As Integer
Dim lngReg As Long
Dim lngMaskColor As Long
Event TurnOn()
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

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

Public Property Get PictureDown() As StdPicture
    Set PictureDown = PicDown
End Property

Public Property Set PictureDown(NewPic As StdPicture)
    Set PicDown = NewPic
    Draw
    PropertyChanged "PictureDown"
End Property

Public Property Get PictureOn() As StdPicture
    Set PictureOn = PicOn
End Property

Public Property Set PictureOn(NewPic As StdPicture)
    Set PicOn = NewPic
    Draw
    PropertyChanged "PictureOn"
End Property

Public Property Get PictureOnOver() As StdPicture
    Set PictureOnOver = PicOnOver
End Property

Public Property Set PictureOnOver(NewPic As StdPicture)
    Set PicOnOver = NewPic
    Draw
    PropertyChanged "PictureOnOver"
End Property

Public Property Get PictureOnDown() As StdPicture
    Set PictureOnDown = PicOnDown
End Property

Public Property Set PictureOnDown(NewPic As StdPicture)
    Set PicOnDown = NewPic
    Draw
    PropertyChanged "PictureOnDown"
End Property

Private Sub Draw()
    On Error Resume Next
        picSource.Move 0, 0
        If Not bolOn Then
            If MouseState = 2 Then
                UserControl.Cls
                UserControl.PaintPicture PicOver, 0, 0
                picSource.Picture = PicOver
                UserControl.Width = picSource.Width * Screen.TwipsPerPixelX
                UserControl.Height = picSource.Height * Screen.TwipsPerPixelY
            ElseIf MouseState = 1 Then
                UserControl.Cls
                UserControl.PaintPicture PicDown, 0, 0
                picSource.Picture = PicDown
                UserControl.Width = picSource.Width * Screen.TwipsPerPixelX
                UserControl.Height = picSource.Height * Screen.TwipsPerPixelY
            Else
                UserControl.Cls
                UserControl.PaintPicture pic, 0, 0
                picSource.Picture = pic
                UserControl.Width = picSource.Width * Screen.TwipsPerPixelX
                UserControl.Height = picSource.Height * Screen.TwipsPerPixelY
            End If
        Else
            If MouseState = 2 Then
                UserControl.Cls
                UserControl.PaintPicture PicOnOver, 0, 0
                picSource.Picture = PicOnOver
                UserControl.Width = picSource.Width * Screen.TwipsPerPixelX
                UserControl.Height = picSource.Height * Screen.TwipsPerPixelY
            ElseIf MouseState = 1 Then
                UserControl.Cls
                UserControl.PaintPicture PicOnDown, 0, 0
                picSource.Picture = PicOnDown
                UserControl.Width = picSource.Width * Screen.TwipsPerPixelX
                UserControl.Height = picSource.Height * Screen.TwipsPerPixelY
            Else
                UserControl.Cls
                UserControl.PaintPicture PicOn, 0, 0
                picSource.Picture = PicOn
                UserControl.Width = picSource.Width * Screen.TwipsPerPixelX
                UserControl.Height = picSource.Height * Screen.TwipsPerPixelY
            End If
        End If
End Sub

'
Private Sub tmrTestState_Timer()
    If Not TestOver Then
        tmrTestState.Enabled = False
        bolOver = False
        MouseState = 0
        Call Draw
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    Set pic = Nothing
    Set PicOver = Nothing
    Set PicDown = Nothing
    Set PicOn = Nothing
    Set PicOnOver = Nothing
    Set PicOnDown = Nothing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
    If Button = vbLeftButton Then
        MouseState = 1
        Call Draw
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
    If Button < 2 Then
        If Not TestOver Then
            bolOver = False
            MouseState = 0
            Call Draw
        Else
            If Button = 0 And Not bolOver Then
                tmrTestState.Enabled = True
                bolOver = True
                MouseState = 2
                Call Draw
            ElseIf Button = 1 Then
                bolOver = True
                MouseState = 1
                Call Draw
                bolOver = False
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
    If Button = vbLeftButton Then
        If TestOver Then
            MouseState = 2
            RaiseEvent TurnOn
            bolOn = Not bolOn
        Else
            MouseState = 0
        End If
        Call Draw
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set pic = .ReadProperty("Picture", Nothing)
        Set PicOver = .ReadProperty("PictureOver", Nothing)
        Set PicDown = .ReadProperty("PictureDown", Nothing)
        Set PicOn = .ReadProperty("PictureOn", Nothing)
        Set PicOnOver = .ReadProperty("PictureOnOver", Nothing)
        Set PicOnDown = .ReadProperty("PictureOnDown", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.Enabled = .ReadProperty("Enabled", True)
    End With
End Sub

Private Sub UserControl_Resize()
    Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Picture", pic, Nothing)
        Call .WriteProperty("PictureOver", PicOver, Nothing)
        Call .WriteProperty("PictureDown", PicDown, Nothing)
        Call .WriteProperty("PictureOn", PicOn, Nothing)
        Call .WriteProperty("PictureOnOver", PicOnOver, Nothing)
        Call .WriteProperty("PictureOnDown", PicOnDown, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
    End With
End Sub

Private Function TestOver() As Boolean
    Dim rad As POINTAPI
    GetCursorPos rad
    TestOver = (WindowFromPoint(rad.x, rad.Y) = UserControl.hwnd)
End Function
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
    UserControl.MousePointer = vNewValue
    PropertyChanged "MousePointer"
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal Img As Picture)
    Set UserControl.MouseIcon = Img
    PropertyChanged "MouseIcon"
End Property
Public Property Get MaskColor() As Long
    MaskColor = lngMaskColor
End Property
Public Property Let MaskColor(ByVal NewVal As Long)
    lngMaskColor = NewVal
    PropertyChanged "MaskColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Sub Refresh()
    UserControl.Cls
    'Draw picture
    Call Draw
End Sub
Public Sub MakeGraph()
    lngReg = MakeRegion(picSource, lngMaskColor)
    SetWindowRgn UserControl.hwnd, lngReg, True
End Sub

