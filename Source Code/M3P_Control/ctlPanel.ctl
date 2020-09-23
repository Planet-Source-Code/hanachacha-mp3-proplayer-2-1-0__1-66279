VERSION 5.00
Begin VB.UserControl ctlPanel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ControlContainer=   -1  'True
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ToolboxBitmap   =   "ctlPanel.ctx":0000
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "ctlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim lngReg As Long
Dim lngMaskColor As Long
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Public Sub Draw(pic As StdPicture, xS As Long, yS As Long, widthS As Long, heightS As Long)
    On Error Resume Next
    picSource.Move 0, 0
    UserControl.Cls
    picSource.PaintPicture pic, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, xS, yS, widthS, heightS, vbSrcCopy
    UserControl.PaintPicture pic, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, xS, yS, widthS, heightS, vbSrcCopy
    UserControl.Picture = UserControl.Image
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.Enabled = .ReadProperty("Enabled", True)
    End With
End Sub

Private Sub UserControl_Resize()
    picSource.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    UserControl.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
    End With
End Sub
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
Public Sub MakeGraph()
    lngReg = MakeRegion(picSource, lngMaskColor)
    SetWindowRgn UserControl.hwnd, lngReg, True
End Sub


