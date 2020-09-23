VERSION 5.00
Begin VB.UserControl Visualization 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ToolboxBitmap   =   "Visualization.ctx":0000
   Begin VB.PictureBox picOsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picPeak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Visualization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetPixel Lib "gdi32" _
                                        (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, _
                                        ByVal crColor As Long) As Long



Private Type Spectrum
    DrawMode As Integer '0 is thick  ; 1 is thin
    FillMode As Integer '0 is normal;1 is fire;2 is line
    ShowPeak As Boolean
    PeakDelay As Long
    PeakDrop As Long
End Type

Private Type Oscilliscope
    DrawMode As Integer '0 is Dot;1 is Line;2 is Solid
End Type

Private tOsc As Oscilliscope
Private tSpec As Spectrum

Public intStyle As Integer '0 is Oscilliscope;1 is Spectrum;2 is none


Private Type Visualization
    Color1 As Long
    Color2 As Long
    Color3 As Long
    Pict As String
    ColorPeak As Long
    ColorOsc1 As Long
    ColorOsc2 As Long
    ColorOsc3 As Long
End Type
Dim tVisual As Visualization



Dim data(0 To 255) As Single  ' Stores y position for BitBlt
Dim Data2(0 To 255) As Single
Dim Data3(0 To 255) As Integer ' Stores FFT Data
Dim picHt As Long
Dim picWt As Long
Dim bolStart(0 To 255) As Boolean
Dim bolDelay(0 To 255) As Single

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event Click()
Public Event DblClick()

Private Function Sqrt(ByVal num As Double) As Double
    Sqrt = num ^ 0.5
End Function

Public Sub doStop()
    On Error Resume Next
    UserControl.Cls
End Sub
Public Sub Oscilliscope(SampleData() As Integer)
    Dim x, r, H As Long
    Dim tmpColor As Long
    
    x = 0
    UserControl.Cls
    For r = 0 To 254 Step 1
        H = ((SampleData(r) + 32768) / 65535 * picHt)
        tmpColor = GetPixel(picOsc.hDC, 1, picHt - H)
        x = x + UserControl.ScaleWidth / 254
        Select Case tOsc.DrawMode
            Case 0 'dot
                SetPixel UserControl.hDC, x, (picHt - H), tmpColor
            Case 1 'Line
                BitBlt UserControl.hDC, x, Round(H), 1, 1, picOsc.hDC, 0, picHt - (picHt - H), vbSrcCopy
            Case 2 'Solid
                BitBlt UserControl.hDC, x, H, 1, picHt / 2 - H, picOsc.hDC, 0, H, vbSrcCopy
        End Select
    Next
End Sub
Public Sub Spectrum(fft() As Single)
    Dim x As Long
    Dim Y As Long
    Dim r As Long
    Dim i As Integer
    Dim H As Long
    Dim tmpColor As Long
    ' Loop
    For i = 0 To 255
        Data2(i) = fft(i)
        If (Data2(i) + (Data2(i) * (i * 0.25))) * 40 > data(i) Then
            data(i) = (Data2(i) + (Data2(i) * (i * 0.25))) * 40
        Else
            data(i) = data(i) - 1
        End If
    Next
    
    r = 0
    x = 1
    UserControl.Cls
    
    
    Select Case tSpec.DrawMode
        '******************Draw Bar********************
        Case 0
            
            For r = 0 To 248
                Y = (data(r) + data(r + 1) + data(r + 2) + data(r + 3) + data(r + 4) + data(r + 5) + data(r + 6) + data(r + 7)) / 16
                If Y > picHt Then Y = picHt
                
                '-------------------Bar------------------
                Select Case tSpec.FillMode
                    Case 0 'Normal
                        StretchBlt UserControl.hDC, x, Round(picHt - Y), 3, Y, picBar.hDC, 0, (picHt - Y), 3, Y, vbSrcCopy
                    Case 1 'Fire
                        StretchBlt UserControl.hDC, x, Round(picHt - Y), 3, Y, picBar.hDC, 0, 0, 3, Y, vbSrcCopy
                    Case 2 'Line
                        tmpColor = IIf((picHt - Y) >= picHt / 3 * 2, tVisual.Color1, IIf(picHt - Y >= picHt / 3, tVisual.Color2, tVisual.Color3))
                        UserControl.Line (x, picHt)-(x + 2, picHt - Y), tmpColor, BF
                   End Select
                
                '----------------------Peak----------------------
                If bolDelay(r) = tSpec.PeakDelay Then
                    Data3(r) = Data3(r) - tSpec.PeakDrop
                    bolStart(r) = True
                End If
                bolDelay(r) = bolDelay(r) + 1
                If Data3(r) = Y Then
                    bolDelay(r) = 0
                    bolStart(r) = False
                End If
                If Data3(r) > Y And bolStart(r) Then
                    bolDelay(r) = tSpec.PeakDelay
                End If
                If Data3(r) < Y Then
                    Data3(r) = (Y)
                End If
                If tSpec.ShowPeak Then
                    BitBlt UserControl.hDC, x, Round(picHt - Data3(r)), 3, 1, picPeak.hDC, 0, 0, vbSrcCopy
                End If
                r = r + 15 'Int(256 / (UserControl.ScaleWidth / 4))
                x = x + 4
            Next r
       
       '****************Line***************
        Case 1
            
            For i = 0 To picWt
                H = (Sqrt(fft(i + 1) + fft(i + 2))) * picHt * 2
                If H > picHt Then H = picHt
                
                If bolDelay(i) = tSpec.PeakDelay Then
                    Data3(i) = Data3(i) - tSpec.PeakDrop
                    bolStart(i) = True
                End If
                bolDelay(i) = bolDelay(1) + 1
                If Data3(i) = H Then
                    bolDelay(i) = 0
                    bolStart(i) = False
                End If
                If Data3(i) > H And bolStart(i) Then
                    bolDelay(i) = tSpec.PeakDelay
                End If
                If Data3(i) < H Then
                    Data3(i) = H
                End If
                
                'draw line
                Select Case tSpec.FillMode
                    Case 0 'Normal
                        BitBlt UserControl.hDC, x, Round(picHt - H), 1, Round(H), picBar.hDC, 0, (picHt - H), vbSrcCopy
                    Case 1 'Fire
                        BitBlt UserControl.hDC, x, Round(picHt - H), 1, Round(H), picBar.hDC, 0, 0, vbSrcCopy
                    Case 2 'Line
                        tmpColor = GetPixel(picBar.hDC, 1, picHt - H)
                        'tmpColor = IIf((picHt - H) >= picHt / 3 * 2, tVisual.Color1, IIf(picHt - H >= picHt / 3, tVisual.Color2, tVisual.Color3))
                        UserControl.Line (x, picHt)-(x, picHt - H), tmpColor
                End Select
                
                If tSpec.ShowPeak Then
                    BitBlt UserControl.hDC, x, Round(picHt - Data3(i)), 1, 1, picPeak.hDC, 0, 0, vbSrcCopy
                End If
                
                x = x + 1
            Next i
    End Select
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
Private Sub UserControl_Resize()
    On Error Resume Next
    picHt = UserControl.ScaleHeight
    picWt = IIf(UserControl.ScaleWidth > 252, 252, UserControl.ScaleWidth)
    picBar.Width = UserControl.ScaleWidth
    picBar.Height = UserControl.ScaleHeight
    picPeak.Width = UserControl.ScaleWidth
    picOsc.Width = UserControl.ScaleWidth
    picOsc.Height = UserControl.ScaleHeight
End Sub
Public Sub setupVisual()
    On Error Resume Next
    picPeak.BackColor = tVisual.ColorPeak
    picBar.Cls
    If Not FileExists(tVisual.Pict) Then
        Call FillGradient(picBar.hDC, 0, 0, picBar.Width, picBar.Height / 2, tVisual.Color2, tVisual.Color3, Fill_Vertical)
        Call FillGradient(picBar.hDC, 0, picBar.Height / 2, picBar.Width, picBar.Height / 2, tVisual.Color1, tVisual.Color2, Fill_Vertical)
    Else
        Set picBar.Picture = LoadPicture(tVisual.Pict)
    End If
    'Setup Oscillscope Color
    picOsc.Cls
    Call FillGradient(picOsc.hDC, 0, 0, picOsc.Width, picOsc.Height \ 4, tVisual.ColorOsc2, tVisual.ColorOsc3, Fill_Vertical)
    Call FillGradient(picOsc.hDC, 0, picOsc.Height \ 4, picOsc.Width, picOsc.Height \ 4, tVisual.ColorOsc1, tVisual.ColorOsc2, Fill_Vertical)
    Call FillGradient(picOsc.hDC, 0, picOsc.Height \ 2 - 1, picOsc.Width, picOsc.Height \ 4 + 1, tVisual.ColorOsc2, tVisual.ColorOsc1, Fill_Vertical)
    Call FillGradient(picOsc.hDC, 0, picOsc.Height - picOsc.Height \ 4 - 2, picOsc.Width, picOsc.Height \ 4 + 3, tVisual.ColorOsc3, tVisual.ColorOsc2, Fill_Vertical)
End Sub
Public Sub Draw(pic As StdPicture, xS As Long, yS As Long, widthS As Long, heightS As Long)
    UserControl.Cls
    UserControl.PaintPicture pic, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, xS, yS, widthS, heightS, vbSrcCopy
    UserControl.Picture = UserControl.Image
End Sub
Public Property Let SpecLowColor(ByVal NewColor As OLE_COLOR)
    tVisual.Color1 = NewColor
    PropertyChanged "SpecLowColor"
End Property
Public Property Get SpecLowColor() As OLE_COLOR
    SpecLowColor = tVisual.Color1
End Property
Public Property Let SpecMidColor(ByVal NewColor As OLE_COLOR)
    tVisual.Color2 = NewColor
    PropertyChanged "SpecMidColor"
End Property
Public Property Get SpecMidColor() As OLE_COLOR
    SpecMidColor = tVisual.Color2
End Property
Public Property Let SpecHiColor(ByVal NewColor As OLE_COLOR)
    tVisual.Color3 = NewColor
    PropertyChanged "SpecHiColor"
End Property
Public Property Get SpecHiColor() As OLE_COLOR
    SpecHiColor = tVisual.Color3
End Property
Public Property Let SpecPeakColor(ByVal NewColor As OLE_COLOR)
    tVisual.ColorPeak = NewColor
    PropertyChanged "SpecPeakColor"
End Property
Public Property Get SpecPeakColor() As OLE_COLOR
    SpecPeakColor = tVisual.ColorPeak
End Property
Public Property Let SpecBarPic(ByVal NewPic As String)
    tVisual.Pict = NewPic
    PropertyChanged "SpecBarPic"
End Property
Public Property Get SpecBarPic() As String
    SpecBarPic = tVisual.Pict
End Property

Public Property Let SpecDrawMode(ByVal NewVal As Integer)
    tSpec.DrawMode = NewVal
    PropertyChanged "SpecDrawMode"
End Property
Public Property Get SpecDrawMode() As Integer
    SpecDrawMode = tSpec.DrawMode
End Property
Public Property Let SpecFillMode(ByVal NewVal As Integer)
    tSpec.FillMode = NewVal
    PropertyChanged "SpecFillMode"
End Property
Public Property Get SpecFillMode() As Integer
    SpecFillMode = tSpec.FillMode
End Property
Public Property Let SpecShowPeak(ByVal NewVal As Boolean)
    tSpec.ShowPeak = NewVal
    PropertyChanged "SpecShowPeak"
End Property
Public Property Get SpecShowPeak() As Boolean
    SpecShowPeak = tSpec.ShowPeak
End Property
Public Property Let SpecPeakDelay(ByVal NewVal As Long)
    tSpec.PeakDelay = NewVal
    PropertyChanged "SpecPeakDelay"
End Property
Public Property Get SpecPeakDelay() As Long
    SpecPeakDelay = tSpec.PeakDelay
End Property
Public Property Let SpecPeakDrop(ByVal NewVal As Long)
    tSpec.PeakDrop = NewVal
    PropertyChanged "SpecPeakDrop"
End Property
Public Property Get SpecPeakDrop() As Long
    SpecPeakDrop = tSpec.PeakDrop
End Property
Public Property Let OscDrawMode(ByVal NewVal As Integer)
    tOsc.DrawMode = NewVal
    PropertyChanged "OscDrawMode"
End Property
Public Property Get OscDrawMode() As Integer
    OscDrawMode = tOsc.DrawMode
End Property

Public Property Let OscLowColor(ByVal NewColor As OLE_COLOR)
    tVisual.ColorOsc1 = NewColor
    PropertyChanged "OscLowColor"
End Property
Public Property Get OscLowColor() As OLE_COLOR
    OscLowColor = tVisual.ColorOsc1
End Property
Public Property Let OscMidColor(ByVal NewColor As OLE_COLOR)
    tVisual.ColorOsc2 = NewColor
    PropertyChanged "OscMidColor"
End Property
Public Property Get OscMidColor() As OLE_COLOR
    OscMidColor = tVisual.ColorOsc2
End Property
Public Property Let OscHiColor(ByVal NewColor As OLE_COLOR)
    tVisual.ColorOsc3 = NewColor
    PropertyChanged "OscHiColor"
End Property
Public Property Get OscHiColor() As OLE_COLOR
    OscHiColor = tVisual.ColorOsc3
End Property
Public Property Let StyleVis(NewVal As Integer)
    intStyle = NewVal
    PropertyChanged "StyleVis"
End Property
Public Property Get StyleVis() As Integer
    StyleVis = intStyle
End Property

