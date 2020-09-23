VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P_Spectrum Plugin"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   33
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   240
      Width           =   6015
      Begin VB.HScrollBar hscZoom 
         Height          =   315
         Left            =   600
         Max             =   50
         Min             =   5
         TabIndex        =   34
         Top             =   2520
         Value           =   5
         Width           =   5295
      End
      Begin MSComDlg.CommonDialog cdloColor 
         Left            =   120
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboSpec 
         Height          =   315
         Index           =   0
         ItemData        =   "frmConfig.frx":0000
         Left            =   960
         List            =   "frmConfig.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox chkSpectrum 
         Appearance      =   0  'Flat
         Caption         =   "Show peak"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1000
         Width           =   1120
      End
      Begin VB.CheckBox chkSpectrum 
         Appearance      =   0  'Flat
         Caption         =   "Peak gradient"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1370
         Width           =   1335
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   16
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Fire"
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   15
         Top             =   150
         Width           =   615
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Column"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   14
         Top             =   150
         Width           =   855
      End
      Begin VB.TextBox txtSpectrum 
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   13
         Top             =   970
         Width           =   735
      End
      Begin VB.TextBox txtSpectrum 
         Height          =   315
         Index           =   2
         Left            =   3840
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtSpectrum 
         Height          =   315
         Index           =   3
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   11
         Top             =   970
         Width           =   735
      End
      Begin VB.TextBox txtSpectrum 
         Height          =   315
         Index           =   4
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1340
         Width           =   735
      End
      Begin VB.TextBox txtSpectrum 
         Height          =   315
         Index           =   0
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkSpectrum 
         Appearance      =   0  'Flat
         Caption         =   "Show bar"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   1095
      End
      Begin VB.ComboBox cboSpec 
         Height          =   315
         Index           =   1
         ItemData        =   "frmConfig.frx":0022
         Left            =   3840
         List            =   "frmConfig.frx":002F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboSpec 
         Height          =   315
         Index           =   2
         ItemData        =   "frmConfig.frx":0044
         Left            =   2040
         List            =   "frmConfig.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cboSpec 
         Height          =   315
         Index           =   3
         ItemData        =   "frmConfig.frx":006A
         Left            =   2040
         List            =   "frmConfig.frx":0077
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblZoom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5820
         TabIndex        =   36
         Top             =   2160
         Width           =   60
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   2580
         Width           =   405
      End
      Begin VB.Shape shpVisual 
         BorderColor     =   &H80000010&
         Height          =   3015
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Style"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   765
      End
      Begin VB.Label lblSpectrum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LowColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblSpectrum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MidColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblSpectrum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HightColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblSpectrum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PeakColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PeakPause"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   3
         Left            =   2985
         TabIndex        =   27
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bars"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   5
         Left            =   1560
         TabIndex        =   26
         Top             =   660
         Width           =   315
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scale"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   1560
         TabIndex        =   25
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PeakHeight"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   2985
         TabIndex        =   24
         Top             =   1770
         Width           =   840
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PeakDrop"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   9
         Left            =   2985
         TabIndex        =   23
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PeakStyle"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   13
         Left            =   2985
         TabIndex        =   22
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PosX"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   14
         Left            =   1560
         TabIndex        =   21
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label lblCaptionOpt6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PosY"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   15
         Left            =   1560
         TabIndex        =   20
         Top             =   1740
         Width           =   375
      End
   End
   Begin VB.PictureBox picPeak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1920
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   3
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   -360
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picMiddle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   -120
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   1
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   360
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                        (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                                        (ByVal HWND As Long, ByVal nIndex As Long) As Long
Const GWL_STYLE = (-16)
Const ES_NUMBER = &H2000&
Const SW_SHOWNORMAL = 1

Private Sub cboSpec_Click(Index As Integer)
    Select Case Index
        Case 0
            tSpec.DrawStyle = cboSpec(0).ListIndex
        Case 1
            tSpec.PeakDraw = cboSpec(1).ListIndex
        Case 2
            tSpec.PosX = cboSpec(2).ListIndex
        Case 3
            tSpec.PosY = cboSpec(3).ListIndex
    End Select
End Sub

Private Sub chkSpectrum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            tSpec.ShowBar = CBool(chkSpectrum(0).Value)
        Case 1
            tSpec.ShowPeak = CBool(chkSpectrum(1).Value)
        Case 2
            tSpec.PeakGradient = CBool(chkSpectrum(2).Value)
    End Select
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    InitCommonControls
    
    Call SetNumber(txtSpectrum(0), True)
    Call SetNumber(txtSpectrum(1), True)
    Call SetNumber(txtSpectrum(2), True)
    Call SetNumber(txtSpectrum(3), True)
    Call SetNumber(txtSpectrum(4), True)

    
    hscZoom.Value = tSpec.Zoom * 10
    cboSpec(0).ListIndex = tSpec.DrawStyle
    cboSpec(1).ListIndex = tSpec.PeakDraw
    cboSpec(2).ListIndex = tSpec.PosX
    cboSpec(3).ListIndex = tSpec.PosY
    optFillMode(tSpec.FillMode).Value = True
    If tSpec.ShowBar Then chkSpectrum(0).Value = Checked
    If tSpec.ShowPeak Then chkSpectrum(1).Value = Checked
    If tSpec.PeakGradient Then chkSpectrum(2).Value = Checked
    txtSpectrum(0).Text = tSpec.BarNumber
    txtSpectrum(1).Text = tSpec.ScaleW
    txtSpectrum(2).Text = tSpec.PeakHeight
    txtSpectrum(3).Text = tSpec.PeakInteval
    txtSpectrum(4).Text = tSpec.PeakDrop
        
    lblSpectrum(0).BackColor = tSpec.LowColor
    lblSpectrum(1).BackColor = tSpec.MidColor
    lblSpectrum(2).BackColor = tSpec.HightColor
    lblSpectrum(3).BackColor = tSpec.PeakColor
    picPeak.BackColor = tSpec.PeakColor
End Sub

Private Sub Form_Terminate()
    Dim strFileconfig As String
    strFileconfig = App.Path & "\M3P_vis.ini"
    
    WriteINI "visSpectrum", "Zoom", tSpec.Zoom, strFileconfig
    WriteINI "visSpectrum", "DrawStyle", tSpec.DrawStyle, strFileconfig
    WriteINI "visSpectrum", "FillMode", tSpec.FillMode, strFileconfig
    WriteINI "visSpectrum", "PosX", tSpec.PosX, strFileconfig
    WriteINI "visSpectrum", "PosY", tSpec.PosY, strFileconfig
    WriteINI "visSpectrum", "LowColor", tSpec.LowColor, strFileconfig
    WriteINI "visSpectrum", "MidColor", tSpec.MidColor, strFileconfig
    WriteINI "visSpectrum", "HightColor", tSpec.HightColor, strFileconfig
    WriteINI "visSpectrum", "ShowBar", tSpec.ShowBar, strFileconfig
    WriteINI "visSpectrum", "BarNumber", tSpec.BarNumber, strFileconfig
    WriteINI "visSpectrum", "ScaleW", tSpec.ScaleW, strFileconfig
    WriteINI "visSpectrum", "ShowPeak", tSpec.ShowPeak, strFileconfig
    WriteINI "visSpectrum", "PeakHeight", tSpec.PeakHeight, strFileconfig
    WriteINI "visSpectrum", "PeakGradient", tSpec.PeakGradient, strFileconfig
    WriteINI "visSpectrum", "PeakColor", tSpec.PeakColor, strFileconfig
    WriteINI "visSpectrum", "PeakInteval", tSpec.PeakInteval, strFileconfig
    WriteINI "visSpectrum", "PeakDrop", tSpec.PeakDrop, strFileconfig
    WriteINI "visSpectrum", "PeakDraw", tSpec.PeakDraw, strFileconfig
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strFileconfig As String
    strFileconfig = App.Path & "\M3P_vis.ini"
    
    WriteINI "visSpectrum", "Zoom", tSpec.Zoom, strFileconfig
    WriteINI "visSpectrum", "DrawStyle", tSpec.DrawStyle, strFileconfig
    WriteINI "visSpectrum", "FillMode", tSpec.FillMode, strFileconfig
    WriteINI "visSpectrum", "PosX", tSpec.PosX, strFileconfig
    WriteINI "visSpectrum", "PosY", tSpec.PosY, strFileconfig
    WriteINI "visSpectrum", "LowColor", tSpec.LowColor, strFileconfig
    WriteINI "visSpectrum", "MidColor", tSpec.MidColor, strFileconfig
    WriteINI "visSpectrum", "HightColor", tSpec.HightColor, strFileconfig
    WriteINI "visSpectrum", "ShowBar", tSpec.ShowBar, strFileconfig
    WriteINI "visSpectrum", "BarNumber", tSpec.BarNumber, strFileconfig
    WriteINI "visSpectrum", "ScaleW", tSpec.ScaleW, strFileconfig
    WriteINI "visSpectrum", "ShowPeak", tSpec.ShowPeak, strFileconfig
    WriteINI "visSpectrum", "PeakHeight", tSpec.PeakHeight, strFileconfig
    WriteINI "visSpectrum", "PeakGradient", tSpec.PeakGradient, strFileconfig
    WriteINI "visSpectrum", "PeakColor", tSpec.PeakColor, strFileconfig
    WriteINI "visSpectrum", "PeakInteval", tSpec.PeakInteval, strFileconfig
    WriteINI "visSpectrum", "PeakDrop", tSpec.PeakDrop, strFileconfig
    WriteINI "visSpectrum", "PeakDraw", tSpec.PeakDraw, strFileconfig
End Sub

Private Sub hscZoom_Change()
    lblZoom.Caption = tSpec.Zoom
End Sub

Private Sub hscZoom_Scroll()
    tSpec.Zoom = hscZoom.Value / 10
    lblZoom.Caption = tSpec.Zoom
End Sub

Private Sub lblSpectrum_Click(Index As Integer)
    On Error Resume Next
    Dim lngColor As Long
    With cdloColor
        .CancelError = True
        .DialogTitle = "MP3_proPlayer : Choice Color"
        .Flags = &H1&
        .ShowColor
        lngColor = .Color
    End With
    lblSpectrum(Index).BackColor = lngColor
    Select Case Index
        Case 0, 1, 2
            tSpec.LowColor = lblSpectrum(0).BackColor
            tSpec.MidColor = lblSpectrum(1).BackColor
            tSpec.HightColor = lblSpectrum(2).BackColor
            picTop.Cls
            Call FillGradient(picTop.hDc, 0, 0, picTop.Width, picTop.Height / 2, tSpec.MidColor, tSpec.LowColor, Fill_Vertical)
            Call FillGradient(picTop.hDc, 0, picTop.Height / 2, picTop.Width, picTop.Height / 2, tSpec.HightColor, tSpec.MidColor, Fill_Vertical)
            picBottom.Cls
            Call FillGradient(picBottom.hDc, 0, 0, picBottom.Width, picBottom.Height / 2, tSpec.MidColor, tSpec.HightColor, Fill_Vertical)
            Call FillGradient(picBottom.hDc, 0, picBottom.Height / 2, picBottom.Width, picBottom.Height / 2, tSpec.LowColor, tSpec.MidColor, Fill_Vertical)
            picMiddle.Cls
            Call FillGradient(picMiddle.hDc, 0, 0, picMiddle.Width, picMiddle.Height / 4, tSpec.MidColor, tSpec.HightColor, Fill_Vertical)
            Call FillGradient(picMiddle.hDc, 0, picMiddle.Height / 4, picMiddle.Width, picMiddle.Height / 4, tSpec.LowColor, tSpec.MidColor, Fill_Vertical)
            Call FillGradient(picMiddle.hDc, 0, picMiddle.Height / 2, picMiddle.Width, picMiddle.Height / 4, tSpec.MidColor, tSpec.LowColor, Fill_Vertical)
            Call FillGradient(picMiddle.hDc, 0, picMiddle.Height - picMiddle.Height / 4, picMiddle.Width, picMiddle.Height / 4, tSpec.HightColor, tSpec.MidColor, Fill_Vertical)
        Case 3
            tSpec.PeakColor = lblSpectrum(3).BackColor
            picPeak.BackColor = tSpec.PeakColor
    End Select
End Sub

Private Sub optFillMode_Click(Index As Integer)
    tSpec.FillMode = Index
End Sub

Private Sub picTop_Resize()
    picBottom.Height = picTop.Height
    picBottom.Width = picTop.Width
    picMiddle.Height = picTop.Height
    picMiddle.Width = picTop.Width
    picPeak.Width = picTop.Width
    picPeak.Height = picTop.Height

    picBottom.Cls
    Call FillGradient(picBottom.hDc, 0, 0, picBottom.Width, picBottom.Height / 2, tSpec.MidColor, tSpec.HightColor, Fill_Vertical)
    Call FillGradient(picBottom.hDc, 0, picBottom.Height / 2, picBottom.Width, picBottom.Height / 2, tSpec.LowColor, tSpec.MidColor, Fill_Vertical)
    
    picTop.Cls
    Call FillGradient(picTop.hDc, 0, 0, picTop.Width, picTop.Height / 2, tSpec.MidColor, tSpec.LowColor, Fill_Vertical)
    Call FillGradient(picTop.hDc, 0, picTop.Height / 2, picTop.Width, picTop.Height / 2, tSpec.HightColor, tSpec.MidColor, Fill_Vertical)
    
    picMiddle.Cls
    Call FillGradient(picMiddle.hDc, 0, 0, picMiddle.Width, picMiddle.Height / 4, tSpec.MidColor, tSpec.HightColor, Fill_Vertical)
    Call FillGradient(picMiddle.hDc, 0, picMiddle.Height / 4, picMiddle.Width, picMiddle.Height / 4, tSpec.LowColor, tSpec.MidColor, Fill_Vertical)
    Call FillGradient(picMiddle.hDc, 0, picMiddle.Height / 2, picMiddle.Width, picMiddle.Height / 4, tSpec.MidColor, tSpec.LowColor, Fill_Vertical)
    Call FillGradient(picMiddle.hDc, 0, picMiddle.Height - picMiddle.Height / 4, picMiddle.Width, picMiddle.Height / 4, tSpec.HightColor, tSpec.MidColor, Fill_Vertical)
End Sub

Private Sub txtSpectrum_Change(Index As Integer)
    On Error GoTo beep
    Select Case Index
        Case 0
            If CInt(txtSpectrum(0).Text) > 0 And CInt(txtSpectrum(0).Text) < 255 Then
                If CInt(txtSpectrum(0).Text) > 127 And cboSpec(2).ListIndex = 1 Then txtSpectrum(0).Text = 32
                tSpec.BarNumber = CInt(txtSpectrum(0).Text)
            Else
                tSpec.BarNumber = 32
            End If
            If tSpec.BarNumber > 0 Then
                tSpec.BarWidth = (picTop.Width - tSpec.BarNumber * tSpec.ScaleW) \ tSpec.BarNumber
            End If
        Case 1
            tSpec.ScaleW = CInt(txtSpectrum(1).Text)
        Case 2
            tSpec.PeakHeight = CInt(txtSpectrum(2).Text)
        Case 3
            If txtSpectrum(3).Text = "" Or txtSpectrum(3).Text < 0 Then txtSpectrum(3).Text = 6
            tSpec.PeakInteval = CInt(txtSpectrum(3).Text)
        Case 4
            tSpec.PeakDrop = CInt(txtSpectrum(4).Text)
    End Select
beep:
    If Err.Number <> 0 Then
        txtSpectrum(0).Text = 32
        txtSpectrum(1).Text = 1
        txtSpectrum(2).Text = 1
        txtSpectrum(3).Text = 50
        txtSpectrum(4).Text = 6
    End If
End Sub
Sub SetNumber(NumberText As TextBox, flag As Boolean)
    Dim curstyle As Long
    Dim newstyle As Long
    
    curstyle = GetWindowLong(NumberText.HWND, GWL_STYLE)
    If flag Then
       curstyle = curstyle Or ES_NUMBER
    Else
       curstyle = curstyle And (Not ES_NUMBER)
    End If
    newstyle = SetWindowLong(NumberText.HWND, GWL_STYLE, curstyle)
    NumberText.Refresh
End Sub

