VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0179B2D7-CD62-439D-BE78-CF820F5A4B44}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmMedia 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   1200
   Icon            =   "frmMedia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ShowInTaskbar   =   0   'False
   Begin M3P_Control.Player PlayerBack 
      Height          =   375
      Left            =   6720
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   5400
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin M3P_Control.Player Player 
      Height          =   375
      Left            =   7440
      TabIndex        =   26
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picMini 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   0
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   294
      TabIndex        =   2
      ToolTipText     =   "MP3_ProPlayer : main window"
      Top             =   2880
      Width           =   4410
      Begin VB.PictureBox picMiniMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   0
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   3
         ToolTipText     =   "MP3_ProPlayer : main window"
         Top             =   0
         Width           =   4215
         Begin M3P_Control.ctlSlider sldMiniPosition 
            Height          =   135
            Left            =   240
            TabIndex        =   50
            Top             =   1080
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   238
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   2
            CueHeight       =   9
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":000C
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.ScrollLabel srlMiniInfor 
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
         End
         Begin M3P_Control.Progress prgVU 
            Height          =   150
            Index           =   0
            Left            =   840
            TabIndex        =   9
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   265
            Max             =   32768
            Value           =   32768
            Enabled         =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   255
            BorderStyle     =   0
            MouseIcon       =   "frmMedia.frx":0028
         End
         Begin M3P_Control.Progress prgVU 
            Height          =   150
            Index           =   1
            Left            =   2280
            TabIndex        =   10
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   265
            Max             =   32768
            Value           =   0
            Enabled         =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   255
            BorderStyle     =   0
            MouseIcon       =   "frmMedia.frx":0044
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   13
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   14
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   15
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   16
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   17
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   18
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   19
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   8
            Left            =   3840
            TabIndex        =   20
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   9
            Left            =   3240
            TabIndex        =   21
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   10
            Left            =   3600
            TabIndex        =   22
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   11
            Left            =   3960
            TabIndex        =   23
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblMiniDur 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   7
            Top             =   600
            Width           =   405
         End
         Begin VB.Label lblMiniPos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   405
         End
         Begin VB.Label lblMiniMpgInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   120
            Width           =   540
         End
         Begin VB.Label lblMiniMpgInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00.0"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   4
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.PictureBox picMiniEqua 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   0
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   8
         Top             =   1920
         Width           =   4215
         Begin M3P_Control.ctlSlider sldMiniBal 
            Height          =   135
            Left            =   2400
            TabIndex        =   64
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            Min             =   -100
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   9
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0060
         End
         Begin M3P_Control.ctlSlider sldMiniVol 
            Height          =   135
            Left            =   480
            TabIndex        =   63
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   9
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":007C
         End
         Begin M3P_Control.ctlSlider sldMiniAmp 
            Height          =   975
            Left            =   480
            TabIndex        =   52
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0098
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMiniMedia 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   0
            Left            =   840
            TabIndex        =   53
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":00B4
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   1
            Left            =   1080
            TabIndex        =   54
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":00D0
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   2
            Left            =   1320
            TabIndex        =   55
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":00EC
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   3
            Left            =   1560
            TabIndex        =   56
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0108
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   4
            Left            =   1800
            TabIndex        =   57
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0124
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   5
            Left            =   2040
            TabIndex        =   58
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0140
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   6
            Left            =   2280
            TabIndex        =   59
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":015C
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   7
            Left            =   2520
            TabIndex        =   60
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0178
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   8
            Left            =   2760
            TabIndex        =   61
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0194
         End
         Begin M3P_Control.ctlSlider sldMiniEqua 
            Height          =   975
            Index           =   9
            Left            =   3000
            TabIndex        =   62
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":01B0
         End
      End
   End
   Begin VB.Timer tmrTaskbar 
      Interval        =   200
      Left            =   5280
      Top             =   4080
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   6000
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Timer tmrPos 
      Interval        =   50
      Left            =   5280
      Top             =   3600
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   0
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   582
      TabIndex        =   0
      ToolTipText     =   "MP3_ProPlayer : main window"
      Top             =   0
      Width           =   8730
      Begin VB.PictureBox picMainMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   28
         Top             =   120
         Width           =   4335
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   11
            Left            =   3120
            TabIndex        =   29
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   10
            Left            =   2760
            TabIndex        =   30
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   31
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   8
            Left            =   3240
            TabIndex        =   32
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   7
            Left            =   2880
            TabIndex        =   33
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   6
            Left            =   2520
            TabIndex        =   34
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   35
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   36
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   37
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   38
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   39
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.Visualization Vis 
            Height          =   255
            Left            =   960
            TabIndex        =   40
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
         End
         Begin M3P_Control.ScrollLabel srlInfor 
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   49
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.ctlSlider sldPosition 
            Height          =   135
            Left            =   120
            TabIndex        =   51
            Top             =   1560
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   238
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   2
            CueHeight       =   9
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":01CC
         End
         Begin VB.Label lblPosition 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblDuration 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   43
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblMpgInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   42
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblMpgInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00.0"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   41
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.PictureBox picEqua 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   4560
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   241
         TabIndex        =   45
         Top             =   360
         Width           =   3615
         Begin M3P_Control.ctlSlider sldAmp 
            Height          =   975
            Left            =   240
            TabIndex        =   65
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":01E8
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   46
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.DynamicButton btnMedia 
            Height          =   255
            Index           =   12
            Left            =   3240
            TabIndex        =   47
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   0
            Left            =   840
            TabIndex        =   66
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0204
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   1
            Left            =   1080
            TabIndex        =   67
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0220
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   2
            Left            =   1320
            TabIndex        =   68
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":023C
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   3
            Left            =   1560
            TabIndex        =   69
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0258
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   4
            Left            =   1800
            TabIndex        =   70
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0274
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   5
            Left            =   2040
            TabIndex        =   71
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0290
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   6
            Left            =   2280
            TabIndex        =   72
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":02AC
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   7
            Left            =   2520
            TabIndex        =   73
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":02C8
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   8
            Left            =   2760
            TabIndex        =   74
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":02E4
         End
         Begin M3P_Control.ctlSlider sldEqua 
            Height          =   975
            Index           =   9
            Left            =   3000
            TabIndex        =   75
            Top             =   120
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1720
            Max             =   -12
            Min             =   12
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   65
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0300
         End
         Begin M3P_Control.ctlSlider sldVolume 
            Height          =   135
            Left            =   120
            TabIndex        =   76
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   9
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":031C
         End
         Begin M3P_Control.ctlSlider sldBalance 
            Height          =   135
            Left            =   1920
            TabIndex        =   77
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            Min             =   -100
            BackColor       =   -2147483643
            BorderStyle     =   0
            SliderOrien     =   0
            CueWidth        =   1
            CueHeight       =   9
            ForeColor       =   -2147483647
            MouseIcon       =   "frmMedia.frx":0338
         End
      End
   End
   Begin VB.Timer tmrTestPic 
      Interval        =   100
      Left            =   5280
      Top             =   3120
   End
   Begin VB.Timer tmrVisual 
      Left            =   4800
      Top             =   3600
   End
   Begin VB.Timer tmrPlay 
      Interval        =   900
      Left            =   4800
      Top             =   3120
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Long
Private posTaskbar As Integer
Private intEQdownX As Single
Private intEQdownY As Single
Private SampleData(1000) As Integer
Private channel As Long
Private fft(512) As Single

Private Sub VUMeter()
    Dim lngLevel As Long
    Dim lngLeftLevel As Long
    Dim lngRightLevel As Long
    Dim H As Long, i As Integer, x As Long
    lngLevel = BASS_ChannelGetLevel(Player.handle)
    If lngLevel = -1 Then Exit Sub
    lngLeftLevel = LoWord(lngLevel)
    lngRightLevel = HiWord(lngLevel)
    prgVU(0).value = 32768 - lngLeftLevel
    prgVU(1).value = lngRightLevel
End Sub

Private Sub btnMedia_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0 'prev
            Call BackTrack
        Case 1 'play
            Call btnPlayClick
        Case 2 'pause
            Call btnPauseClick
        Case 3 'next
            Call NextTrack
        Case 4 'Stop
            Call StopPlayer
        Case 5 'Shuffe
            Call btnShuffeClick
        Case 6 'Repeat
            Call btnRepeatClick
        Case 7 'Show EQ
            Call btnShowEQClick
        Case 8 'Show Playlist
            Call btnShowPLClick
        Case 9
            Call ChangeMask(True)
            Call ShowEQ
        Case 10
            If tAppConfig.bolTaskbar Then
                frmMenu.WindowState = vbMinimized
            Else
                frmMedia.Visible = False
                If tPlaylistConfig.bolHidePL = False Then frmPlayList.Visible = False
            End If
        Case 11
            If Player.PlayState = 1 Then Call StopPlayer
            Unload Me
        Case 12
            bolEQEnabled = Not bolEQEnabled
            btnMedia(12).bolOn = bolEQEnabled
            If bolEQEnabled Then
                Call SetEqual(Player)
                For i = 0 To 9
                    Call UpdateEqual(i, sldEqua(i).value)
                Next i
                'Call BASS_ChannelSetDSP(Player.handle, AddressOf PreAmp, 0, 2)
            Else
                For i = 0 To 9
                    Call BASS_ChannelRemoveFX(Player.handle, EQ(i))
                Next i
                'Call BASS_ChannelRemoveDSP(Player.handle, AddressOf PreAmp)
            End If
            WriteINI "Equalizer", "EqualEnabled", bolEQEnabled, strFileconfig
        Case 13
            PopupMenu frmMenu.mnuEqua, vbPopupMenuLeftAlign
    End Select

End Sub

Private Sub btnMedia_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Select Case Index
            Case 0 'prev
                srlInfor.Title = LangCap(118)
            Case 1 'play
                If Player.PlayState = 1 Then
                    srlInfor.Title = LangCap(119)
                Else
                    If Player.PlayState = 2 Then
                        srlInfor.Title = LangCap(120)
                    Else
                        srlInfor.Title = LangCap(121)
                    End If
                End If
            Case 2 'pause
                If Player.PlayState = 1 Then
                    srlInfor.Title = LangCap(122)
                Else
                    srlInfor.Title = LangCap(123)
                End If
            Case 3 'next
                srlInfor.Title = LangCap(124)
            Case 4 'stop
                srlInfor.Title = LangCap(125)
            Case 5 'Shuffe
                If tPlayerConfig.bolShuffe Then
                    srlInfor.Title = LangCap(126)
                Else
                    srlInfor.Title = LangCap(127)
                End If
            Case 6 'Repeat
                If tPlayerConfig.bolRepeat Then
                    srlInfor.Title = LangCap(128)
                Else
                    srlInfor.Title = LangCap(129)
                End If
            Case 7 'ShowEQ
                If tPlayerConfig.bolShowEQ Then
                    srlInfor.Title = LangCap(130)
                Else
                    srlInfor.Title = LangCap(131)
                End If
            Case 8 'ShowPL
                If tPlaylistConfig.bolHidePL Then
                    srlInfor.Title = LangCap(132)
                Else
                    srlInfor.Title = LangCap(133)
                End If
            Case 9
                srlInfor.Title = LangCap(134)
            Case 10 'Minimize
                srlInfor.Title = LangCap(136)
            Case 11 'Close
                srlInfor.Title = LangCap(137)
            Case 12 'Equalizer Enabled
                If bolEQEnabled Then
                    srlInfor.Title = LangCap(138)
                Else
                    srlInfor.Title = LangCap(139)
                End If
            Case 13 'Show Equalizer Preset
                srlInfor.Title = LangCap(140)
        End Select
    End If
End Sub

Private Sub btnMedia_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
End Sub

Private Sub btnMiniMedia_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0 'prev
            Call BackTrack
        Case 1 'play
            Call btnPlayClick
        Case 2 'pause
            Call btnPauseClick
        Case 3
            Call NextTrack
        Case 4
            Call StopPlayer
        Case 5 'Shuffe
            Call btnShuffeClick
        Case 6 'Repeat
            Call btnRepeatClick
        Case 7 'Show EQ
            Call btnShowEQClick
        Case 8 'Show Playlist
            Call btnShowPLClick
        Case 9
            Call ChangeMask(False)
            Call ShowEQ
        Case 10
            If tAppConfig.bolTaskbar Then
                frmMenu.WindowState = vbMinimized
            Else
                frmMedia.Visible = False
                If tPlaylistConfig.bolHidePL = False Then frmPlayList.Visible = False
            End If
        Case 11
            If Player.PlayState = 1 Then Call StopPlayer
            Unload Me
        Case 12
            bolEQEnabled = Not bolEQEnabled
            btnMiniMedia(12).bolOn = bolEQEnabled
            If bolEQEnabled Then
                Call SetEqual(Player)
                For i = 0 To 9
                    Call UpdateEqual(i, sldEqua(i).value)
                Next i
                'Call BASS_ChannelSetDSP(Player.handle, AddressOf PreAmp, 0, 2)
            Else
                For i = 0 To 9
                    Call BASS_ChannelRemoveFX(Player.handle, EQ(i))
                Next i
                'Call BASS_ChannelRemoveDSP(Player.handle, AddressOf PreAmp)
            End If
            WriteINI "Equalizer", "EqualEnabled", bolEQEnabled, strFileconfig
        Case 13
            PopupMenu frmMenu.mnuEqua, vbPopupMenuLeftAlign
    End Select

End Sub

Private Sub btnMiniMedia_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Select Case Index
            Case 0 'prev
                srlMiniInfor.Title = LangCap(118)
            Case 1 'play
                If Player.PlayState = 1 Then
                    srlMiniInfor.Title = LangCap(119)
                Else
                    If Player.PlayState = 2 Then
                        srlMiniInfor.Title = LangCap(120)
                    Else
                        srlMiniInfor.Title = LangCap(121)
                    End If
                End If
            Case 2 'pause
                If Player.PlayState = 1 Then
                    srlMiniInfor.Title = LangCap(122)
                Else
                    srlMiniInfor.Title = LangCap(123)
                End If
            Case 3 'next
                srlMiniInfor.Title = LangCap(124)
            Case 4 'stop
                srlMiniInfor.Title = LangCap(125)
            Case 5 'Shuffe
                If tPlayerConfig.bolShuffe Then
                    srlMiniInfor.Title = LangCap(126)
                Else
                    srlMiniInfor.Title = LangCap(127)
                End If
            Case 6 'Repeat
                If tPlayerConfig.bolRepeat Then
                    srlMiniInfor.Title = LangCap(128)
                Else
                    srlMiniInfor.Title = LangCap(129)
                End If
            Case 7 'ShowEQ
                If tPlayerConfig.bolShowEQ Then
                    srlMiniInfor.Title = LangCap(130)
                Else
                    srlMiniInfor.Title = LangCap(131)
                End If
            Case 8 'ShowPL
                If tPlaylistConfig.bolHidePL Then
                    srlMiniInfor.Title = LangCap(132)
                Else
                    srlMiniInfor.Title = LangCap(133)
                End If
            Case 9
                srlMiniInfor.Title = LangCap(135)
            Case 10 'Minimize
                srlMiniInfor.Title = LangCap(136)
            Case 11 'Close
                srlMiniInfor.Title = LangCap(137)
            Case 12 'Equalizer Enabled
                If bolEQEnabled Then
                    srlMiniInfor.Title = LangCap(138)
                Else
                    srlMiniInfor.Title = LangCap(139)
                End If
            Case 13 'Show Equalizer Preset
                srlMiniInfor.Title = LangCap(140)
        End Select
    End If
End Sub
Private Sub btnMiniMedia_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
End Sub
Private Sub Form_Activate()
    bolLoading = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
    If KeyCode = 107 Then
        If sldVolume.value < 100 Then
            Call frmMedia.volume(frmMedia.sldVolume.value + 1)
        End If
    End If
    If KeyCode = 109 Then
        If sldVolume.value >= 1 Then
            Call frmMedia.volume(frmMedia.sldVolume.value - 1)
        End If
    End If
    If KeyCode = 36 Then Call GoTop 'Home
    If KeyCode = 66 Then Call BackTrack    'B
    If KeyCode = 37 Then Call Prev    'Left <--
    If KeyCode = 39 Then Call Forw    'Right -->
    If KeyCode = 78 Then Call NextTrack  'N
    If KeyCode = 35 Then Call GoBottom 'End
    If KeyCode = 76 Then Call btnMedia_Click(1) 'L
    If KeyCode = 80 Then Call btnMedia_Click(2) 'P
    If KeyCode = 32 Then Call StopPlayer   'Spacebar
    If KeyCode = 83 Then Call btnMedia_Click(5) 'S
    If KeyCode = 82 Then Call btnMedia_Click(6)  'R
    If KeyCode = 77 Then frmMedia.WindowState = vbMinimized ' M
    If KeyCode = 112 Then
        ShellExecute Me.hwnd, vbNullString, App.path & "\ReadMe.rtf", vbNullString, vbNullString, SW_SHOWNORMAL
    End If
    If KeyCode = 88 Then Unload Me 'X
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 69 Then
        Call btnMedia_Click(7) 'E
        btnMedia(7).bolOn = Not btnMedia(7).bolOn
        btnMedia(7).Refresh
    End If
    If KeyCode = 65 Then 'A
        If frmMenu.mnuOnTop.Checked Then
            Call AlwaysOnTop(Me, False)
            frmMenu.mnuOnTop.Checked = False
        Else
            Call AlwaysOnTop(Me, True)
            frmMenu.mnuOnTop.Checked = True
        End If
    End If
    If KeyCode = 87 Then 'W
        Call btnMedia_Click(8)
        btnMedia(8).bolOn = Not btnMedia(8).bolOn
        btnMedia(8).Refresh
    End If
    If KeyCode = 68 Then 'D
        If frmMenu.mnuDSP.Checked Then
            Unload frmDSP
        Else
            frmDSP.Show
            frmMenu.mnuDSP.Checked = True
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Local Error Resume Next
    Err.Clear
    Static bolBusy As Boolean
        If bolBusy = False Then
            bolBusy = True
            Select Case CLng(x)
                Case WM_LBUTTONDBLCLK
                    If tAppConfig.bolTaskbar Then
                        If frmMenu.WindowState = vbNormal Then
                            frmMenu.WindowState = vbMinimized
                        Else
                            frmMenu.WindowState = vbNormal
                        End If
                    Else
                        frmMedia.Visible = Not frmMedia.Visible
                        If frmMedia.Visible Then
                            If tPlaylistConfig.bolHidePL = False Then frmPlayList.Visible = True
                        Else
                            If tPlaylistConfig.bolHidePL = False Then frmPlayList.Visible = False
                        End If
                    End If
                    DoEvents
                Case WM_LBUTTONDOWN
                Case WM_LBUTTONUP
                Case WM_RBUTTONDBLCLK
                Case WM_RBUTTONDOWN
                Case WM_RBUTTONUP
                    'Right mouse button released: display popup menu
                    PopupMenu frmMenu.mnuMain, vbPopupMenuRightAlign
            End Select
            bolBusy = False
        End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo handle
    
    Dim DirList As New Collection
    Dim temp As String
    
    If tCurrentSkin.Infor.Name <> "Default" Then
        DirList.Add tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\"
        Do While DirList.Count
            temp = Dir$(DirList(1), vbDirectory)
            Do Until temp = ""
                If temp = "." Or temp = ".." Then
                ElseIf (GetAttr(RepairPath(DirList(1), temp)) And vbDirectory) = vbDirectory Then
                    ElseIf InStr(temp, ".") Then
                    Kill tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\" & temp
                End If
                temp = Dir$
            Loop
            DirList.Remove 1
        Loop
        If Dir(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name, vbDirectory) <> "" Then
            RemoveDirectory tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name
        End If
    End If
    
    Call StopDSP
    Call Stop_VisPlg
    
    WriteINI "Demension", "MainTop", Me.Top, strFileconfig
    WriteINI "Demension", "MainLeft", Me.Left, strFileconfig
    WriteINI "Equalizer", "EqualEnabled", bolEQEnabled, strFileconfig
    WriteINI "Equalizer", "LastPreset", strCurrentEQPreset, strFileconfig
    
    For i = 0 To 9
        intEQ(i) = sldEqua(i).value
        WriteINI "Equalizer", "Equa_" & i, intEQ(i), strFileconfig
    Next i
    
    If tAppConfig.bolSysTray Then sysTray.RemoveIcon (frmMedia.hwnd)
    If frmVisual.bolShow Then Unload frmVisual
    If frmDSP.bolShow Then Unload frmDSP
    If frmRate.bolShow Then Unload frmRate
    If frmEQLoadPreset.bolShow Then Unload frmEQLoadPreset
    If frmEQSavePreset.bolShow Then Unload frmEQSavePreset
    If frmLibrary.bolShow Then Unload frmLibrary
    If frmOption.bolShow Then Unload frmOption
    
    Unload frmApp
    Unload frmPlayList
    Unload frmMenu
    
    'Now upload file M3P.M3PData
    If FileExists(strLibrary) Then
        Name strLibrary As App.path & "\temp.M3PData"
        'Kill strLibrary
    End If
    
    i = FreeFile
    Open strLibrary For Output As #i
        Print #i, "[M3P_Library]"
        Print #i, "FileLen=" & UBound(Library)
    Close #i
    For i = 0 To UBound(Library) - 1
        Call WriteDataFile(i)
    Next i
    For i = 0 To UBound(Playlist) - 1
        Call WriteDataList(i)
    Next i
    
    Kill App.path & "\temp.M3PData"
    Exit Sub
handle:
    Kill strLibrary
    If FileExists(App.path & "\temp.M3PData") Then
        Name App.path & "\temp.M3PData" As strLibrary
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo handle
    
    Call GlobalFree(ByVal recPTR)
    Call BASS_RecordFree
    Call BASS_WA_FreeVisInfo
    Call BASS_WADSP_Free
    Call SaveConfig
    
    If tPlayerConfig.bolAutoShutdow Then
        Call WinAPI.SHUTDOWN
    End If

    End
handle:
    If Err.Number <> 0 Then
        Open App.path & "\Error.log" For Append As #1
            Print #1, "Date :" & date
            Print #1, "Error " & Err.Number & "_Des " & Err.Description
        Close #1
        End
    End If
End Sub
Private Sub lblMiniPos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If tPlayerConfig.bolTimer Then
            tPlayerConfig.bolTimer = False
            srlMiniInfor.Title = LangCap(141)
        Else
            tPlayerConfig.bolTimer = True
            srlMiniInfor.Title = LangCap(142)
        End If
    End If
End Sub

Private Sub lblMiniPos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
End Sub


Private Sub lblPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If tPlayerConfig.bolTimer Then
            tPlayerConfig.bolTimer = False
            srlInfor.Title = LangCap(141)
        Else
            tPlayerConfig.bolTimer = True
            srlInfor.Title = LangCap(142)
        End If
    End If
End Sub

Private Sub lblPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
End Sub




Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        Dim oldLeft As Long, oldTop As Long
        oldLeft = Me.Left
        oldTop = Me.Top
        DragForm Me
        frmMenu.Left = Me.Left
        Call MoveSnapForm(oldLeft, oldTop)
    End If
End Sub


Private Sub picMini_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        Dim oldLeft As Long, oldTop As Long
        oldLeft = Me.Left
        oldTop = Me.Top
        DragForm Me
        frmMenu.Left = Me.Left
        Call MoveSnapForm(oldLeft, oldTop)
    End If
End Sub

Private Sub picMiniEqua_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        intEQdownX = x
        intEQdownY = y
    End If
End Sub

Private Sub picMiniEqua_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        tPlayerConfig.bolShowEQ = ManualShowEQ(True, intEQdownX, intEQdownY, x, y)
        btnMiniMedia(7).bolOn = tPlayerConfig.bolShowEQ
        btnMiniMedia(7).Refresh
    End If
End Sub

Private Sub picMiniEqua_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMain, vbPopupMenuLeftAlign
    End If
End Sub
Private Sub picMiniMask_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        Dim oldLeft As Long, oldTop As Long
        oldLeft = Me.Left
        oldTop = Me.Top
        DragForm Me
        frmMenu.Left = Me.Left
        Call MoveSnapForm(oldLeft, oldTop)
    End If
End Sub

Private Sub picMiniMask_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMain, vbPopupMenuLeftAlign
    End If
End Sub





Private Sub Player_EndOfStream()
    If tDevice.SoundD.WaveWrite = True Then
        Call Play(currentPlayIndex + 1)
    End If
End Sub



Private Sub picEqua_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        intEQdownX = x
        intEQdownY = y
    End If
End Sub

Private Sub picEqua_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        tPlayerConfig.bolShowEQ = ManualShowEQ(False, intEQdownX, intEQdownY, x, y)
        btnMedia(7).bolOn = tPlayerConfig.bolShowEQ
        btnMedia(7).Refresh
    End If
End Sub

Private Sub picEqua_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMain, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub picMainMask_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Dim oldLeft As Long, oldTop As Long
        oldLeft = Me.Left
        oldTop = Me.Top
        DragForm Me
        frmMenu.Left = Me.Left
        Call MoveSnapForm(oldLeft, oldTop)
    End If
End Sub

Private Sub picMainMask_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMain, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub Player_PlayStateChange()
    Debug.Print Player.PlayState
    If Player.PlayState <> 0 Then
        If Player.PlayState = 1 Then
            btnMedia(1).bolOn = True
            btnMedia(2).bolOn = False
            btnMedia(2).Enabled = True
        Else
            btnMedia(2).bolOn = True
            btnMedia(2).Enabled = True
            btnMedia(1).bolOn = False
        End If
    Else
        btnMedia(1).bolOn = False
        btnMedia(2).bolOn = False
        btnMedia(2).Enabled = False
    End If
    btnMedia(1).Refresh
    btnMedia(2).Refresh
    btnMiniMedia(1).bolOn = btnMedia(1).bolOn
    btnMiniMedia(2).bolOn = btnMedia(2).bolOn
    btnMiniMedia(1).Refresh
    btnMiniMedia(2).Refresh
    If frmPlayList.List.ListItemCount <= 1 Then
        btnMedia(0).Enabled = False
        btnMedia(3).Enabled = False
        btnMiniMedia(0).Enabled = False
        btnMiniMedia(3).Enabled = False
    Else
        btnMedia(0).Enabled = True
        btnMedia(3).Enabled = True
        btnMiniMedia(0).Enabled = True
        btnMiniMedia(3).Enabled = True
    End If
End Sub

Private Sub sldAmp_Change(lValue As Long)
    srlInfor.Title = "PreAmp =" & CStr(sldAmp.value) & "db"
    lngPreAmpVal = lValue
    'Wait 200
    'Call BASS_ChannelSetDSP(Player.handle, AddressOf PreAmp, 0, 2)
End Sub

Private Sub sldAmp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
End Sub

Private Sub sldBalance_Change(lValue As Long)
    Dim strB As String
    If Not bolVideoOn Then
        Player.Balance = lValue
    Else
        frmVD.Video.Balance = lValue
    End If
    tPlayerConfig.intBalance = lValue
    If lValue = 0 Then
        srlInfor.Title = "[Center]"
    Else
        If lValue < 0 Then strB = "[Left " & Mid(CStr(lValue), 2) & "%]"
        If lValue > 0 Then strB = "[Right " & CStr(lValue) & "%]"
        srlInfor.Title = strB
    End If
    sldMiniBal.value = lValue
End Sub

Private Sub sldBalance_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
    If Button = vbRightButton Then
        sldBalance.value = 0
        Player.Balance = 0
    End If
End Sub

Private Sub sldEqua_Change(Index As Integer, lValue As Long)
    Dim strCenter As String
    
    Select Case Index
        Case 0: strCenter = "80 [Hz]"
        Case 1: strCenter = "180 [Hz]"
        Case 2: strCenter = "340 [Hz]"
        Case 3: strCenter = "650 [Hz]"
        Case 4: strCenter = "1 [Khz]"
        Case 5: strCenter = "3 [Khz]"
        Case 6: strCenter = "6 [Khz]"
        Case 7: strCenter = "12 [Khz]"
        Case 8: strCenter = "14 [Khz]"
        Case 9: strCenter = "16 [Khz]"
    End Select
    If sldEqua(Index).value > 0 Then strCenter = strCenter & "= +" & CStr(sldEqua(Index).value) & "db"
    If sldEqua(Index).value <= 0 Then strCenter = strCenter & "=" & CStr(sldEqua(Index).value) & "db"
    srlInfor.Title = strCenter
    sldMiniEqua(Index).value = sldEqua(Index).value
    Call UpdateEqual(Index, sldEqua(Index).value)
    strCurrentEQPreset = ""
End Sub

Private Sub sldEqua_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
End Sub

Private Sub sldMiniAmp_Change(lValue As Long)
    srlMiniInfor.Title = "PreAmp =" & CStr(sldMiniAmp.value) & "db"
    lngPreAmpVal = lValue
    'Wait 200
    'Call BASS_ChannelSetDSP(Player.handle, AddressOf PreAmp, 0, 2)
End Sub

Private Sub sldMiniAmp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
End Sub

Private Sub sldMiniBal_Change(lValue As Long)
    Dim strB As String
    If Not bolVideoOn Then
        Player.Balance = lValue
    Else
        frmVD.Video.Balance = lValue
    End If
    tPlayerConfig.intBalance = lValue
    If lValue = 0 Then
        srlMiniInfor.Title = "[Center]"
    Else
        If lValue < 0 Then strB = "[Left " & Mid(CStr(lValue), 2) & "%]"
        If lValue > 0 Then strB = "[Right " & CStr(lValue) & "%]"
        srlMiniInfor.Title = strB
    End If
    sldBalance.value = lValue
End Sub

Private Sub sldMiniBal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
    If Button = vbRightButton Then
        sldMiniBal.value = 0
        Player.Balance = 0
    End If
End Sub

Private Sub sldMiniEqua_Change(Index As Integer, lValue As Long)
    Dim strCenter As String
    Select Case Index
        Case 0: strCenter = "80 [Hz]"
        Case 1: strCenter = "180 [Hz]"
        Case 2: strCenter = "340 [Hz]"
        Case 3: strCenter = "650 [Hz]"
        Case 4: strCenter = "1 [Khz]"
        Case 5: strCenter = "3 [Khz]"
        Case 6: strCenter = "6 [Khz]"
        Case 7: strCenter = "12 [Khz]"
        Case 8: strCenter = "14 [Khz]"
        Case 9: strCenter = "16 [Khz]"
    End Select
    If sldMiniEqua(Index).value > 0 Then strCenter = strCenter & "= +" & CStr(sldMiniEqua(Index).value) & "db"
    If sldMiniEqua(Index).value <= 0 Then strCenter = strCenter & "=" & CStr(sldMiniEqua(Index).value) & "db"
    srlMiniInfor.Title = strCenter
    sldEqua(Index).value = sldMiniEqua(Index).value
    Call UpdateEqual(Index, sldEqua(Index).value)
    strCurrentEQPreset = ""
End Sub

Private Sub sldMiniEqua_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
End Sub

Private Sub sldMiniPosition_Change(lValue As Long)
    If Not bolVideoOn Then
        If Player.PlayState = 1 Then
            Player.Position = lValue
        End If
    Else
        If frmVD.Video.State = playing Then
            frmVD.Video.Position = lValue
        End If
    End If
    srlMiniInfor.Title = Time2String(lValue)
End Sub

Private Sub sldMiniPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
End Sub

Private Sub sldMiniVol_Change(lValue As Long)
    Call volume(lValue)
End Sub

Private Sub sldMiniVol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlMiniInfor.Title = strTitle
    End If
End Sub


Private Sub sldPosition_Change(lValue As Long)
    If Not bolVideoOn Then
        If Player.PlayState = 1 Then
            Player.Position = lValue
        End If
    Else
        If frmVD.Video.State = playing Then
            frmVD.Video.Position = lValue
        End If
    End If
    srlInfor.Title = Time2String(lValue)
End Sub

Private Sub sldPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
End Sub

Private Sub sldVolume_Change(lValue As Long)
    Call volume(lValue)
End Sub

Private Sub sldVolume_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        srlInfor.Title = strTitle
    End If
End Sub

Private Sub tmrPos_Timer()
    If Not bolVideoOn Then
        If Player.PlayState = 1 Then
            sldPosition.value = Player.Position
            sldMiniPosition.value = Player.Position
        End If
        If tDevice.SoundD.WaveWrite = False Then
            If tPlayerConfig.bolCrossfade = False Then
                If sldPosition.value = sldPosition.max Then Call PlayNext
            Else
                If sldPosition.value = sldPosition.max - tPlayerConfig.intCrossfade Then Call PlayNext
            End If
        End If
    Else
        If frmVD.Video.State = playing Then
            sldPosition.value = frmVD.Video.Position
        End If
        sldMiniPosition.value = sldPosition.value
        If sldPosition.value = sldPosition.max - tPlayerConfig.intCrossfade Then Call PlayNext
    End If
End Sub
Private Sub tmrPlay_Timer()
    On Error Resume Next
    Dim TimeCurrent As Integer
    Dim min, sec, MinValue, SecValue As Integer
    
    If Not bolVideoOn Then
        If Player.PlayState = 1 Then
            TimeCurrent = Player.Position
            min = TimeCurrent \ 60
            sec = TimeCurrent - (min * 60)
            If sec = "-1" Then sec = "0"
            MinValue = (Player.Duration - TimeCurrent) \ 60
            SecValue = (Player.Duration - TimeCurrent) - MinValue * 60
            If Not tPlayerConfig.bolTimer Then
                lblPosition.Caption = Format(min, "00") & ":" & Format(sec, "00")
            Else
                lblPosition.Caption = "-" & Format(MinValue, "00") & ":" & Format(SecValue, "00")
            End If
        End If
    Else
        If frmVD.Video.State = playing Then
            TimeCurrent = frmVD.Video.Position
            min = TimeCurrent \ 60
            sec = TimeCurrent - (min * 60)
            If sec = "-1" Then sec = "0"
            MinValue = (frmVD.Video.Duration - TimeCurrent) \ 60
            SecValue = (frmVD.Video.Duration - TimeCurrent) - MinValue * 60
            If Not tPlayerConfig.bolTimer Then
                lblPosition.Caption = Format(min, "00") & ":" & Format(sec, "00")
            Else
                lblPosition.Caption = "-" & Format(MinValue, "00") & ":" & Format(SecValue, "00")
            End If
        End If
    End If
    lblMiniPos.Caption = lblPosition.Caption
End Sub
Private Sub tmrTaskbar_Timer()
    Dim temp$
    Dim Lenght As Integer
    Dim str As String
    If tAppConfig.bolTaskbar And tAppConfig.bolTaskbarScroll Then
        If btnMedia(1).bolOn Or btnMedia(2).bolOn Then
            If Not bolCDPlay Then
                temp$ = "*** " & tCurrentTrack.Artist & " _ " & tCurrentTrack.Title & " "
            Else
                temp$ = "*** " & frmPlayList.List.ListItemText(tCDplay.CurrentTrack + 1) & " "
            End If
        Else
            temp$ = "*** " & "MP3_proPlayer " & App.Major & "." & App.Minor & " "
        End If
        Lenght = Len(temp)
        If posTaskbar < Lenght Then
            str = Mid(temp$, IIf(posTaskbar = 0, 1, posTaskbar), Lenght - posTaskbar) & " " & Mid(temp$, 1, posTaskbar)
            frmMenu.Caption = "[" & lblPosition.Caption & "] " & str
            posTaskbar = posTaskbar + 1
        Else
            posTaskbar = 1
        End If
    Else
        If tAppConfig.bolTaskbarScroll Then
            If Not bolCDPlay Then
                frmMenu.Caption = tCurrentTrack.Artist & " _ " & tCurrentTrack.Title
            Else
                frmMenu.Caption = frmPlayList.List.ListItemText(tCDplay.CurrentTrack + 1)
            End If
        End If
    End If
End Sub


Private Sub tmrTestPic_Timer()
    On Error Resume Next
    If Not bolCDPlay Then
        If currentPlayIndex <> 0 Then
            strTitle = tCurrentTrack.Artist & " -- " & tCurrentTrack.Title
        Else
            If currentIndex <> 0 Then
                strTitle = NowPlaying(frmPlayList.List.Key(currentIndex)).Infor.Artist & " -- " & NowPlaying(frmPlayList.List.Key(currentIndex)).Infor.Title
            Else
                strTitle = "HanaSoft@--MP3_proPlayer--Version " & App.Major & "." & App.Minor & "." & App.Revision
            End If
        End If
    Else
        If currentPlayIndex <> 0 Then
            strTitle = frmPlayList.List.ListItemText(tCDplay.CurrentTrack + 1)
        Else
            strTitle = "HanaSoft@--MP3_proPlayer--Version " & App.Major & "." & App.Minor & "." & App.Revision
        End If
    End If
End Sub

Private Sub tmrVisual_Timer()
    If Not bolVideoOn Then
        If btnMedia(1).bolOn Then
            If tCurrentSkin.mini Then
                Call VUMeter
            Else
                Select Case tSkinVis.intStyle
                    Case Is = 0
                        Call Oscilliscope
                    Case Is = 1
                        Call Spectrum
                    Case Is = 2
                        Exit Sub
                    End Select
            End If
        End If
    End If
End Sub
Public Sub PlayNext()
    On Error Resume Next
    Dim NextIndex As Long
    Dim tmpIndex As Long
    
    tmpIndex = currentPlayIndex
    
    If tPlayerConfig.bolAutoRemove Then
        Call SubFile(tmpIndex)
    End If

    If Not bolCDPlay Then
        If tPlayerConfig.bolRepeat = False Then
            If tPlayerConfig.bolShuffe = True Then
                Randomize frmPlayList.List.ListItemCount
                NextIndex = Int(frmPlayList.List.ListItemCount * Rnd)
                If NextIndex < frmPlayList.List.ListItemCount And NextIndex > 0 Then
                    Call Play(NextIndex)
                Else
                    NextIndex = currentIndex
                    Call Play(NextIndex)
                End If
            Else
                If currentPlayIndex < frmPlayList.List.ListItemCount Then
                    NextIndex = currentPlayIndex + 1
                    Call Play(NextIndex)
                Else
                'Loop playlist
                    If tPlayerConfig.bolLoop Then
                        Call Play(1)
                    Else
                        Call StopPlayer
                        NextIndex = 0
                        If tPlayerConfig.bolAutoExit Then
                            Unload Me
                        End If
                    End If
                End If
            End If
        Else
            NextIndex = currentPlayIndex
            Call Play(NextIndex)
        End If
    Else
        If tPlayerConfig.bolRepeat = False Then
            If tPlayerConfig.bolShuffe = True Then
                Randomize tCDplay.TotalTrack
                NextIndex = Int(tCDplay.TotalTrack * Rnd)
                If NextIndex < tCDplay.TotalTrack And NextIndex > 0 Then
                    Call PlayCD(tCDplay.CurrentDrive, NextIndex)
                Else
                    NextIndex = currentIndex
                    Call PlayCD(tCDplay.CurrentDrive, NextIndex)
                End If
            Else
                If tCDplay.CurrentTrack < tCDplay.TotalTrack Then
                    NextIndex = tCDplay.CurrentTrack + 1
                    Call PlayCD(tCDplay.CurrentDrive, NextIndex)
                Else
                    'Loop playlist
                    If tPlayerConfig.bolLoop Then
                        Call PlayCD(tCDplay.CurrentDrive, 0)
                    Else
                        Call StopPlayer
                        If tPlayerConfig.bolAutoExit Then
                            Unload Me
                        End If
                    End If
                End If
            End If
        Else
            NextIndex = tCDplay.CurrentTrack
            Call PlayCD(tCDplay.CurrentDrive, NextIndex)
        End If
    End If
End Sub
Sub Spectrum()
    Dim lRslt As Long
    channel = Player.handle
    lRslt = BASS_ChannelGetData(channel, fft(0), BASS_DATA_FFT1024)
    If lRslt <> BASSFALSE And lRslt > BASSFALSE Then
        Vis.Spectrum fft
    End If
End Sub
Sub Oscilliscope()
    channel = Player.handle
    BASS_ChannelGetData channel, SampleData(0), 1000
    Vis.Oscilliscope SampleData
End Sub
Public Sub PlayCrossFade()
    On Error Resume Next
    PlayerBack.volume = Player.volume
    If Player.PlayState = 1 Then
        If currentPlayIndex <> 0 Then
            If tPlayerConfig.bolCrossfade Then
                If Not bolCDPlay Then
                    PlayerBack.OpenFile NowPlaying(frmPlayList.List.Key(CLng(currentPlayIndex))).Infor.FullName, False
                    PlayerBack.Position = Player.Position
                    PlayerBack.PlayStream
                    BASS_ChannelSlideAttributes PlayerBack.handle, -1, -2, -101, tPlayerConfig.intCrossfade * 1000
                End If
            End If
        End If
    End If
End Sub
Public Sub volume(lngValue As Long) 'lngValue in range 0-100
    If Not bolVideoOn Then
        Player.volume = lngValue
        tPlayerConfig.intVolume = Player.volume
    Else
        frmVD.Video.volume = lngValue
        tPlayerConfig.intVolume = frmVD.Video.volume
    End If
    srlInfor.Title = "[" & LangCap(143) & "= " & (tPlayerConfig.intVolume) & " %]"
    srlMiniInfor.Title = "[" & LangCap(143) & "= " & (tPlayerConfig.intVolume) & " %]"
    sldVolume.value = lngValue
    sldMiniVol.value = lngValue
End Sub
Public Sub btnPlayClick()
    If frmPlayList.List.ListItemCount = 0 Then
        If OpenGetFile Then
            Call Play(1)
        Else
            Exit Sub
        End If
    End If
    If currentIndex <> currentPlayIndex Then
        Call Play(currentIndex)
    Else
        If Not bolVideoOn Then
            If Player.PlayState = 2 Then 'Paused
                Player.Position = sldPosition.value
                Player.PlayStream
                frmMenu.mnuMediaControl(7).Checked = False
            End If
            If Player.PlayState = 0 Then 'Stoped
                Play (currentIndex)
            End If
        Else
            If frmVD.Video.State = Paused Then 'Paused and play from pos
                frmVD.Video.PlayVideo sldPosition.value
                frmMenu.mnuMediaControl(7).Checked = False
            End If
        End If
    End If
End Sub
Public Sub btnPauseClick()
    If currentPlayIndex <> 0 Then
        If Not bolVideoOn Then
            If Player.PlayState = 2 Then 'Paused and play from pos
                Player.Position = sldPosition.value
                Player.PlayStream
                frmMenu.mnuMediaControl(7).Checked = False
            Else 'set paused
                Player.PauseStream
                frmMenu.mnuMediaControl(7).Checked = True
            End If
        Else
            If frmVD.Video.State = Paused Then 'Paused so play from pos
                frmVD.Video.PlayVideo sldPosition.value
                frmMenu.mnuMediaControl(7).Checked = False
            Else 'set pause
                frmVD.Video.PauseVideo
                frmMenu.mnuMediaControl(7).Checked = True
            End If
        End If
    End If
End Sub
Public Sub btnShuffeClick()
    tPlayerConfig.bolShuffe = Not tPlayerConfig.bolShuffe
    frmMenu.mnuOptionC(5).Checked = tPlayerConfig.bolShuffe
    btnMedia(5).bolOn = tPlayerConfig.bolShuffe
    btnMiniMedia(5).bolOn = tPlayerConfig.bolShuffe
End Sub
Public Sub btnRepeatClick()
    tPlayerConfig.bolRepeat = Not tPlayerConfig.bolRepeat
    frmMenu.mnuOptionC(6).Checked = tPlayerConfig.bolRepeat
    btnMedia(6).bolOn = tPlayerConfig.bolRepeat
    btnMiniMedia(6).bolOn = tPlayerConfig.bolRepeat
End Sub
Public Sub btnShowEQClick()
    tPlayerConfig.bolShowEQ = Not tPlayerConfig.bolShowEQ
    btnMedia(7).bolOn = tPlayerConfig.bolShowEQ
    btnMiniMedia(7).bolOn = tPlayerConfig.bolShowEQ
    frmMenu.mnuShowEQ.Checked = tPlayerConfig.bolShowEQ
    Call ShowEQ
End Sub
Public Sub btnShowPLClick()
    tPlaylistConfig.bolHidePL = Not tPlaylistConfig.bolHidePL
    btnMedia(8).bolOn = Not tPlaylistConfig.bolHidePL
    btnMiniMedia(8).bolOn = Not tPlaylistConfig.bolHidePL
    If tPlaylistConfig.bolHidePL Then
        frmPlayList.Hide
    Else
        frmPlayList.Show
    End If
    frmMenu.mnuShowPL.Checked = Not tPlaylistConfig.bolHidePL
End Sub

Public Sub Init()
    On Error Resume Next
    
    Load frmApp
    
    'Install device
    Dim intDevice As Integer
    intDevice = 1
    
    While BASS_GetDeviceDescription(intDevice)
        Player.InitBass intDevice, tDevice.SoundD.Freq, 0, 0
        PlayerBack.InitBass intDevice, tDevice.SoundD.Freq, 0, 0
        intDevice = intDevice + 1
    Wend
    If tDevice.SoundD.WaveWrite = False Then
        If Player.SetDevice(tDevice.SoundD.OutputDevice) = False Or PlayerBack.SetDevice(tDevice.SoundD.OutputDevice) = False Then
            MsgBox "Install device failed !!!", vbCritical, "MP3_proPlayer"
        End If
    Else
        If Player.InitBass(0, 44100, 0, 0) = False Then
            MsgBox "Install device for wave write failed !!!", vbCritical, "MP3_proPlayer"
        End If
        If BASS_RecordInit(0) = False Then
            MsgBox "Unable to write video to wav", vbCritical, "Error"
        Else
            ' get list of inputs
            Dim c As Integer
            Dim T As String
            input_ = -1
            'Simple to write a video is use bass record with input line is Stereo Mix
            While BASS_RecordGetInputName(c)
                T = LCase(VBStrFromAnsiPtr(BASS_RecordGetInputName(c)))
                If T = "stereo mix" Then
                    input_ = c
                    Call BASS_RecordSetInput(c, BASS_INPUT_ON) ' enable the selected input
                End If
                c = c + 1
            Wend
            recPTR = 0
            reclen = 0
            BUFSTEP = 200000    ' memory allocation unit
        End If
    End If
        
    If Not FileExists(GetProperPath(App.path & "\Plugins") & "bass_wa.dll") Then
        Call MsgBox("BASS_WA.dll does not exists and you can't use winamp vis plugin", vbCritical, "MP3_proPlayer")
    Else
        Call BASS_WA_SetHwnd(Player.hwnd)
    End If
        
    If Not FileExists(GetProperPath(App.path & "\Plugins") & "bass_wadsp.dll") Then
        Call MsgBox("BASS_WA.dll does not exists and you can't use winamp dsp plugin", vbCritical, "MP3_proPlayer")
    Else
        BASS_WADSP_Init Player.hwnd
    End If
    
    
    
    Me.Left = ReadINI("Demension", "Mainleft", strFileconfig)
    Me.Top = ReadINI("Demension", "MainTop", strFileconfig)
    
        
        
    ReDim DSPPlugins(0) 'workaround
        
    For i = 1 To 12
        imgIcon.ListImages.Add i, , LoadResPicture(100 + i, vbResIcon)
    Next i
    Me.Icon = imgIcon.ListImages(tAppConfig.intIcon).Picture
        
    If tAppConfig.bolSysTray Then
        sysTray.AddIcon Me.hwnd, Me.Icon.handle
    Else
        sysTray.RemoveIcon Me.hwnd
    End If
        
    Load frmOption
    frmOption.Visible = False
    
    tmrVisual.Interval = tSkinVis.intRefresh
    Vis.SpecDrawMode = tSkinVis.intSpecDraw
    Vis.SpecFillMode = tSkinVis.intSpecFill
    Vis.SpecShowPeak = tSkinVis.bolSpecPeak
    Vis.SpecPeakDelay = tSkinVis.intSpecPeakPause
    Vis.SpecPeakDrop = tSkinVis.intSpecPeakDrop
    Vis.OscDrawMode = tSkinVis.intOsc
        
        
    If strCurrentEQPreset = "" Then
        For i = 0 To 9
            sldEqua(i).value = intEQ(i)
            sldMiniEqua(i).value = intEQ(i)
        Next i
    Else
        Call LoadEQPreset(strCurrentEQPreset)
    End If
        
    tPlayerConfig.lngMasterVol = BASS_GetVolume 'soundcard master vol
    sldBalance.value = tPlayerConfig.intBalance
    sldMiniBal.value = tPlayerConfig.intBalance
    sldVolume.value = tPlayerConfig.intVolume
    sldMiniVol.value = tPlayerConfig.intVolume
    sldAmp.value = 0 ' intAmp
        
    With frmDSP
        For i = 0 To 4
            .prgChorus(i).value = intChorus(i)
        Next i
        .cboChorus(0).ListIndex = intChorus(5)
        .cboChorus(1).ListIndex = intChorus(6)
            
        For i = 0 To 5
            .prgCompressor(i).value = intCompressor(i)
        Next i
            
        For i = 0 To 4
            .prgDistortion(i).value = intDistortion(i)
        Next i
            
        For i = 0 To 3
            .prgEcho(i).value = intEcho(i)
        Next i
        .chkEcho.value = intEcho(4)
            
        For i = 0 To 4
            .prgFlanger(i).value = intFlanger(i)
        Next i
        .cboFlanger(0).ListIndex = intFlanger(5)
        .cboFlanger(1).ListIndex = intFlanger(6)
        
        .prgGargle.value = intGargle(0)
        .cboGargle.ListIndex = intGargle(1)
        
        For i = 0 To 11
            .prgLReverb(i).value = intI3DL2Reverb(i)
        Next i
        
        For i = 0 To 3
            .prgReverb(i).value = intReverb(i)
        Next i
        
    End With
    
    Call SetFX(Player)
    For i = 0 To 7
        Call UpdateFX(i)
    Next i
                
    srlInfor.Title = strTitle
    srlMiniInfor.Title = strTitle
    srlInfor.Scroll = tPlayerConfig.bolScroll
    srlMiniInfor.Scroll = tPlayerConfig.bolScroll
        
    If LibOption.bolUse Then
        Call LoadDatabase
    End If
        
    ReDim NowPlaying(0)
        
    If tPlayerConfig.bolShowList = True Then
        If FileExists(App.path & "\M3P.m3u") Then
            With frmPlayList
                Call LoadPlaylistM3U(App.path & "\M3P.m3u")
                currentIndex = tPlaylistConfig.intLastFile
                .sldPl.value = currentIndex
            End With
            If tPlayerConfig.bolAutoPlay Then Play (currentIndex)
        End If
    Else
        currentIndex = 0
        currentRIndex = 0
        currentPlayIndex = 0
    End If
    
End Sub

Private Sub Vis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tSkinVis.intStyle = tSkinVis.intStyle + 1
        If tSkinVis.intStyle = 3 Then tSkinVis.intStyle = 0
        For i = 0 To 2
            frmMenu.mnuMainSpecC(i).Checked = False
        Next
        frmMenu.mnuMainSpecC(tSkinVis.intStyle).Checked = True
        frmOption.optSkinVis(tSkinVis.intStyle).value = True
        Vis.StyleVis = tSkinVis.intStyle
        Select Case tSkinVis.intStyle
            Case 0, 1
                tmrVisual.Enabled = True
            Case 2
                tmrVisual.Enabled = False
                Vis.doStop
        End Select
    End If
End Sub

Private Sub Vis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuMainSpec, vbPopupMenuLeftAlign
    End If
End Sub
