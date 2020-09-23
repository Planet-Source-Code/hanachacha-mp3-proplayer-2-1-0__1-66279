VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOption 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " M3P : Preferences"
   ClientHeight    =   5235
   ClientLeft      =   2820
   ClientTop       =   4740
   ClientWidth     =   6870
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   0
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   4
      Top             =   750
      Width           =   6540
      Begin VB.ComboBox cboLang 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   720
         Width           =   2055
      End
      Begin ComctlLib.Slider sldGeneral 
         Height          =   255
         Left            =   600
         TabIndex        =   131
         Top             =   3180
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Max             =   11
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Exit when end list"
         Height          =   255
         Index           =   10
         Left            =   4320
         TabIndex        =   63
         Top             =   2928
         Width           =   1575
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Remove file played"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   62
         Top             =   2500
         Width           =   1695
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Auto shutdown"
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   61
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Remember lastlist"
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   60
         Top             =   2072
         Width           =   1575
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Auto play on start"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   59
         Top             =   1644
         Width           =   1575
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Enabled scroll title in taskbar"
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   58
         Top             =   2080
         Width           =   2415
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Show context menu"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   57
         Top             =   788
         Width           =   1815
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "System tray"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   36
         Top             =   2520
         Width           =   1150
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Taskbar"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   35
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Show splash screen"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Start with window"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   32
         Top             =   1216
         Width           =   1695
      End
      Begin VB.Label lblCaptionOpt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Language :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   164
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblCaptionOpt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Application"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lblCaptionOpt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Show MP3_proPlayer in :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   37
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label lblCaptionOpt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "System tray icon"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   38
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Shape shpGeneral 
         BorderColor     =   &H80000010&
         Height          =   615
         Index           =   2
         Left            =   480
         Shape           =   4  'Rounded Rectangle
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Image imgGeneral 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   495
      End
      Begin VB.Shape shpGeneral 
         BorderColor     =   &H80000010&
         Height          =   3615
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   6
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   90
      Top             =   750
      Width           =   6540
      Begin VB.PictureBox picVis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Index           =   0
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   6540
         TabIndex        =   111
         Top             =   480
         Width           =   6540
         Begin VB.PictureBox picSkinVis 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Index           =   0
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   6255
            TabIndex        =   126
            Top             =   0
            Width           =   6255
            Begin ComctlLib.Slider sldSkinVis 
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   135
               Top             =   600
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   450
               _Version        =   327682
               LargeChange     =   30
               SmallChange     =   30
               Min             =   10
               Max             =   100
               SelStart        =   70
               TickFrequency   =   30
               Value           =   70
            End
            Begin VB.OptionButton optSkinVis 
               Caption         =   "Oscilliscope"
               Height          =   255
               Index           =   0
               Left            =   2520
               TabIndex        =   129
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton optSkinVis 
               Caption         =   "Spectrum"
               Height          =   255
               Index           =   1
               Left            =   3840
               TabIndex        =   128
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton optSkinVis 
               Caption         =   "None"
               Height          =   255
               Index           =   2
               Left            =   5160
               TabIndex        =   127
               Top             =   120
               Width           =   735
            End
            Begin VB.Label lblCaptionOpt6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Visualization refresh rate"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   130
               Top             =   600
               Width           =   1710
            End
            Begin VB.Shape shpSkinVis 
               BorderColor     =   &H80000010&
               Height          =   1095
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   6255
            End
         End
         Begin VB.PictureBox picSkinVis 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   3
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   6255
            TabIndex        =   122
            Top             =   2640
            Width           =   6255
            Begin VB.OptionButton optOscStyle 
               Caption         =   "Solid"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   125
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optOscStyle 
               Caption         =   "Line"
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   124
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optOscStyle 
               Caption         =   "Dot"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   123
               Top             =   240
               Width           =   855
            End
            Begin VB.Shape shpSkinVis 
               BorderColor     =   &H80000010&
               Height          =   495
               Index           =   2
               Left            =   0
               Top             =   120
               Width           =   6255
            End
         End
         Begin VB.PictureBox picSkinVis 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   1
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   6255
            TabIndex        =   112
            Top             =   1320
            Width           =   6255
            Begin VB.PictureBox picSkinVis 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   240
               ScaleHeight     =   255
               ScaleWidth      =   2535
               TabIndex        =   119
               Top             =   840
               Width           =   2535
               Begin VB.OptionButton optSpecDraw 
                  Caption         =   "Thick"
                  Height          =   255
                  Index           =   0
                  Left            =   1680
                  TabIndex        =   121
                  Top             =   0
                  Width           =   735
               End
               Begin VB.OptionButton optSpecDraw 
                  Caption         =   "Thin"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   120
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.OptionButton optSpecStyle 
               Caption         =   "Line"
               Height          =   255
               Index           =   2
               Left            =   1920
               TabIndex        =   116
               Top             =   120
               Width           =   735
            End
            Begin VB.OptionButton optSpecStyle 
               Caption         =   "Fire"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   115
               Top             =   120
               Width           =   615
            End
            Begin VB.OptionButton optSpecStyle 
               Caption         =   "Normal"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   114
               Top             =   120
               Width           =   855
            End
            Begin VB.CheckBox chkSpec 
               Caption         =   "Show peak"
               Height          =   255
               Left            =   240
               TabIndex        =   113
               Top             =   480
               Width           =   1215
            End
            Begin ComctlLib.Slider sldSkinVis 
               Height          =   255
               Index           =   1
               Left            =   3840
               TabIndex        =   136
               Top             =   120
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   450
               _Version        =   327682
               SmallChange     =   5
               Min             =   5
               Max             =   25
               SelStart        =   25
               TickFrequency   =   5
               Value           =   25
            End
            Begin ComctlLib.Slider sldSkinVis 
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   137
               Top             =   840
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   450
               _Version        =   327682
               LargeChange     =   1
               Min             =   1
               Max             =   5
               SelStart        =   1
               Value           =   1
            End
            Begin VB.Label lblCaptionOpt6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peak drop"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   10
               Left            =   3000
               TabIndex        =   118
               Top             =   840
               Width           =   735
            End
            Begin VB.Label lblCaptionOpt6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peak pause"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   9
               Left            =   2880
               TabIndex        =   117
               Top             =   120
               Width           =   855
            End
            Begin VB.Shape shpSkinVis 
               BorderColor     =   &H80000010&
               Height          =   1215
               Index           =   1
               Left            =   0
               Top             =   0
               Width           =   6255
            End
         End
      End
      Begin VB.PictureBox picVis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Index           =   1
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   6540
         TabIndex        =   91
         Top             =   480
         Width           =   6540
         Begin ComctlLib.Slider sldMainVis 
            Height          =   255
            Left            =   720
            TabIndex        =   138
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327682
            LargeChange     =   10
            SmallChange     =   10
            Min             =   10
            Max             =   100
            SelStart        =   90
            TickFrequency   =   10
            Value           =   90
         End
         Begin VB.CommandButton cmdMainVis 
            Caption         =   "..."
            Height          =   330
            Left            =   3210
            TabIndex        =   97
            Top             =   2760
            Width           =   495
         End
         Begin VB.CommandButton cmdVisualization 
            Caption         =   "Save"
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   102
            Top             =   2760
            Width           =   735
         End
         Begin VB.CommandButton cmdVisualization 
            Caption         =   "Stop"
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   101
            Top             =   2760
            Width           =   735
         End
         Begin VB.CommandButton cmdVisualization 
            Caption         =   "Start"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   100
            Top             =   2760
            Width           =   735
         End
         Begin VB.ListBox lstAVS 
            Height          =   1815
            Left            =   3840
            TabIndex        =   99
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtMainVis 
            Height          =   375
            Left            =   240
            TabIndex        =   98
            Top             =   2745
            Width           =   3495
         End
         Begin VB.CheckBox chkMainVis 
            Caption         =   "Show title"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   96
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox chkMainVis 
            Caption         =   "Use picture background"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   95
            Top             =   2280
            Width           =   2175
         End
         Begin VB.ComboBox cboData 
            Height          =   315
            ItemData        =   "frmOptions.frx":000C
            Left            =   1320
            List            =   "frmOptions.frx":001C
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "Spectrum"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   93
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "Oscillscope"
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   92
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblCaptionOpt6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Visualization"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   106
            Top             =   0
            Width           =   870
         End
         Begin VB.Label lblCaptionOpt6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preset :"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   2
            Left            =   3840
            TabIndex        =   108
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblCaptionOpt6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   107
            Top             =   360
            Width           =   345
         End
         Begin VB.Shape shpVisual 
            BorderColor     =   &H80000010&
            Height          =   3135
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   6255
         End
         Begin VB.Label lblMainVis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1320
            TabIndex        =   105
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label lblCaptionOpt6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Background :"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   104
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label lblCaptionOpt6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Value :"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   103
            Top             =   1140
            Width           =   885
         End
      End
      Begin VB.Label lblVisTab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MVS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   110
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblVisTab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Skin Vis"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   109
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   5
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   5
      Top             =   750
      Width           =   6540
      Begin VB.CommandButton cmdSkin 
         Caption         =   "Author"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   89
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CommandButton cmdSkin 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   88
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSkin 
         Caption         =   "Rename"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   87
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdSkin 
         Caption         =   "Change"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   86
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkSkin 
         Caption         =   "Enable Equalizer Panel Slide"
         Height          =   375
         Left            =   240
         TabIndex        =   85
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtSkin 
         Height          =   2775
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   84
         Top             =   360
         Width           =   2295
      End
      Begin VB.ListBox lstSkin 
         Height          =   2790
         ItemData        =   "frmOptions.frx":0036
         Left            =   240
         List            =   "frmOptions.frx":0038
         TabIndex        =   83
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblCaptionOpt4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Skin"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   46
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape shpSkin 
         BorderColor     =   &H80000010&
         Height          =   3615
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   600
      Pattern         =   "*.mvs"
      TabIndex        =   140
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   3
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   6
      Top             =   750
      Width           =   6540
      Begin VB.TextBox txtPlaylist 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   55
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtPlaylist 
         Height          =   1215
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   51
         Text            =   "frmOptions.frx":003A
         Top             =   2400
         Width           =   4935
      End
      Begin VB.CommandButton cmdPlaylist 
         Caption         =   "Default"
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtPlaylist 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   6015
      End
      Begin VB.CheckBox chkPlaylist 
         Caption         =   "Show number in playlist"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton optPlaylist 
         Caption         =   "Load infor only when file played"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton optPlaylist 
         Caption         =   "Load infor when file view (Recommended)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   540
         Width           =   3375
      End
      Begin VB.OptionButton optPlaylist 
         Caption         =   "Load infor when file added (So slow if add lot file)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblCaptionOpt3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Row Scroll (<25)"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   56
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label lblCaptionOpt3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Format Playlist Text :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Shape shpPlaylist 
         BorderColor     =   &H80000010&
         Height          =   1935
         Index           =   1
         Left            =   120
         Top             =   1800
         Width           =   6255
      End
      Begin VB.Label lblCaptionOpt3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Information"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   0
         Width           =   780
      End
      Begin VB.Shape shpPlaylist 
         BorderColor     =   &H80000010&
         Height          =   1455
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6255
      End
   End
   Begin MSComctlLib.ImageList imlIcon 
      Left            =   0
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdloColor 
      Left            =   7800
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   4800
      Width           =   1305
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   4800
      Width           =   1305
   End
   Begin ComctlLib.TabStrip tbsOption 
      Height          =   4545
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   8017
      MultiRow        =   -1  'True
      TabFixedWidth   =   2646
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   10
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Device"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "File types"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Playlist"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Library"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Skin"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Visualization"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Plugins"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Winamp Plugins"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "MP3 _ proPlayer"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   4
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   67
      Top             =   720
      Width           =   6540
      Begin VB.CheckBox chkLibrary 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdLibrary 
         Caption         =   "Monitor"
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   80
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox lstLibrary 
         Height          =   1425
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   72
         Top             =   720
         Width           =   4815
      End
      Begin VB.CommandButton cmdLibrary 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   71
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdLibrary 
         Caption         =   "Remove"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   70
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtLibrary 
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   69
         Top             =   2670
         Width           =   735
      End
      Begin VB.TextBox txtLibrary 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   68
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Monitor Folder"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   79
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Library Option"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   78
         Top             =   2400
         Width           =   975
      End
      Begin VB.Shape shpLibrary 
         BorderColor     =   &H80000010&
         Height          =   2175
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6255
      End
      Begin VB.Shape shpLibrary 
         BorderColor     =   &H80000010&
         Height          =   1215
         Index           =   1
         Left            =   120
         Top             =   2520
         Width           =   6255
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skip file(s) smaller than :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   77
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audio :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   76
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Video :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   75
         Top             =   3330
         Width           =   495
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(KB)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   3600
         TabIndex        =   74
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label lblCaptionOpt5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(KB)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   3600
         TabIndex        =   73
         Top             =   3330
         Width           =   300
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   7
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   50
      Top             =   750
      Width           =   6540
      Begin VB.CommandButton cmdGenPlugins 
         Caption         =   "About"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   161
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdGenPlugins 
         Caption         =   "Config"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   159
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ListBox lstGenPlugins 
         Height          =   2205
         Left            =   360
         TabIndex        =   158
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label lblPlugins 
         AutoSize        =   -1  'True
         Caption         =   "M3P_Plugins (just beta test only)"
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   360
         TabIndex        =   160
         Top             =   120
         Width           =   2295
      End
      Begin VB.Shape shpPlugins 
         BorderColor     =   &H80000010&
         Height          =   3495
         Left            =   120
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   2
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   34
      Top             =   750
      Width           =   6540
      Begin ComctlLib.ListView lvwPlayer 
         Height          =   2775
         Left            =   240
         TabIndex        =   139
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "FileExt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "CurrentProgram"
            Object.Width           =   5292
         EndProperty
      End
      Begin ComctlLib.Slider sldIcon 
         Height          =   975
         Index           =   0
         Left            =   5280
         TabIndex        =   133
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         _Version        =   327682
         Orientation     =   1
         Min             =   1
         Max             =   12
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "Register"
         Height          =   375
         Index           =   4
         Left            =   5160
         TabIndex        =   52
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "Video Only"
         Height          =   375
         Index           =   3
         Left            =   3930
         TabIndex        =   43
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "Audio Only"
         Height          =   375
         Index           =   2
         Left            =   2700
         TabIndex        =   42
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "Select None"
         Height          =   375
         Index           =   1
         Left            =   1470
         TabIndex        =   41
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "Select All"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   3240
         Width           =   1095
      End
      Begin MSComctlLib.ImageList imgProgram 
         Left            =   5760
         Top             =   3840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
      End
      Begin ComctlLib.Slider sldIcon 
         Height          =   975
         Index           =   1
         Left            =   5280
         TabIndex        =   134
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         _Version        =   327682
         Orientation     =   1
         Min             =   1
         Max             =   12
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblCaptionOpt2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Playlist Icon"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   6
         Left            =   5280
         TabIndex        =   54
         Top             =   1800
         Width           =   840
      End
      Begin VB.Shape shpPlayer 
         BorderColor     =   &H80000010&
         Height          =   1215
         Index           =   5
         Left            =   5160
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblCaptionOpt2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "File Icon"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   5280
         TabIndex        =   53
         Top             =   240
         Width           =   600
      End
      Begin VB.Shape shpPlayer 
         BorderColor     =   &H80000010&
         Height          =   1215
         Index           =   4
         Left            =   5160
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblCaptionOpt2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Files type"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   39
         Top             =   0
         Width           =   660
      End
      Begin VB.Shape shpPlayer 
         BorderColor     =   &H80000010&
         Height          =   3615
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.PictureBox picOption 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   8
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   141
      Top             =   750
      Width           =   6540
      Begin VB.PictureBox picWinamp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3015
         Index           =   0
         Left            =   420
         ScaleHeight     =   3015
         ScaleWidth      =   5820
         TabIndex        =   142
         Top             =   480
         Width           =   5820
         Begin VB.ListBox lstPlugins 
            Height          =   1425
            ItemData        =   "frmOptions.frx":0129
            Left            =   240
            List            =   "frmOptions.frx":012B
            Sorted          =   -1  'True
            TabIndex        =   148
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Start"
            Height          =   375
            Index           =   0
            Left            =   4800
            TabIndex        =   147
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Stop"
            Height          =   375
            Index           =   1
            Left            =   4800
            TabIndex        =   146
            Top             =   1260
            Width           =   855
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Config"
            Height          =   375
            Index           =   2
            Left            =   4800
            TabIndex        =   145
            Top             =   1800
            Width           =   855
         End
         Begin VB.ComboBox cboPlugins 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   2520
            Width           =   4455
         End
         Begin VB.CheckBox chkPlugins 
            Caption         =   "Use Winamp Visualization"
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblCaptionOpt7 
            AutoSize        =   -1  'True
            Caption         =   "Plugins (Vis plugins for winamp)"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   150
            Top             =   0
            Width           =   2205
         End
         Begin VB.Label lblCaptionOpt7 
            AutoSize        =   -1  'True
            Caption         =   "Winamp Plugins SubType :"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   149
            Top             =   2280
            Width           =   1920
         End
         Begin VB.Shape shpWinamp 
            BorderColor     =   &H80000010&
            Height          =   2895
            Index           =   0
            Left            =   0
            Top             =   120
            Width           =   5775
         End
      End
      Begin VB.PictureBox picWinamp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3015
         Index           =   1
         Left            =   420
         ScaleHeight     =   3015
         ScaleWidth      =   5820
         TabIndex        =   151
         Top             =   480
         Width           =   5820
         Begin VB.CommandButton cmdDSP 
            Caption         =   "Load"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   156
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton cmdDSP 
            Caption         =   "Unload"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   155
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton cmdDSP 
            Caption         =   "Configure"
            Height          =   375
            Index           =   2
            Left            =   4320
            TabIndex        =   154
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox chkDSP 
            Caption         =   "Enabled  Winamp DSP Plugins"
            Height          =   255
            Left            =   240
            TabIndex        =   153
            Top             =   360
            Width           =   2535
         End
         Begin ComctlLib.ListView lvwDSP 
            Height          =   1695
            Left            =   240
            TabIndex        =   152
            Top             =   720
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   2990
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Plugin"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Load"
               Object.Width           =   1676
            EndProperty
         End
         Begin VB.Label lblCaptionOpt8 
            AutoSize        =   -1  'True
            Caption         =   "Winamp DSP"
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   240
            TabIndex        =   157
            Top             =   0
            Width           =   960
         End
         Begin VB.Shape shpWinamp 
            BorderColor     =   &H80000010&
            Height          =   2895
            Index           =   1
            Left            =   0
            Top             =   120
            Width           =   5775
         End
      End
      Begin ComctlLib.TabStrip tbsWinamp 
         Height          =   3420
         Left            =   360
         TabIndex        =   162
         Top             =   120
         Width           =   5920
         _ExtentX        =   10451
         _ExtentY        =   6033
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Vis Plugins"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "DSP Plugins"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   9
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   44
      Top             =   750
      Width           =   6540
      Begin VB.TextBox txtCredit 
         Alignment       =   2  'Center
         ForeColor       =   &H00DD5500&
         Height          =   3375
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   163
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label lblCaptionOpt10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MP3_ProPlayer"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   47
         Top             =   0
         Width           =   1095
      End
      Begin VB.Shape shpMP3_proPlayer 
         BorderColor     =   &H80000010&
         Height          =   3615
         Left            =   120
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   1
      Left            =   170
      ScaleHeight     =   3855
      ScaleWidth      =   6540
      TabIndex        =   7
      Top             =   750
      Width           =   6540
      Begin ComctlLib.Slider sldPlayer 
         Height          =   255
         Left            =   480
         TabIndex        =   132
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Max             =   25
      End
      Begin VB.TextBox txtDevice 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   81
         Text            =   "3"
         Top             =   2505
         Width           =   495
      End
      Begin VB.TextBox txtPlayer 
         Height          =   375
         Left            =   5040
         TabIndex        =   65
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3960
         TabIndex        =   31
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtDevice 
         Height          =   405
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CheckBox chkDevice 
         Caption         =   "Auto file name"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   28
         Top             =   3300
         Width           =   1335
      End
      Begin VB.TextBox txtDevice 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   24
         Text            =   "4"
         Top             =   2505
         Width           =   495
      End
      Begin VB.CheckBox chkDevice 
         Caption         =   "Lock Video window aspect ratio :"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   23
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ComboBox cboDevice 
         Height          =   315
         Index           =   2
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1980
         Width           =   1695
      End
      Begin VB.ComboBox cboDevice 
         Height          =   315
         Index           =   1
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   1215
      End
      Begin VB.ComboBox cboDevice 
         Height          =   315
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time jum "
         Height          =   195
         Index           =   11
         Left            =   4245
         TabIndex        =   66
         Top             =   1230
         Width           =   675
      End
      Begin VB.Label lblCaptionOpt1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crossfade (0 - 25) :"
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   64
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblCaptionOpt1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Video"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   405
      End
      Begin VB.Shape shpDevice 
         BorderColor     =   &H80000010&
         Height          =   1095
         Index           =   1
         Left            =   120
         Top             =   1792
         Width           =   6255
      End
      Begin VB.Label lblCaptionOpt1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note : This option will not enabled until you restart file."
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   49
         Top             =   630
         Width           =   3795
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directory"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   30
         Top             =   3360
         Width           =   630
      End
      Begin VB.Label lblCaptionOpt1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Wave write"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   2960
         Width           =   810
      End
      Begin VB.Shape shpDevice 
         BorderColor     =   &H80000010&
         Height          =   680
         Index           =   2
         Left            =   120
         Top             =   3060
         Width           =   6255
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(usually 4:3)"
         Height          =   195
         Index           =   7
         Left            =   4920
         TabIndex        =   26
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Index           =   6
         Left            =   4200
         TabIndex        =   25
         Top             =   2520
         Width           =   45
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Screen"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   22
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label lblCaptionOpt1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Output Device"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   0
         Width           =   1035
      End
      Begin VB.Shape shpDevice 
         BorderColor     =   &H80000010&
         Height          =   1500
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample rate"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   18
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblCaptionOpt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[App]
Private tTempMainVis As MainVis
Private tTempSkinVis As SkinVis
Private tTempAppConfig As AppGeneral
Private tTempPlaylistConfig As PlaylistConfig
Private tTempPlayerConfig As PlayerConfig
Private tTempDevice As Device
Private tTempCurrentSkin As M3PSkin

Private fso As New FileSystemObject

Public intTabIndex As Integer
Public bolShow As Boolean


Dim i As Long

Private Sub cboData_Click()
    tMainWin.Data = cboData.ListIndex
End Sub

Private Sub cboDevice_Click(Index As Integer)
    Select Case Index
        Case 0
            tDevice.SoundD.OutputDevice = cboDevice(0).ItemData(cboDevice(0).ListIndex)
            If cboDevice(0).ListIndex = cboDevice(0).ListCount - 1 Then
                tDevice.SoundD.WaveWrite = True
            Else
                tDevice.SoundD.WaveWrite = False
                frmMedia.Player.SetDevice (tDevice.SoundD.OutputDevice)
            End If
        Case 1
            tDevice.SoundD.Freq = cboDevice(1).List(cboDevice(1).ListIndex)
        Case 2
            tDevice.VideoD.intDefaultScreen = cboDevice(2).ListIndex
    End Select
End Sub



Private Sub cboLang_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cboPlugins_Click()
    tWinamp.intSubPlugin = cboPlugins.ListIndex
End Sub


Private Sub chkDevice_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Select Case Index
            Case 0
                tDevice.VideoD.bolLockRatio = CBool(chkDevice(0).value)
                txtDevice(0).Enabled = tDevice.VideoD.bolLockRatio
                txtDevice(1).Enabled = tDevice.VideoD.bolLockRatio
            Case 1
                tDevice.WaveD.bolAutoFilename = CBool(chkDevice(1).value)
        End Select
        cmdApply.Enabled = True
    End If
End Sub

Private Sub chkDSP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tWinampDSP.bolEnabled = CBool(chkDSP.value)
    End If
    If tWinampDSP.bolEnabled Then
        cmdDSP(0).Enabled = True
        cmdDSP(1).Enabled = True
        cmdDSP(2).Enabled = True
        Call StartDSP
    Else
        Call StopDSP
        cmdDSP(0).Enabled = False
        cmdDSP(1).Enabled = False
        cmdDSP(2).Enabled = False
    End If
End Sub

Private Sub chkGeneral_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Dim reg As New clsRegistry
        Select Case Index
            Case 0 'Auto start
                tAppConfig.bolAutoStart = CBool(chkGeneral(0).value)
                reg.ClassKey = HKEY_CURRENT_USER
                reg.SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                reg.ValueKey = "MP3_proPlayer"
                reg.ValueType = REG_SZ
                reg.value = Chr(34) & GetProperPath(App.path) & "MP3_proPlayer.exe" & Chr(34)
                If tAppConfig.bolAutoStart Then
                    If reg.KeyExists = False Then reg.CreateKey
                Else
                    If reg.KeyExists Then reg.DeleteValue
                End If
                MsgBox "You need restart to apply this option", vbInformation And vbOKOnly, "M3P _ Message"
            Case 1
                tAppConfig.bolShowSplash = CBool(chkGeneral(1).value)
            Case 2
                tAppConfig.bolMenu = CBool(chkGeneral(4).value)
                reg.ClassKey = HKEY_CLASSES_ROOT
                If tAppConfig.bolMenu Then
                    'With directory
                        reg.SectionKey = "Directory\shell\Add to M3P"
                        reg.ValueKey = ""
                        reg.ValueType = REG_SZ
                        reg.value = "&Add to M3P"
                        reg.CreateKey
                        reg.SectionKey = "Directory\shell\Add to M3P\command"
                        reg.ValueKey = ""
                        reg.ValueType = REG_SZ
                        reg.value = Chr(34) & GetProperPath(App.path) & "MP3_proPlayer.exe" & Chr(34) & " " & "/add %1"
                        reg.CreateKey
                    'With folder
                        reg.SectionKey = "Folder\shell\Add to M3P"
                        reg.ValueKey = ""
                        reg.ValueType = REG_SZ
                        reg.value = "&Add to M3P"
                        reg.CreateKey
                        reg.SectionKey = "Folder\shell\Add to M3P\command"
                        reg.ValueKey = ""
                        reg.ValueType = REG_SZ
                        reg.value = Chr(34) & GetProperPath(App.path) & "MP3_proPlayer.exe" & Chr(34) & " " & "/add %1"
                        reg.CreateKey
                    'With drive
                        reg.SectionKey = "Drive\shell\Add to M3P"
                        reg.ValueKey = ""
                        reg.ValueType = REG_SZ
                        reg.value = "&Add to M3P"
                        reg.CreateKey
                        reg.SectionKey = "Drive\shell\Add to M3P\command"
                        reg.ValueKey = ""
                        reg.ValueType = REG_SZ
                        reg.value = Chr(34) & GetProperPath(App.path) & "MP3_proPlayer.exe" & Chr(34) & " " & "/add %1"
                        reg.CreateKey
                Else
                    reg.ClassKey = HKEY_CLASSES_ROOT
                    reg.SectionKey = "Directory\shell\Add to M3P\command"
                    reg.DeleteKey
                    reg.SectionKey = "Directory\shell\Add to M3P"
                    reg.DeleteKey
                    
                    reg.SectionKey = "Folder\shell\Add to M3P\command"
                    reg.DeleteKey
                    reg.SectionKey = "Folder\shell\Add to M3P"
                    reg.DeleteKey
                    
                    reg.SectionKey = "Drive\shell\Add to M3P\command"
                    reg.DeleteKey
                    reg.SectionKey = "Drive\shell\Add to M3P"
                    reg.DeleteKey
                    
                End If
            Case 3
                tAppConfig.bolTaskbar = CBool(chkGeneral(3).value)
                frmMenu.Visible = tAppConfig.bolTaskbar
            Case 4
                tAppConfig.bolTaskbarScroll = CBool(chkGeneral(4).value)
            Case 5
                tAppConfig.bolSysTray = CBool(chkGeneral(5).value)
                If tAppConfig.bolSysTray Then
                    sysTray.AddIcon frmMedia.hwnd, frmMedia.Icon.handle
                Else
                    sysTray.RemoveIcon frmMedia.hwnd
                End If
            Case 6
                tPlayerConfig.bolAutoPlay = CBool(chkGeneral(6).value)
            Case 7
                tPlayerConfig.bolShowList = CBool(chkGeneral(7).value)
            Case 8
                tPlayerConfig.bolAutoShutdow = CBool(chkGeneral(8).value)
            Case 9
                tPlayerConfig.bolAutoRemove = CBool(chkGeneral(9).value)
            Case 10
                tPlayerConfig.bolAutoExit = CBool(chkGeneral(10).value)
        End Select
        cmdApply.Enabled = True
    End If
End Sub



Private Sub chkLibrary_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        LibOption.bolUse = CBool(chkLibrary.value)
        If LibOption.bolUse Then
            Call LoadDatabase
            frmMenu.mnuLibrary.Enabled = True
            cmdLibrary(0).Enabled = True
            cmdLibrary(1).Enabled = True
            cmdLibrary(2).Enabled = True
        Else
            ReDim Library(0)
            cmdLibrary(0).Enabled = False
            cmdLibrary(1).Enabled = False
            cmdLibrary(2).Enabled = False
            frmMenu.mnuLibrary.Enabled = False
        End If
        cmdApply.Enabled = True
    End If
End Sub

Private Sub chkMainVis_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        Select Case Index
            Case 0
                tMainWin.bolShowTitle = CBool(chkMainVis(0).value)
                If frmMenu.mnuVisualC(0).Checked Then frmVisual.lblTitle.Visible = tMainWin.bolShowTitle
            Case 1
                tMainWin.bolUsePic = CBool(chkMainVis(1).value)
                If tMainWin.bolUsePic = False Then
                    frmVisual.picVis.Picture = Nothing
                    frmVisual.picVis.BackColor = tMainWin.BackColor
                Else
                    If FileExists(tMainWin.BackGround) Then
                        frmVisual.picVis.Cls
                        frmVisual.picVis.PaintPicture frmVisual.picBG, 0, 0, frmVisual.picVis.ScaleWidth, frmVisual.picVis.ScaleHeight, 0, 0
                        frmVisual.picVis.Picture = frmVisual.picVis.Image
                    End If
                End If
        End Select
        cmdApply.Enabled = True
    End If
End Sub

Private Sub chkPlaylist_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As Integer
    tPlaylistConfig.bolShowNumber = CBool(chkPlaylist.value)
    frmPlayList.List.Number = tPlaylistConfig.bolShowNumber
    cmdApply.Enabled = True
End Sub

Private Sub chkPlugins_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tWinamp.bolEnabled = CBool(chkPlugins.value)
    End If
    cmdApply.Enabled = True
End Sub

Private Sub chkSkin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tSkinOption.bolEQSlide = CBool(chkSkin.value)
    End If
    cmdApply.Enabled = True
End Sub


Private Sub chkSpec_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmMedia.Vis.SpecShowPeak = CBool(chkSpec.value)
    tSkinVis.bolSpecPeak = CBool(chkSpec.value)
End Sub

Private Sub cmdApply_Click()
    If tTempPlaylistConfig.strDisplay <> tPlaylistConfig.strDisplay Then
        Dim x As Long
        Dim strText As String
        Dim y As Long
        If frmPlayList.List.ListItemCount > 0 Then
            For x = 1 To frmPlayList.List.ListItemCount
                y = frmPlayList.List.Key(x)
                strText = ""
                strText = tPlaylistConfig.strDisplay
                strText = Replace(strText, "%1", NowPlaying(y).Infor.Artist)
                strText = Replace(strText, "%2", NowPlaying(y).Infor.Title)
                strText = Replace(strText, "%3", NowPlaying(y).Infor.Album)
                strText = Replace(strText, "%4", NowPlaying(y).Infor.Genre)
                strText = Replace(strText, "%5", NowPlaying(y).Infor.Year)
                strText = Replace(strText, "%6", NowPlaying(y).Infor.Filename)
                strText = Replace(strText, "%7", NowPlaying(y).Infor.FullName)
                NowPlaying(y).strText = strText
            Next x
            For x = 1 To frmPlayList.List.ListItemCount
                y = frmPlayList.List.Key(x)
                frmPlayList.List.ListItemText(x) = NowPlaying(y).strText
            Next x
            frmPlayList.List.Number = tPlaylistConfig.bolShowNumber
        End If
    End If
    LoadLang (cboLang.List(cboLang.ListIndex))
    tTempAppConfig = tAppConfig
    tTempPlayerConfig = tPlayerConfig
    tTempPlayerConfig = tPlayerConfig
    tTempPlaylistConfig = tPlaylistConfig
    tTempMainVis = tMainWin
    tTempDevice = tDevice
    Call SaveConfig
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    tAppConfig = tTempAppConfig
    tPlayerConfig = tTempPlayerConfig
    tPlayerConfig = tTempPlayerConfig
    tPlaylistConfig = tTempPlaylistConfig
    tMainWin = tTempMainVis
    tDevice = tTempDevice
    tSkinVis = tTempSkinVis
    Call SaveConfig
    Me.Hide
End Sub

Private Sub cmdDevice_Click()
    On Error GoTo beep
        Dim Browse As Shell
        Dim strFolder As String
        Set Browse = New Shell
        strFolder = Browse.BrowseForFolder(Me.hwnd, "Select a Folder", 0).Items.Item.path
        txtDevice(2).text = strFolder
        tDevice.WaveD.strWaveOutput = strFolder
beep:
    Set Browse = Nothing
End Sub

Private Sub cmdDSP_Click(Index As Integer)
    Select Case Index
        Case 0
            Call StartDSP
        Case 1
            Call StopDSP
        Case 2
            If Len(lvwDSP.SelectedItem.Key) > 0 Then
                DspOpenConfig lvwDSP.SelectedItem.Key
            End If
    End Select
End Sub


Private Sub cmdGenPlugins_Click(Index As Integer)
    If Index = 0 Then
        genPlugin(lstGenPlugins.ListIndex).config
    Else
        genPlugin(lstGenPlugins.ListIndex).about
    End If
End Sub

Private Sub cmdLibrary_Click(Index As Integer)
    On Error GoTo beep
    Select Case Index
        Case 0
            Dim Browse As Shell
            Dim strFolder As String
            Set Browse = New Shell
            strFolder = Browse.BrowseForFolder(Me.hwnd, "Select a Folder", 0).Items.Item.path
            lstLibrary.AddItem strFolder
            Set Browse = Nothing
        Case 1
            For i = 0 To lstLibrary.ListCount - 1
                If lstLibrary.Selected(i) = True Then
                    lstLibrary.RemoveItem i
                End If
            Next i
        Case 2
            For i = 0 To lstLibrary.ListCount - 1
                If lstLibrary.Selected(i) = True Then
                    SearchMedia (lstLibrary.List(i))
                End If
            Next i
    End Select
    cmdApply.Enabled = True
beep:
    Set Browse = Nothing
End Sub

Private Sub cmdMainVis_Click()
    On Error GoTo beep
    Dim strFileName As String
    With cdloColor
        .DialogTitle = "Choice picture"
        .DefaultExt = "*.bmp"
        .Filter = "All Supported Files |*.bmp;*.jpg;*.gif"
        .flags = cdlOFNFileMustExist
        .CancelError = True
        .ShowOpen
        strFileName = .Filename
    End With
    If strFileName <> "" Then
        tMainWin.BackGround = strFileName
        txtMainVis.text = strFileName
        frmVisual.picBG.Picture = LoadPicture(strFileName)
        If tMainWin.bolUsePic Then
            frmVisual.picVis.Cls
            frmVisual.picVis.PaintPicture frmVisual.picBG, 0, 0, frmVisual.picVis.ScaleWidth, frmVisual.picVis.ScaleHeight, 0, 0
            frmVisual.picVis.Picture = frmVisual.picVis.Image
        End If
    End If
beep:
    If Err.Number <> 0 Then
        Exit Sub
    End If
End Sub



Private Sub cmdOk_Click()
    If cmdApply.Enabled = True Then
    If tTempPlaylistConfig.strDisplay <> tPlaylistConfig.strDisplay Then
        Dim x As Long
        Dim strText As String
        Dim y As Long
            If frmPlayList.List.ListItemCount > 0 Then
                For x = 1 To frmPlayList.List.ListItemCount
                    y = frmPlayList.List.Key(x)
                    strText = ""
                    strText = Replace(strText, "%1", NowPlaying(y).Infor.Artist)
                    strText = Replace(strText, "%2", NowPlaying(y).Infor.Title)
                    strText = Replace(strText, "%3", NowPlaying(y).Infor.Album)
                    strText = Replace(strText, "%4", NowPlaying(y).Infor.Genre)
                    strText = Replace(strText, "%5", NowPlaying(y).Infor.Year)
                    strText = Replace(strText, "%6", NowPlaying(y).Infor.Filename)
                    strText = Replace(strText, "%7", NowPlaying(y).Infor.FullName)
                    NowPlaying(y).strText = strText
                Next x
                
                For x = 1 To frmPlayList.List.ListItemCount
                    y = frmPlayList.List.Key(x)
                    frmPlayList.List.ListItemText(x) = NowPlaying(y).strText
                Next x
                
                frmPlayList.List.Number = tPlaylistConfig.bolShowNumber
            End If
        End If
        LoadLang (cboLang.List(cboLang.ListIndex))
        Call SaveConfig
    End If
    Me.Hide
End Sub

Private Sub cmdPlayer_Click(Index As Integer)
    Dim reg As New clsRegistry
    Dim regBak As New clsRegistry
    Dim strOldAss As String
    If Index <> 4 Then
        For i = 1 To lvwPlayer.ListItems.Count
            lvwPlayer.ListItems(i).Selected = False
        Next i
    End If
    Select Case Index
        Case 0
            For i = 1 To lvwPlayer.ListItems.Count
                lvwPlayer.ListItems(i).Selected = True
            Next i
            cmdPlayer(4).Enabled = True
        Case 1
            tPlayerConfig.strFileType = ""
                For i = 1 To lvwPlayer.ListItems.Count
                    regBak.ClassKey = HKEY_CLASSES_ROOT
                    regBak.SectionKey = "." & lvwPlayer.ListItems(i).text
                    regBak.ValueKey = "MP3_proPlayer_Bak"
                    regBak.ValueType = REG_SZ
                    strOldAss = regBak.value
                    
                    regBak.ValueKey = ""
                    regBak.ValueType = REG_SZ
                    regBak.value = strOldAss
                    regBak.CreateKey
                    
                    regBak.ClassKey = HKEY_LOCAL_MACHINE
                    regBak.SectionKey = "SOFTWARE\Classes\" & "." & lvwPlayer.ListItems(i).text
                    regBak.ValueKey = "MP3_proPlayer_Bak"
                    regBak.ValueType = REG_SZ
                    strOldAss = regBak.value
                    
                    regBak.ValueKey = ""
                    regBak.ValueType = REG_SZ
                    regBak.value = strOldAss
                    regBak.CreateKey
                Next i
                For i = 1 To lvwPlayer.ListItems.Count
                    reg.ClassKey = HKEY_CLASSES_ROOT
                    reg.SectionKey = "." & lvwPlayer.ListItems(i).text
                    reg.ValueType = REG_SZ
                    reg.SectionKey = reg.value
                    lvwPlayer.ListItems(i).SubItems(1) = reg.value
                    
                    reg.SectionKey = reg.SectionKey & "\shell\open\command"
                    reg.ValueKey = ""
                    lvwPlayer.ListItems(i).SubItems(2) = Mid(reg.value, 2, InStr(1, reg.value, ".exe", vbBinaryCompare) + 2)
                Next i
                cmdPlayer(4).Enabled = False
        Case 2
                lvwPlayer.ListItems(4).Selected = True
                lvwPlayer.ListItems(7).Selected = True
                lvwPlayer.ListItems(11).Selected = True
                lvwPlayer.ListItems(15).Selected = True
                lvwPlayer.ListItems(16).Selected = True
                lvwPlayer.ListItems(17).Selected = True
                cmdPlayer(4).Enabled = True
        Case 3
                lvwPlayer.ListItems(1).Selected = True
                lvwPlayer.ListItems(2).Selected = True
                lvwPlayer.ListItems(3).Selected = True
                lvwPlayer.ListItems(5).Selected = True
                lvwPlayer.ListItems(6).Selected = True
                lvwPlayer.ListItems(8).Selected = True
                lvwPlayer.ListItems(9).Selected = True
                lvwPlayer.ListItems(10).Selected = True
                lvwPlayer.ListItems(12).Selected = True
                lvwPlayer.ListItems(13).Selected = True
                lvwPlayer.ListItems(14).Selected = True
                lvwPlayer.ListItems(18).Selected = True
                cmdPlayer(4).Enabled = True
        Case 4
                For i = 1 To lvwPlayer.ListItems.Count
                    
                    regBak.ClassKey = HKEY_CLASSES_ROOT
                    regBak.SectionKey = "." & lvwPlayer.ListItems(i).text
                    regBak.ValueKey = ""
                    regBak.ValueType = REG_SZ
                    If regBak.value <> "M3P.File" And regBak.value <> "M3P.Playlist" Then
                        strOldAss = regBak.value
                        regBak.ValueKey = "MP3_proPlayer_Bak"
                        regBak.ValueType = REG_SZ
                        regBak.value = strOldAss
                        regBak.CreateKey
                    End If
                    
                    regBak.ClassKey = HKEY_LOCAL_MACHINE
                    regBak.SectionKey = "SOFTWARE\Classes\" & "." & lvwPlayer.ListItems(i).text
                    regBak.ValueKey = ""
                    regBak.ValueType = REG_SZ
                    If regBak.value <> "M3P.File" And regBak.value <> "M3P.Playlist" Then
                        strOldAss = regBak.value
                        regBak.ValueKey = "MP3_proPlayer_Bak"
                        regBak.ValueType = REG_SZ
                        regBak.value = strOldAss
                        regBak.CreateKey
                    End If
                    If lvwPlayer.ListItems(i).Selected Then
                            If i < 19 Then
                                Call reg.CreateEXEAssociation(GetProperPath(App.path) & "MP3_proPlayer.exe", "M3P.File", "M3P Media File", lvwPlayer.ListItems(i).text, "&Play in M3P", "/open", False, , , , , , tAppConfig.intFileIcon)
                            Else
                                Call reg.CreateEXEAssociation(GetProperPath(App.path) & "MP3_proPlayer.exe", "M3P.Playlist", "M3P Playlist File", lvwPlayer.ListItems(i).text, "&Play in M3P", "/open", False, , , , , , tAppConfig.intPLIcon)
                            End If
                    End If
                Next i
                For i = 1 To lvwPlayer.ListItems.Count
                    reg.ClassKey = HKEY_CLASSES_ROOT
                    reg.SectionKey = "." & lvwPlayer.ListItems(i).text
                    reg.ValueType = REG_SZ
                    reg.SectionKey = reg.value
                    lvwPlayer.ListItems(i).SubItems(1) = reg.value
                    
                    reg.SectionKey = reg.SectionKey & "\shell\open\command"
                    reg.ValueKey = ""
                    lvwPlayer.ListItems(i).SubItems(2) = Mid(reg.value, 2, InStr(1, reg.value, ".exe", vbBinaryCompare) + 2)
                Next i
                cmdPlayer(4).Enabled = False
    End Select
    tPlayerConfig.strFileType = ""
    For i = 1 To lvwPlayer.ListItems.Count
        If lvwPlayer.ListItems(i).Selected Then
            If Len(lvwPlayer.ListItems(i).text) = 3 Then
                tPlayerConfig.strFileType = tPlayerConfig.strFileType & "." & lvwPlayer.ListItems(i).text & ","
            Else
                tPlayerConfig.strFileType = tPlayerConfig.strFileType & lvwPlayer.ListItems(i).text & ","
            End If
        End If
    Next i
    cmdApply.Enabled = True
End Sub

Private Sub cmdPlaylist_Click()
    txtPlaylist(0).text = "%1 - %2"
End Sub

Private Sub cmdPlugins_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            Call Stop_VisPlg
            Call Start_VisPlg
        Case 1
            Call Stop_VisPlg
        Case 2
            Dim PluginsIndex As Long
            PluginsIndex = lstPlugins.ListIndex
            Call BASS_WA_Config_Vis(PluginsIndex, cboPlugins.ListIndex)
    End Select
    cmdApply.Enabled = True
End Sub

Private Sub cmdSkin_Click(Index As Integer)
    Select Case Index
        Case 0
            If lstSkin.List(lstSkin.ListIndex) <> tCurrentSkin.Infor.Name Then
                Call LoadSkin(lstSkin.List(lstSkin.ListIndex) & ".skn", tCurrentSkin.mini)
                If Not FileExists(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Readme.txt") Then
                    txtSkin.text = tCurrentSkin.Infor.Author & vbCrLf & tCurrentSkin.Infor.Comment & vbCrLf & tCurrentSkin.Infor.Location
                Else
                    Dim strFileRead As String
                    Dim strText As String
                    strFileRead = tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Readme.txt"
                    Open strFileRead For Input As #1
                        Do Until EOF(1)
                            Input #1, strText
                            txtSkin.text = strText & vbCrLf
                        Loop
                    Close #1
                End If
                cmdSkin(3).Caption = tCurrentSkin.Infor.Author
            End If
        Case 1
            Dim strNewName As String
            strNewName = InputBox("Enter a new name", "Change Skin name")
            If strNewName <> "" Then
                Name tSkinOption.SkinDir & "\" & lstSkin.List(lstSkin.ListIndex) & ".skn" As tSkinOption.SkinDir & "\" & strNewName & ".skn"
                lstSkin.List(lstSkin.ListIndex) = strNewName
            End If
        Case 2
            If MsgBox("Are you sure want to delete " & lstSkin.List(lstSkin.ListIndex) & " ?", vbQuestion, "MP3_proPlayer") Then
                Kill tSkinOption.SkinDir & "\" & lstSkin.List(lstSkin.ListIndex) & ".skn"
                lstSkin.RemoveItem lstSkin.ListIndex
                Call frmMenu.Add_SkinInstall(tSkinOption.SkinDir)
            End If
        Case 3
            ShellExecute Me.hwnd, vbNullString, tCurrentSkin.Infor.Location, vbNullString, vbNullString, SW_SHOWNORMAL
    End Select
    cmdApply.Enabled = True
End Sub

Private Sub cmdVisualization_Click(Index As Integer)
    Select Case Index
        Case 0
            If lstAVS.ListIndex >= 0 Then
                tMainWin.plugin = lstAVS.List(lstAVS.ListIndex)
                If Not frmVisual.bolShow Then frmVisual.Show
            End If
        Case 1
            If frmVisual.bolShow Then Unload frmVisual
        Case 2
            '[Visualization]
            WriteINI "Visualization", "GetStyle", tMainWin.Style, strFileconfig
            WriteINI "Visualization", "Data", tMainWin.Data, strFileconfig
            WriteINI "Visualization", "Plugin", tMainWin.plugin, strFileconfig
            WriteINI "Visualization", "FontColor", tMainWin.FontColor, strFileconfig
            WriteINI "Visualization", "BackColor", tMainWin.BackColor, strFileconfig
            WriteINI "Visualization", "BackGround", tMainWin.BackGround, strFileconfig
            WriteINI "Visualization", "Interval", tMainWin.TimeDisplay, strFileconfig
            WriteINI "Visualization", "UsePictureBG", tMainWin.bolUsePic, strFileconfig
            WriteINI "Visualization", "ShowTitle", tMainWin.bolShowTitle, strFileconfig
    End Select
    cmdApply.Enabled = True
End Sub



Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = LoadResPicture(112, vbResIcon)
    bolShow = True
    For i = 1 To 12
        imlIcon.ListImages.Add i, , LoadResPicture(100 + i, vbResIcon)
    Next i
    tTempAppConfig = tAppConfig
    tTempPlayerConfig = tPlayerConfig
    tTempPlayerConfig = tPlayerConfig
    tTempPlaylistConfig = tPlaylistConfig
    tTempMainVis = tMainWin
    tTempDevice = tDevice
    
    Call SetNumber(txtDevice(0), True)
    Call SetNumber(txtDevice(1), True)
    Call SetNumber(txtPlayer, True)
    Call SetNumber(txtPlaylist(2), True)
    Call SetNumber(txtLibrary(0), True)
    Call SetNumber(txtLibrary(1), True)
    Call Init
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolShow = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    WriteINI "Library", "Total", lstLibrary.ListCount, strFileconfig
    For i = 0 To lstLibrary.ListCount - 1
        WriteINI "Library", "Folder" & i, lstLibrary.List(i), strFileconfig
    Next i
    If frmMenu.mnuVisualC(0).Checked Then frmVisual.Enabled = True
    For i = 0 To UBound(genPlugin) - 1
        If ObjPtr(genPlugin(i)) > 0 Then
            Set genPlugin(i) = Nothing
        End If
    Next i

End Sub




Private Sub lblMainVis_Click()
    On Error Resume Next
    Dim lngColor As Long
    
    With cdloColor
        .CancelError = True
        .DialogTitle = "MP3_proPlayer : Choice Color"
        .flags = &H1&
        .ShowColor
        lngColor = .Color
    End With
    lblMainVis.BackColor = lngColor
    tMainWin.BackColor = lngColor
    If tMainWin.bolUsePic = False Then frmVisual.picVis.BackColor = tMainWin.BackColor
    cmdApply.Enabled = True
End Sub


Private Sub lblVisTab_Click(Index As Integer)
    picVis(Index).ZOrder 0
    For i = 0 To lblVisTab.Count - 1
        lblVisTab(i).BackColor = &HFFFFFF
        lblVisTab(i).ForeColor = &H0
    Next i
    lblVisTab(Index).BackColor = &H0
    lblVisTab(Index).ForeColor = &HFFFFFF
End Sub

Private Sub lblVisTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For i = 0 To lblVisTab.Count - 1
        lblVisTab(i).BackColor = &HFFFFFF
        lblVisTab(i).ForeColor = &H0
    Next i
    lblVisTab(Index).BackColor = &H0
    lblVisTab(Index).ForeColor = &HFFFFFF
End Sub


Private Sub lstPlugins_Click()
    Dim Index As Long
    Dim ModuleInfo As String
    Dim lpModuleInfo As Long
    Dim cntModule As Long
    Dim NumOfModules As Long
    
        cboPlugins.Clear
        Index = lstPlugins.ListIndex
    
        NumOfModules = BASS_WA_GetModuleCount(Index)
    
        For cntModule = 0 To NumOfModules - 1
            lpModuleInfo = BASS_WA_GetModuleInfo(Index, cntModule)
            ModuleInfo = GetStringFromPointer(lpModuleInfo)
            cboPlugins.AddItem ModuleInfo
        Next cntModule
            
        cboPlugins.ListIndex = 0
        tWinamp.intCurrentPlugin = lstPlugins.ListIndex
        tWinamp.intSubPlugin = cboPlugins.ListIndex
End Sub


Private Sub lvwDSP_Click()
    tWinampDSP.intCurrentPlugin = lvwDSP.SelectedItem.Index
End Sub




Private Sub optOscStyle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    frmMedia.Vis.OscDrawMode = Index
    tSkinVis.intOsc = Index
End Sub

Private Sub optPlaylist_Click(Index As Integer)
    tPlaylistConfig.intLoadID = Index
    cmdApply.Enabled = True
End Sub


Private Sub optSkinVis_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For i = 0 To 2
        frmMenu.mnuMainSpecC(i).Checked = False
    Next i
    tSkinVis.intStyle = Index
    If Index = 2 Then frmMedia.Vis.doStop
    frmMedia.Vis.StyleVis = Index
    frmMenu.mnuMainSpecC(Index).Checked = True
End Sub


Private Sub optSpecDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    frmMedia.Vis.SpecDrawMode = Index
    tSkinVis.intSpecDraw = Index
End Sub

Private Sub optSpecStyle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    frmMedia.Vis.SpecFillMode = Index
    tSkinVis.intSpecFill = Index
End Sub

Private Sub optStyle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    tMainWin.Style = Index
End Sub



Private Sub sldGeneral_Scroll()
    tAppConfig.intIcon = sldGeneral.value + 1
    imgGeneral.Picture = frmMedia.imgIcon.ListImages(tAppConfig.intIcon).Picture
    frmMedia.Icon = imgGeneral.Picture
    If tAppConfig.bolSysTray Then
        sysTray.RemoveIcon frmMedia.hwnd
        sysTray.AddIcon frmMedia.hwnd, frmMedia.Icon.handle
    End If
End Sub
Private Sub sldIcon_Scroll(Index As Integer)
    Select Case Index
        Case 0
            tAppConfig.intFileIcon = sldIcon(0).value
            imgIcon(0).Picture = imlIcon.ListImages(tAppConfig.intFileIcon).Picture
        Case 1
            tAppConfig.intPLIcon = sldIcon(1).value
            imgIcon(1).Picture = imlIcon.ListImages(tAppConfig.intPLIcon).Picture
    End Select
    cmdApply.Enabled = True
End Sub


Private Sub sldMainVis_Change()
    If frmVisual.bolShow Then
        tMainWin.TimeDisplay = CInt(1000 \ sldMainVis.value)
        frmVisual.tmrVisual.Interval = tMainWin.TimeDisplay
    End If
    cmdApply.Enabled = True
End Sub

Private Sub sldPlayer_Change()
    lblCaptionOpt1(10).Caption = "Crossfade (0 - 25) : " & sldPlayer.value
    If sldPlayer.value = 0 Then lblCaptionOpt1(10).Caption = "Crossfade (0 - 25) : Disabled"
End Sub

Private Sub sldPlayer_Scroll()
    tPlayerConfig.intCrossfade = sldPlayer.value
    If tPlayerConfig.intCrossfade = 0 Then
        tPlayerConfig.bolCrossfade = False
    Else
        tPlayerConfig.bolCrossfade = True
    End If
End Sub

Private Sub sldSkinVis_Change(Index As Integer)
    Select Case Index
        Case 0
            frmMedia.tmrVisual.Interval = CInt(1000 \ sldSkinVis(0).value)
            tSkinVis.intRefresh = frmMedia.tmrVisual.Interval
        Case 1
            frmMedia.Vis.SpecPeakDelay = sldSkinVis(1).value
            tSkinVis.intSpecPeakPause = sldSkinVis(1).value
        Case 2
            frmMedia.Vis.SpecPeakDrop = sldSkinVis(2).value
            tSkinVis.intSpecPeakDrop = sldSkinVis(2).value
    End Select
End Sub
Private Sub tbsOption_Click()
    picOption(tbsOption.SelectedItem.Index - 1).ZOrder 0
    intTabIndex = tbsOption.SelectedItem.Index
End Sub


Private Sub tbsWinamp_Click()
    picWinamp(tbsWinamp.SelectedItem.Index - 1).ZOrder 0
End Sub


Private Sub txtDevice_Change(Index As Integer)
    Select Case Index
        Case 1
            tDevice.VideoD.intRatioHeight = txtDevice(1).text
        Case 0
            tDevice.VideoD.intRatioWidth = txtDevice(0).text
    End Select
    cmdApply.Enabled = True
End Sub



Private Sub txtLibrary_Change(Index As Integer)
    Select Case Index
        Case 0
            LibOption.lngAudioSkip = txtLibrary(0).text
        Case 1
            LibOption.lngVideoSkip = txtLibrary(1).text
    End Select
    cmdApply.Enabled = True
End Sub

Private Sub txtMainVis_Change()
    On Error Resume Next
    If Not FileExists(txtMainVis.text) Then
        Set frmVisual.picBG.Picture = Nothing
        Set frmVisual.picVis.Picture = Nothing
        tMainWin.BackGround = ""
    Else
        tMainWin.BackGround = txtMainVis.text
    End If
End Sub

Private Sub txtPlayer_Change()
    tPlayerConfig.intTime = CInt(txtPlayer.text)
    cmdApply.Enabled = True
End Sub

Private Sub txtPlaylist_Change(Index As Integer)
    Select Case Index
        Case 0
            tPlaylistConfig.strDisplay = txtPlaylist(0).text
        Case 2
            If txtPlaylist(2).text < 25 Then
                tPlaylistConfig.intRowS = txtPlaylist(2).text
            End If
    End Select
    cmdApply.Enabled = True
End Sub
Public Sub Init()
    On Error Resume Next
        
        
    '[tab General]
        If tAppConfig.bolAutoStart Then chkGeneral(0).value = Checked
        If tAppConfig.bolShowSplash Then chkGeneral(1).value = Checked
        If tAppConfig.bolMenu Then chkGeneral(2).value = Checked
        If tAppConfig.bolTaskbar Then chkGeneral(3).value = Checked
        If tAppConfig.bolTaskbarScroll Then chkGeneral(4).value = Checked
        If tAppConfig.bolSysTray Then chkGeneral(5).value = Checked
        
        If tPlayerConfig.bolAutoPlay Then chkGeneral(6).value = Checked
        If tPlayerConfig.bolShowList Then chkGeneral(7).value = Checked
        If tPlayerConfig.bolAutoShutdow Then chkGeneral(8).value = Checked
        If tPlayerConfig.bolAutoRemove Then chkGeneral(9).value = Checked
        If tPlayerConfig.bolAutoExit Then chkGeneral(10).value = Checked
        
        InitLang
        For i = 0 To cboLang.ListCount - 1
            If cboLang.List(i) = CurrentLang Then
                cboLang.ListIndex = i
                Exit For
            End If
        Next i
        sldGeneral.value = tAppConfig.intIcon - 1
        imgGeneral.Picture = frmMedia.imgIcon.ListImages(tAppConfig.intIcon).Picture
        
    '[End tabGeneral]

    '[tab Device]
        Dim strDevice As String
        Dim intDevice As Integer
        intDevice = 1
        While BASS_GetDeviceDescription(intDevice)
            strDevice = VBStrFromAnsiPtr(BASS_GetDeviceDescription(intDevice))
            cboDevice(0).AddItem strDevice, intDevice - 1
            cboDevice(0).ItemData(cboDevice(0).NewIndex) = intDevice
            intDevice = intDevice + 1
        Wend
        cboDevice(0).AddItem "Wave write", intDevice - 1
        cboDevice(0).ItemData(cboDevice(0).NewIndex) = intDevice
        cboDevice(0).ListIndex = tDevice.SoundD.OutputDevice - 1
        
        cboDevice(1).AddItem "11025", 0
        cboDevice(1).AddItem "22050", 1
        cboDevice(1).AddItem "41100", 2
        cboDevice(1).AddItem "48000", 3
        cboDevice(1).AddItem "96000", 4
        For i = 0 To cboDevice(1).ListCount - 1
            If cboDevice(1).List(i) = CStr(tDevice.SoundD.Freq) Then
                cboDevice(1).ListIndex = i
                Exit For
            End If
        Next
        cboDevice(2).AddItem "Half screen", 0
        cboDevice(2).AddItem "Default screen", 1
        cboDevice(2).AddItem "Double screen", 2
        cboDevice(2).AddItem "Maximize screen", 3
        cboDevice(2).ListIndex = tDevice.VideoD.intDefaultScreen
        If tDevice.VideoD.bolLockRatio Then
            chkDevice(0).value = Checked
            txtDevice(0).Enabled = True
            txtDevice(1).Enabled = True
        End If
        txtDevice(0).text = tDevice.VideoD.intRatioWidth
        txtDevice(1).text = tDevice.VideoD.intRatioHeight
        txtDevice(2).text = tDevice.WaveD.strWaveOutput
        If tDevice.WaveD.bolAutoFilename Then chkDevice(1).value = Checked
    '[End tabDevice]
    
    '[TabPlayer]
        Dim atype() As String
        Dim J As Integer
        lvwPlayer.ListItems.Add 1, , "AIF"
        lvwPlayer.ListItems.Add 2, , "ASF"
        lvwPlayer.ListItems.Add 3, , "AVI"
        lvwPlayer.ListItems.Add 4, , "CDA"
        lvwPlayer.ListItems.Add 5, , "DAT"
        lvwPlayer.ListItems.Add 6, , "M1V"
        lvwPlayer.ListItems.Add 7, , "MID"
        lvwPlayer.ListItems.Add 8, , "MOV"
        lvwPlayer.ListItems.Add 9, , "MP1"
        lvwPlayer.ListItems.Add 10, , "MP2"
        lvwPlayer.ListItems.Add 11, , "MP3"
        lvwPlayer.ListItems.Add 12, , "MPE"
        lvwPlayer.ListItems.Add 13, , "MPG"
        lvwPlayer.ListItems.Add 14, , "MPEG"
        lvwPlayer.ListItems.Add 15, , "OGG"
        lvwPlayer.ListItems.Add 16, , "WAV"
        lvwPlayer.ListItems.Add 17, , "WMA"
        lvwPlayer.ListItems.Add 18, , "WMV"
        lvwPlayer.ListItems.Add 19, , "M3U"
        lvwPlayer.ListItems.Add 20, , "PLS"
        atype = Split(tPlayerConfig.strFileType, ",")
        For i = 1 To lvwPlayer.ListItems.Count
            For J = 0 To UBound(atype)
                If Len(lvwPlayer.ListItems(i).text) = 3 Then
                    If atype(J) = "." & lvwPlayer.ListItems(i).text Then lvwPlayer.ListItems(i).Selected = True
                End If
                If Len(lvwPlayer.ListItems(i).text) = 4 Then
                    If atype(J) = lvwPlayer.ListItems(i).text Then lvwPlayer.ListItems(i).Selected = True
                End If
            Next J
        Next i
         
        Dim reg As New clsRegistry
        Dim tmpSec As String
        
        For i = 1 To lvwPlayer.ListItems.Count
            tmpSec = "." & lvwPlayer.ListItems(i).text
            reg.ClassKey = HKEY_CLASSES_ROOT
            reg.SectionKey = tmpSec
            reg.ValueType = REG_SZ
            reg.SectionKey = reg.value
            lvwPlayer.ListItems(i).SubItems(1) = reg.value
            
            reg.SectionKey = reg.SectionKey & "\shell\open\command"
            reg.ValueKey = ""
            lvwPlayer.ListItems(i).SubItems(2) = Mid(reg.value, 2, InStr(1, reg.value, ".exe", vbBinaryCompare) + 2)
        Next i
                
        sldIcon(0).value = tAppConfig.intFileIcon
        sldIcon(1).value = tAppConfig.intPLIcon
        imgIcon(0).Picture = imlIcon.ListImages(tAppConfig.intFileIcon).Picture
        imgIcon(1).Picture = imlIcon.ListImages(tAppConfig.intPLIcon).Picture
        sldPlayer.value = tPlayerConfig.intCrossfade
        txtPlayer.text = tPlayerConfig.intTime
    '[End tabPlayer]
    
    '[tab Playlist]
        optPlaylist(tPlaylistConfig.intLoadID).value = True
        If tPlaylistConfig.bolShowNumber Then chkPlaylist.value = Checked
        txtPlaylist(0).text = tPlaylistConfig.strDisplay
        txtPlaylist(2).text = tPlaylistConfig.intRowS
    '[End tabplaylist]
    
    '[Tab Library]
        Dim lngFile As Long
        
        If LibOption.bolUse Then
            chkLibrary.value = Checked
            cmdLibrary(0).Enabled = True
            cmdLibrary(1).Enabled = True
            cmdLibrary(2).Enabled = True
        Else
            chkLibrary.value = Unchecked
            cmdLibrary(0).Enabled = False
            cmdLibrary(1).Enabled = False
            cmdLibrary(2).Enabled = False
        End If
        lngFile = ReadINI("Library", "Total", strFileconfig, 0)
        If lngFile > 0 Then
            For i = 0 To lngFile - 1
                If FileExists(ReadINI("Library", "Folder" & i, strFileconfig)) Then
                    frmOption.lstLibrary.AddItem ReadINI("Library", "Folder" & i, strFileconfig)
                End If
            Next i
        End If
        txtLibrary(0).text = LibOption.lngAudioSkip
        txtLibrary(1).text = LibOption.lngVideoSkip
    '[End TabLibrary]
    
    '[tab Skin]
        Dim strFileRead As String
        Dim strText As String
        For i = 3 To frmMenu.mnuSkinC.Count - 1
            lstSkin.AddItem frmMenu.mnuSkinC(i).Caption
        Next i
        For i = 0 To lstSkin.ListCount - 1
            If lstSkin.List(i) = tCurrentSkin.Infor.Name Then lstSkin.Selected(i) = True
        Next i
        If Not FileExists(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Readme.txt") Then
            txtSkin.text = tCurrentSkin.Infor.Author & vbCrLf & tCurrentSkin.Infor.Comment & vbCrLf & tCurrentSkin.Infor.Location
        Else
            strFileRead = tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Readme.txt"
            Open strFileRead For Input As #1
                Do Until EOF(1)
                    Input #1, strText
                    txtSkin.text = txtSkin.text & strText & vbCrLf
                Loop
            Close #1
            strText = ""
        End If
        cmdSkin(3).Caption = tCurrentSkin.Infor.Author
        If tSkinOption.bolEQSlide Then chkSkin.value = Checked
    '[End tabSkin]
    
    
    '[Tab Visualization]
        sldSkinVis(0).value = 1000 / tSkinVis.intRefresh
        sldSkinVis(1).value = tSkinVis.intSpecPeakPause
        sldSkinVis(2).value = tSkinVis.intSpecPeakDrop
        If tSkinVis.bolSpecPeak Then chkSpec.value = vbChecked
        optSkinVis(tSkinVis.intStyle).value = True
        optSpecDraw(tSkinVis.intSpecDraw).value = True
        optSpecStyle(tSkinVis.intSpecFill).value = True
        optOscStyle(tSkinVis.intOsc).value = True
        
        Call InitVis
        sldMainVis.value = tMainWin.TimeDisplay
        If tMainWin.bolShowTitle Then chkMainVis(0).value = vbChecked
        If tMainWin.bolUsePic Then chkMainVis(1).value = vbChecked
        txtMainVis.text = tMainWin.BackGround
        lblMainVis.BackColor = tMainWin.BackColor
        cboData.ListIndex = tMainWin.Data
        optStyle(tMainWin.Style).value = True
    '[End tabVisualization]
        
    '[Tab Plugins]
        Call InitGenPlugin
    '[Tab Winamp Plugins]
        Call InitWinampVis
        If tWinamp.bolEnabled Then chkPlugins.value = 1
        Call SetupDSPPlugins
        If tWinampDSP.bolEnabled Then
            chkDSP.value = 1
            cmdDSP(0).Enabled = True
            cmdDSP(1).Enabled = True
            cmdDSP(2).Enabled = True
        Else
            chkDSP.value = 0
            cmdDSP(0).Enabled = False
            cmdDSP(1).Enabled = False
            cmdDSP(2).Enabled = False
        End If
        If tWinampDSP.intCurrentPlugin <> 0 Then
            lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Selected = True
        End If
    '[About]
    Dim tmp As String
    strFileRead = App.path & "\Readme.txt"
    Open strFileRead For Input As #1
        Do Until EOF(1)
            Input #1, tmp
            strText = strText & vbCrLf & tmp
        Loop
    Close #1
    txtCredit.text = strText
End Sub
Public Sub CallTab(Optional Index As Integer)
    If Index = vbNull Then Index = intTabIndex
    If Index > 0 And Index <= 10 Then
        tbsOption.Tabs(Index).Selected = True
        picOption(tbsOption.SelectedItem.Index - 1).ZOrder 0
    End If
End Sub
