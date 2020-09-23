VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{21897D67-63B7-480F-BB8D-CE51D4D25E82}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmDSP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P : DSP Studio"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7425
   Icon            =   "frmDSP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   0
      Left            =   180
      TabIndex        =   20
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgChorus 
         Height          =   1575
         Index           =   0
         Left            =   480
         TabIndex        =   155
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Value           =   5
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   65280
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":000C
      End
      Begin VB.ComboBox cboChorus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmDSP.frx":0028
         Left            =   4800
         List            =   "frmDSP.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cboChorus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmDSP.frx":0070
         Left            =   4800
         List            =   "frmDSP.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin M3P_Control.Progress prgChorus 
         Height          =   1575
         Index           =   1
         Left            =   1320
         TabIndex        =   156
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Value           =   5
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   65280
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0096
      End
      Begin M3P_Control.Progress prgChorus 
         Height          =   1575
         Index           =   2
         Left            =   2160
         TabIndex        =   157
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   99
         Min             =   -99
         Value           =   5
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   65280
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":00B2
      End
      Begin M3P_Control.Progress prgChorus 
         Height          =   1575
         Index           =   3
         Left            =   3105
         TabIndex        =   158
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   10
         Value           =   5
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   65280
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":00CE
      End
      Begin M3P_Control.Progress prgChorus 
         Height          =   1575
         Index           =   4
         Left            =   3900
         TabIndex        =   159
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   20
         Value           =   5
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   65280
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":00EA
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Index           =   16
         Left            =   3780
         TabIndex        =   64
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         Height          =   195
         Index           =   15
         Left            =   3690
         TabIndex        =   63
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Index           =   14
         Left            =   2980
         TabIndex        =   62
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   195
         Index           =   13
         Left            =   2880
         TabIndex        =   61
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         Height          =   195
         Index           =   12
         Left            =   1926
         TabIndex        =   60
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-99"
         Height          =   195
         Index           =   11
         Left            =   1920
         TabIndex        =   59
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Index           =   10
         Left            =   1200
         TabIndex        =   58
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   195
         Index           =   9
         Left            =   1038
         TabIndex        =   57
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Index           =   8
         Left            =   330
         TabIndex        =   56
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   55
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lPhase"
         Height          =   195
         Index           =   6
         Left            =   5535
         TabIndex        =   30
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waveform"
         Height          =   195
         Index           =   5
         Left            =   5400
         TabIndex        =   29
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delay"
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   28
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency"
         Height          =   195
         Index           =   3
         Left            =   2865
         TabIndex        =   27
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Feedback"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depth"
         Height          =   195
         Index           =   1
         Left            =   1245
         TabIndex        =   25
         Top             =   600
         Width           =   450
      End
      Begin VB.Label lblChorus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WetDryMix"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   7
      Left            =   180
      TabIndex        =   0
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgReverb 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   193
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Max             =   0
         Min             =   -96
         Value           =   -48
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0106
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enable"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin M3P_Control.Progress prgReverb 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   194
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Max             =   0
         Min             =   -96
         Value           =   -48
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0122
      End
      Begin M3P_Control.Progress prgReverb 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   195
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Max             =   3000
         Min             =   1
         Value           =   1500
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":013E
      End
      Begin M3P_Control.Progress prgReverb 
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   196
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Max             =   999
         Min             =   1
         Value           =   999
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":015A
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HFR Time Ratio"
         Height          =   195
         Index           =   11
         Left            =   4800
         TabIndex        =   13
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reverb Mix"
         Height          =   195
         Index           =   10
         Left            =   1200
         TabIndex        =   12
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reverb Time"
         Height          =   195
         Index           =   9
         Left            =   4920
         TabIndex        =   11
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input Gain"
         Height          =   195
         Index           =   8
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.001"
         Height          =   195
         Index           =   7
         Left            =   3840
         TabIndex        =   9
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.999"
         Height          =   195
         Index           =   6
         Left            =   6375
         TabIndex        =   8
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3000 ms"
         Height          =   195
         Index           =   5
         Left            =   6375
         TabIndex        =   7
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1 ms"
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 db"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   5
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-96 db"
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   4
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 db"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblReverb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-96 db"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   600
         Width           =   450
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   6
      Left            =   180
      TabIndex        =   19
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   181
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   0
         Min             =   -10000
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0176
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   182
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   0
         Min             =   -10000
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0192
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   183
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   10
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":01AE
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   184
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   20000
         Min             =   100
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":01CA
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   185
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   2000
         Min             =   100
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":01E6
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   186
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   1000
         Min             =   -10000
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0202
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   187
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   300
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":021E
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   188
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   2000
         Min             =   -10000
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":023A
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   189
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0256
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   190
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0272
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   191
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":028E
      End
      Begin M3P_Control.Progress prgLReverb 
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   192
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Max             =   20000
         Min             =   20
         Value           =   -5000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":02AA
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   35
         Left            =   6720
         TabIndex        =   152
         Top             =   1980
         Width           =   180
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.02"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   4920
         TabIndex        =   151
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   6720
         TabIndex        =   150
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   32
         Left            =   5040
         TabIndex        =   149
         Top             =   1500
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   6720
         TabIndex        =   148
         Top             =   1020
         Width           =   270
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   5040
         TabIndex        =   147
         Top             =   1020
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   6720
         TabIndex        =   146
         Top             =   540
         Width           =   225
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   5040
         TabIndex        =   145
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2000"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   4320
         TabIndex        =   144
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-10000"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   2400
         TabIndex        =   143
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   4400
         TabIndex        =   142
         Top             =   1500
         Width           =   225
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   2640
         TabIndex        =   141
         Top             =   1500
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   4320
         TabIndex        =   140
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-10000"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   2400
         TabIndex        =   139
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   4400
         TabIndex        =   138
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   2640
         TabIndex        =   137
         Top             =   540
         Width           =   225
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   2040
         TabIndex        =   136
         Top             =   1980
         Width           =   180
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   135
         Top             =   1980
         Width           =   225
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   2040
         TabIndex        =   134
         Top             =   1500
         Width           =   180
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   133
         Top             =   1500
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   2040
         TabIndex        =   132
         Top             =   1020
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-10000"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   131
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   2040
         TabIndex        =   130
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-10000"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   129
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HFReference (Khz)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   5325
         TabIndex        =   128
         Top             =   1980
         Width           =   1365
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Density (%)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   5580
         TabIndex        =   127
         Top             =   1500
         Width           =   780
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Diffusion (%)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   5535
         TabIndex        =   126
         Top             =   1020
         Width           =   870
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ReverbDelay (s)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   5400
         TabIndex        =   125
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reverb (mB)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   3203
         TabIndex        =   124
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ReflectionsDelay (s)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   2940
         TabIndex        =   123
         Top             =   1500
         Width           =   1410
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reflections (mB)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3068
         TabIndex        =   122
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DecayHFRatio"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   121
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DecayTime (s)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   120
         Top             =   1980
         Width           =   1020
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RoomRolloffFactor"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   119
         Top             =   1500
         Width           =   1320
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RoomHF (mB)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   118
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label lblLRevrb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Room (mB)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   117
         Top             =   540
         Width           =   780
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   5
      Left            =   180
      TabIndex        =   18
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgGargle 
         Height          =   255
         Left            =   240
         TabIndex        =   180
         Top             =   1800
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         Max             =   1000
         Min             =   1
         Value           =   500
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":02C6
      End
      Begin VB.ComboBox cboGargle 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmDSP.frx":02E2
         Left            =   360
         List            =   "frmDSP.frx":02EC
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblGargle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1000 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   6360
         TabIndex        =   116
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label lblGargle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   115
         Top             =   1605
         Width           =   330
      End
      Begin VB.Label lblGargle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WaveShape :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   114
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblGargle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RateHz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   113
         Top             =   1600
         Width           =   540
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   4
      Left            =   180
      TabIndex        =   17
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgFlanger 
         Height          =   1575
         Index           =   0
         Left            =   480
         TabIndex        =   165
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0308
      End
      Begin VB.ComboBox cboFlanger 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "frmDSP.frx":0324
         Left            =   5520
         List            =   "frmDSP.frx":0334
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cboFlanger 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmDSP.frx":036C
         Left            =   5520
         List            =   "frmDSP.frx":0376
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin M3P_Control.Progress prgFlanger 
         Height          =   1575
         Index           =   1
         Left            =   1560
         TabIndex        =   166
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0392
      End
      Begin M3P_Control.Progress prgFlanger 
         Height          =   1575
         Index           =   2
         Left            =   2565
         TabIndex        =   167
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   99
         Min             =   -99
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":03AE
      End
      Begin M3P_Control.Progress prgFlanger 
         Height          =   1575
         Index           =   3
         Left            =   3600
         TabIndex        =   168
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   10
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":03CA
      End
      Begin M3P_Control.Progress prgFlanger 
         Height          =   1575
         Index           =   4
         Left            =   4680
         TabIndex        =   169
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   4
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   255
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":03E6
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   4560
         TabIndex        =   111
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   4560
         TabIndex        =   110
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   3450
         TabIndex        =   109
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   3360
         TabIndex        =   108
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-99"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   2280
         TabIndex        =   107
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   2325
         TabIndex        =   106
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   1380
         TabIndex        =   105
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   1200
         TabIndex        =   104
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   300
         TabIndex        =   103
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   102
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Phase :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   101
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Waveform :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   5520
         TabIndex        =   100
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Delay"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   4600
         TabIndex        =   99
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   98
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Feedback"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   97
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Depth (LFO)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   96
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lblFlanger 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WetDryMix"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   3
      Left            =   180
      TabIndex        =   16
      Top             =   800
      Width           =   7095
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin M3P_Control.Progress prgEcho 
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   176
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0402
      End
      Begin VB.CheckBox chkEcho 
         Caption         =   "Swap (Left - Right)"
         Height          =   315
         Left            =   2760
         TabIndex        =   80
         Top             =   720
         Width           =   1695
      End
      Begin M3P_Control.Progress prgEcho 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   177
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":041E
      End
      Begin M3P_Control.Progress prgEcho 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   178
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Max             =   2000
         Min             =   1
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":043A
      End
      Begin M3P_Control.Progress prgEcho 
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   179
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Max             =   2000
         Min             =   1
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0456
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2000 ms"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   6360
         TabIndex        =   92
         Top             =   1850
         Width           =   600
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 ms"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   4440
         TabIndex        =   91
         Top             =   1850
         Width           =   330
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2000 ms"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   6360
         TabIndex        =   90
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 ms"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   4440
         TabIndex        =   89
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   88
         Top             =   1845
         Width           =   270
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   87
         Top             =   1850
         Width           =   90
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2505
         TabIndex        =   86
         Top             =   1140
         Width           =   270
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   85
         Top             =   1140
         Width           =   90
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RightDelay"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   5280
         TabIndex        =   84
         Top             =   1850
         Width           =   780
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LeftDelay"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   5280
         TabIndex        =   83
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Feedback"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   82
         Top             =   1850
         Width           =   720
      End
      Begin VB.Label lblEcho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WetDryMix"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   81
         Top             =   1140
         Width           =   780
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgDistortion 
         Height          =   1575
         Index           =   0
         Left            =   480
         TabIndex        =   160
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   0
         Min             =   -60
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0472
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin M3P_Control.Progress prgDistortion 
         Height          =   1575
         Index           =   1
         Left            =   1920
         TabIndex        =   161
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":048E
      End
      Begin M3P_Control.Progress prgDistortion 
         Height          =   1575
         Index           =   2
         Left            =   3360
         TabIndex        =   162
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   8000
         Min             =   100
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":04AA
      End
      Begin M3P_Control.Progress prgDistortion 
         Height          =   1575
         Index           =   3
         Left            =   4920
         TabIndex        =   163
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   8000
         Min             =   100
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":04C6
      End
      Begin M3P_Control.Progress prgDistortion 
         Height          =   1575
         Index           =   4
         Left            =   6480
         TabIndex        =   164
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2778
         Max             =   8000
         Min             =   100
         Value           =   0
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         Orientation     =   1
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":04E2
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   5940
         TabIndex        =   79
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8000 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   5850
         TabIndex        =   78
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   4380
         TabIndex        =   77
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8000 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   4290
         TabIndex        =   76
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   2850
         TabIndex        =   75
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8000 Hz"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   2760
         TabIndex        =   74
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1635
         TabIndex        =   73
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1455
         TabIndex        =   72
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-60 db"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   71
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label lblDistortion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 db"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   70
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lblDistortion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PreLowpassCutoff"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   5760
         TabIndex        =   69
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label lblDistortion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EQBandwidth"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   68
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDistortion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EQCenterFrequency"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   67
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label lblDistortion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Edge"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   66
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblDistortion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gain"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   65
         Top             =   600
         Width           =   330
      End
   End
   Begin VB.Frame frmDSP 
      Height          =   2655
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   800
      Width           =   7095
      Begin M3P_Control.Progress prgCompressor 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   170
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Max             =   0
         Min             =   -60
         Value           =   30
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":04FE
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin M3P_Control.Progress prgCompressor 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   171
         Top             =   1500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Max             =   50000
         Min             =   1
         Value           =   10000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":051A
      End
      Begin M3P_Control.Progress prgCompressor 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   172
         Top             =   2280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Max             =   3000
         Min             =   50
         Value           =   10000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0536
      End
      Begin M3P_Control.Progress prgCompressor 
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   173
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Max             =   0
         Min             =   -60
         Value           =   10000
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":0552
      End
      Begin M3P_Control.Progress prgCompressor 
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   174
         Top             =   1500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Min             =   1
         Value           =   10
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":056E
      End
      Begin M3P_Control.Progress prgCompressor 
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   175
         Top             =   2280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Max             =   4
         Value           =   2
         BackColor       =   14737632
         BorderColor     =   -2147483646
         ForeColor       =   12632256
         MinColor        =   9786150
         ProgressStyle   =   1
         MouseIcon       =   "frmDSP.frx":058A
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4 ms"
         Height          =   195
         Index           =   17
         Left            =   6240
         TabIndex        =   54
         Top             =   2100
         Width           =   330
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 ms"
         Height          =   195
         Index           =   16
         Left            =   3840
         TabIndex        =   53
         Top             =   2100
         Width           =   330
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   195
         Index           =   15
         Left            =   6240
         TabIndex        =   52
         Top             =   1320
         Width           =   270
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   195
         Index           =   14
         Left            =   3960
         TabIndex        =   51
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 db"
         Height          =   195
         Index           =   13
         Left            =   6240
         TabIndex        =   50
         Top             =   540
         Width           =   315
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-60 db"
         Height          =   195
         Index           =   12
         Left            =   3840
         TabIndex        =   49
         Top             =   540
         Width           =   450
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3000 ms"
         Height          =   195
         Index           =   11
         Left            =   2520
         TabIndex        =   48
         Top             =   2100
         Width           =   600
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50 ms"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   47
         Top             =   2100
         Width           =   420
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "500 ms"
         Height          =   195
         Index           =   9
         Left            =   2520
         TabIndex        =   46
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.01 ms"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 db"
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   44
         Top             =   540
         Width           =   315
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-60 db"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   540
         Width           =   450
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Predelay"
         Height          =   195
         Index           =   5
         Left            =   5040
         TabIndex        =   42
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ratio"
         Height          =   195
         Index           =   4
         Left            =   5160
         TabIndex        =   41
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Threshold"
         Height          =   195
         Index           =   3
         Left            =   4995
         TabIndex        =   40
         Top             =   540
         Width           =   705
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Release"
         Height          =   195
         Index           =   2
         Left            =   1215
         TabIndex        =   39
         Top             =   2100
         Width           =   585
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attack"
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   38
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblCompressor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output Gain"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   37
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6360
      TabIndex        =   154
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkApplyAll 
      Caption         =   "Disabled All Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   153
      Top             =   120
      Width           =   2055
   End
   Begin ComctlLib.TabStrip tbsDSP 
      Height          =   3015
      Left            =   120
      TabIndex        =   197
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5318
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   8
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Chorus"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Compressor"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Distortion"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Echo"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Flanger"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Gargle"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "I3DL2Reverb"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reverb"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim intTabDSP As Integer
Public bolShow As Boolean


Private Sub cboChorus_Change(Index As Integer)
    Call UpdateFX(0)
End Sub

Private Sub cboFlanger_Change(Index As Integer)
    Call UpdateFX(4)
End Sub

Private Sub cboGargle_Change()
    Call UpdateFX(5)
End Sub

Private Sub chkApplyAll_Click()
    Wait 500
    If chkApplyAll.value = 1 Then
        For i = 0 To 7
            Call BASS_ChannelRemoveFX(frmMedia.Player.handle, FX(i))
            chkDSP(i).Enabled = False
        Next i
        bolUseDirectX = False
    Else
        bolUseDirectX = True
        Call SetFX(frmMedia.Player.handle)
        For i = 0 To 7
            chkDSP(i).Enabled = True
            Call UpdateFX(i)
        Next i
    End If
End Sub

Private Sub chkDSP_Click(Index As Integer)
    Wait 500
    If chkDSP(Index).value = 0 Then
        Call BASS_ChannelRemoveFX(frmMedia.Player.handle, FX(Index))
        bolDSP(Index) = False
    Else
        bolDSP(Index) = True
        Call SetFX(frmMedia.Player.handle)
        Call UpdateFX(CLng(Index))
    End If
End Sub


Private Sub chkEcho_Click()
    Call UpdateFX(3)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = LoadResPicture(112, vbResIcon)
    bolShow = True
    intTabDSP = ReadINI("DSP Effect", "TabIndex", strFileconfig)
    'Call lblDSP_Click(intTabDSP)
    tbsDSP.Tabs(intTabDSP).Selected = True
    frmDSP(tbsDSP.SelectedItem.Index - 1).ZOrder 0
    For i = 0 To chkDSP.Count - 1
        If bolDSP(i) Then
            chkDSP(i).value = Checked
        Else
            chkDSP(i).value = Unchecked
        End If
    Next i
    '
            For i = 0 To 4
                prgChorus(i).value = intChorus(i)
            Next i
                cboChorus(0).ListIndex = intChorus(5)
                cboChorus(1).ListIndex = intChorus(6)
            For i = 0 To 5
                prgCompressor(i).value = intCompressor(i)
            Next i
            For i = 0 To 4
                prgDistortion(i).value = intDistortion(i)
            Next i
            For i = 0 To 3
                prgEcho(i).value = intEcho(i)
            Next i
                chkEcho.value = intEcho(4)
            For i = 0 To 4
                prgFlanger(i).value = intFlanger(i)
            Next i
                cboFlanger(0).ListIndex = intFlanger(5)
                cboFlanger(1).ListIndex = intFlanger(6)
                prgGargle.value = intGargle(0)
                cboGargle.ListIndex = intGargle(1)
            For i = 0 To 11
                prgLReverb(i).value = intI3DL2Reverb(i)
            Next i
            For i = 0 To 3
                prgReverb(i).value = intReverb(i)
            Next i
    If bolUseDirectX Then
        chkApplyAll.value = Unchecked
        Call SetFX(frmMedia.Player)
        For i = 0 To 7
            chkDSP(i).Enabled = True
            Call UpdateFX(i)
        Next i
    Else
        chkApplyAll.value = Checked
        For i = 0 To 7
            Call BASS_ChannelRemoveFX(frmMedia.Player.handle, FX(i))
            chkDSP(i).Enabled = False
        Next i
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    For i = 0 To chkDSP.Count - 1
        bolDSP(i) = CBool(chkDSP(i).value)
    Next i
    '[Save chorus]
        For i = 0 To prgChorus.Count - 1
            intChorus(i) = prgChorus(i).value
        Next i
            intChorus(5) = cboChorus(0).ListIndex
            intChorus(6) = cboChorus(1).ListIndex
     '[Compressor]
        For i = 0 To prgCompressor.Count - 1
            intCompressor(i) = prgCompressor(i).value
        Next i
    '[Distortion]
        For i = 0 To prgDistortion.Count - 1
            intDistortion(i) = prgDistortion(i).value
        Next i
    '[Echo]
        For i = 0 To prgEcho.Count - 1
            intEcho(i) = prgEcho(i).value
        Next i
            intEcho(4) = chkEcho.value
    '[Flanger]
        For i = 0 To prgFlanger.Count - 1
            intFlanger(i) = prgFlanger(i).value
        Next i
            intFlanger(5) = cboFlanger(0).ListIndex
            intFlanger(6) = cboFlanger(1).ListIndex
    '[Gargle]
        intGargle(0) = prgGargle.value
        intGargle(1) = cboGargle.ListIndex
    '[I3DL2Reverb]
        For i = 0 To prgLReverb.Count - 1
            intI3DL2Reverb(i) = prgLReverb(i).value
        Next i
    '[Reverb]
        For i = 0 To prgReverb.Count - 1
            intReverb(i) = prgReverb(i).value
        Next i
        WriteINI "DSP Effect", "UseDirectX", bolUseDirectX, strFileconfig
        WriteINI "DSP Effect", "Chorus", bolDSP(0), strFileconfig
        WriteINI "DSP Effect", "Compressor", bolDSP(1), strFileconfig
        WriteINI "DSP Effect", "Distortion", bolDSP(2), strFileconfig
        WriteINI "DSP Effect", "Echo", bolDSP(3), strFileconfig
        WriteINI "DSP Effect", "Flanger", bolDSP(4), strFileconfig
        WriteINI "DSP Effect", "Gargle", bolDSP(5), strFileconfig
        WriteINI "DSP Effect", "I3DL2Reverb", bolDSP(6), strFileconfig
        WriteINI "DSP Effect", "Reverb", bolDSP(7), strFileconfig
        For i = 0 To 6
            WriteINI "DSP Effect", "Chorus_" & i, intChorus(i), strFileconfig
        Next i
        For i = 0 To 5
            WriteINI "DSP Effect", "Compressor_" & i, intCompressor(i), strFileconfig
        Next i
        For i = 0 To 4
            WriteINI "DSP Effect", "Distortion_" & i, intDistortion(i), strFileconfig
        Next i
        For i = 0 To 4
            WriteINI "DSP Effect", "Echo_" & i, intEcho(i), strFileconfig
        Next i
        For i = 0 To 6
            WriteINI "DSP Effect", "Flanger_" & i, intFlanger(i), strFileconfig
        Next i
        For i = 0 To 1
            WriteINI "DSP Effect", "Gargle_" & i, intGargle(i), strFileconfig
        Next i
        For i = 0 To 11
            WriteINI "DSP Effect", "I3DL2Reverb_" & i, intI3DL2Reverb(i), strFileconfig
        Next i
        For i = 0 To 3
            WriteINI "DSP Effect", "Reverb_" & i, intReverb(i), strFileconfig
        Next i
        WriteINI "DSP Effect", "TabIndex", intTabDSP, strFileconfig
End Sub


Private Sub Form_Unload(Cancel As Integer)
    bolShow = False
    frmMenu.mnuDSP.Checked = False
End Sub
Private Sub prgChorus_Change(Index As Integer, lValue As Long)
    Call UpdateFX(0)
End Sub

Private Sub prgDistortion_Change(Index As Integer, lValue As Long)
    Call UpdateFX(2)
End Sub

Private Sub prgFlanger_Change(Index As Integer, lValue As Long)
    Call UpdateFX(4)
End Sub

Private Sub prgCompressor_Change(Index As Integer, lValue As Long)
    Call UpdateFX(1)
End Sub

Private Sub prgEcho_Change(Index As Integer, lValue As Long)
    Call UpdateFX(3)
End Sub

Private Sub prgGargle_Change(lValue As Long)
    Call UpdateFX(5)
End Sub

Private Sub prgLReverb_Change(Index As Integer, lValue As Long)
    Call UpdateFX(6)
End Sub

Private Sub prgReverb_Change(Index As Integer, lValue As Long)
    Call UpdateFX(7)
End Sub

Private Sub tbsDSP_Click()
    frmDSP(tbsDSP.SelectedItem.Index - 1).ZOrder 0
    intTabDSP = tbsDSP.SelectedItem.Index
End Sub

