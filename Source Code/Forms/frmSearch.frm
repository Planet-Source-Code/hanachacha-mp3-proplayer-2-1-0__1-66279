VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Media ..."
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame frmSearch 
      Caption         =   "Process"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      Begin ComctlLib.ProgressBar prgStatus 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label lblProgess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   1125
         Width           =   45
      End
      Begin VB.Label lblCurrent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files added :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1400
         Width           =   900
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progess percent :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1120
         Width           =   1245
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current file :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   840
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6000
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AlwaysOnTop Me, True
    frmLibrary.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AlwaysOnTop Me, False
    frmLibrary.Enabled = True
End Sub
