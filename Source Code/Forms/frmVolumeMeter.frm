VERSION 5.00
Begin VB.Form frmVolumeMeter 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   1245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerVol 
      Interval        =   100
      Left            =   360
      Top             =   4440
   End
   Begin VB.PictureBox pctVolRMask 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3890
      Left            =   710
      ScaleHeight     =   3885
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   10
      Width           =   495
   End
   Begin VB.PictureBox pctVolLMask 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3890
      Left            =   10
      ScaleHeight     =   3885
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   10
      Width           =   495
   End
   Begin VB.PictureBox pctVolL 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3890
      Left            =   10
      ScaleHeight     =   3885
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   10
      Width           =   495
   End
   Begin VB.PictureBox pctVolR 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3890
      Left            =   710
      ScaleHeight     =   3885
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   10
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Left            =   845
      TabIndex        =   5
      Top             =   3960
      Width           =   225
   End
   Begin VB.Label lblVolL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Left            =   167
      TabIndex        =   4
      Top             =   3960
      Width           =   180
   End
End
Attribute VB_Name = "frmVolumeMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    pctVolL.Picture = LoadPicture(App.Path & "Grapphics\nen\VolLFG.jpg")
    pctVolLMask.Picture = LoadPicture(App.Path & "Grapphics\nen\VolLBG.jpg")
    pctVolR.Picture = LoadPicture(App.Path & "Grapphics\nen\VolRFG.jpg")
    pctVolRMask.Picture = LoadPicture(App.Path & "Grapphics\nen\VolRBG.jpg")
End Sub

Private Sub TimerVol_Timer()
Randomize TimerVol
    If frmMedia.allowplay = "yes" And frmMedia.paused = False Then
        pctVolLMask.Height = Val(frmMedia.lblVol.Caption) * Rnd
        pctVolRMask.Height = Val(frmMedia.lblVol.Caption) * Rnd
    End If
End Sub
