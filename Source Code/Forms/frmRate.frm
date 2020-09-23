VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P : Pitch"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ComctlLib.Slider sldRate 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   327682
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   2
      Left            =   3405
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000"
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   270
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngRate As Long
Public bolShow As Boolean
Private Sub cmdReset_Click()
    sldRate.value = lngRate
End Sub

Private Sub Form_Load()
    bolShow = True
    sldRate.MousePointer = vbCustom
    If Not bolVideoOn Then
        lngRate = frmMedia.Player.Rate
        sldRate.max = 100000
        sldRate.min = 100
        sldRate.TickFrequency = 1000
    Else
        lngRate = frmVD.Video.Rate * 100
        sldRate.max = 200
        sldRate.min = 10
        sldRate.TickFrequency = 10
    End If
    sldRate.value = lngRate
    lblRate(2).Caption = "[ " & lngRate & " ]"
    lblRate(0).Caption = sldRate.min
    lblRate(1).Caption = sldRate.max
End Sub

Private Sub Form_Terminate()
    bolShow = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bolShow = False
End Sub

Private Sub sldRate_Change()
    If Not bolVideoOn Then
        frmMedia.Player.Rate = sldRate.value
    Else
        frmVD.Video.Rate = sldRate.value / 100
    End If
    lblRate(2).Caption = "[ " & sldRate.value & " ]"
End Sub
