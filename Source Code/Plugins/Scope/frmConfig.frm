VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P_Scope"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cboConfig 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfig.frx":0000
         Left            =   1080
         List            =   "frmConfig.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Double"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cboData 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfig.frx":0023
         Left            =   1080
         List            =   "frmConfig.frx":0033
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Size"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Style"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   765
      End
      Begin VB.Label lblOscilliscope 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HightColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblOscilliscope 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MidColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblOscilliscope 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LowColor"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Shape shpOption 
         BackColor       =   &H8000000F&
         BorderColor     =   &H8000000B&
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog cdloColor 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboConfig_Click()
    tOscill.DrawStyle = cboConfig.ListIndex
End Sub


Private Sub cboData_Click()
    tOscill.DataSize = cboData.ListIndex
End Sub

Private Sub chkConfig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        tOscill.Double = CBool(chkConfig.Value)
    End If
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    lblOscilliscope(0).BackColor = tOscill.LowColor
    lblOscilliscope(1).BackColor = tOscill.MidColor
    lblOscilliscope(2).BackColor = tOscill.HightColor
    If tOscill.Double = True Then chkConfig.Value = 1
    cboConfig.ListIndex = tOscill.DrawStyle
    cboData.ListIndex = tOscill.DataSize

End Sub

Private Sub Form_Terminate()
    Dim strINI As String
    strINI = App.Path & "\M3P_vis.ini"
    WriteINI "visOscillscope", "Color1", tOscill.LowColor, strINI
    WriteINI "visOscillscope", "Color2", tOscill.MidColor, strINI
    WriteINI "visOscillscope", "Color3", tOscill.HightColor, strINI
    WriteINI "visOscillscope", "Double", tOscill.Double, strINI
    WriteINI "visOscillscope", "Style", tOscill.DrawStyle, strINI
    WriteINI "visOscillscope", "Data", tOscill.DataSize, strINI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strINI As String
    strINI = App.Path & "\M3P_vis.ini"
    WriteINI "visOscillscope", "Color1", tOscill.LowColor, strINI
    WriteINI "visOscillscope", "Color2", tOscill.MidColor, strINI
    WriteINI "visOscillscope", "Color3", tOscill.HightColor, strINI
    WriteINI "visOscillscope", "Double", tOscill.Double, strINI
    WriteINI "visOscillscope", "Style", tOscill.DrawStyle, strINI
    WriteINI "visOscillscope", "Data", tOscill.DataSize, strINI
End Sub

Private Sub lblOscilliscope_Click(Index As Integer)
    On Error Resume Next
    Dim lngColor As Long
    With cdloColor
        .CancelError = True
        .DialogTitle = "MP3_Scope : Choice Color"
        .Flags = &H1&
        .ShowColor
        lngColor = .Color
    End With
    lblOscilliscope(Index).BackColor = lngColor
    Select Case Index
        Case 0
            tOscill.LowColor = lngColor
        Case 1
            tOscill.MidColor = lngColor
        Case 2
            tOscill.HightColor = lngColor
    End Select
    pic.Cls
    Call FillGradient(pic.hDc, 0, 0, pic.Width, pic.Height / 4, tOscill.MidColor, tOscill.HightColor, Fill_Vertical)
    Call FillGradient(pic.hDc, 0, pic.Height / 4, pic.Width, pic.Height / 4, tOscill.LowColor, tOscill.MidColor, Fill_Vertical)
    Call FillGradient(pic.hDc, 0, pic.Height / 2, pic.Width, pic.Height / 4, tOscill.MidColor, tOscill.LowColor, Fill_Vertical)
    Call FillGradient(pic.hDc, 0, pic.Height - pic.Height / 4, pic.Width, pic.Height / 4, tOscill.HightColor, tOscill.MidColor, Fill_Vertical)
End Sub

Private Sub pic_Resize()
    pic.Cls
    Call FillGradient(pic.hDc, 0, 0, pic.Width, pic.Height / 4, tOscill.MidColor, tOscill.HightColor, Fill_Vertical)
    Call FillGradient(pic.hDc, 0, pic.Height / 4, pic.Width, pic.Height / 4, tOscill.LowColor, tOscill.MidColor, Fill_Vertical)
    Call FillGradient(pic.hDc, 0, pic.Height / 2, pic.Width, pic.Height / 4, tOscill.MidColor, tOscill.LowColor, Fill_Vertical)
    Call FillGradient(pic.hDc, 0, pic.Height - pic.Height / 4, pic.Width, pic.Height / 4, tOscill.HightColor, tOscill.MidColor, Fill_Vertical)
End Sub
