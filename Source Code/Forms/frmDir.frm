VERSION 5.00
Object = "{1FD693CE-8CF2-4FB2-99C9-7BFC2F22A0B5}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmBrowse 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M3P _ Open folder"
   ClientHeight    =   5175
   ClientLeft      =   195
   ClientTop       =   435
   ClientWidth     =   4575
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin M3P_Control.ctlExplorer Browser 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      Arrange         =   2
      LabelWrap       =   -1  'True
      MouseIcon       =   "frmDir.frx":000C
      ShowFolders     =   0   'False
      View            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkInsub 
      Caption         =   "Include subfolders"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      Top             =   4680
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H000080FF&
      TabIndex        =   0
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To add all file in a parent folder please check Include subfolders"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   3795
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, x As Integer


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
        cmdCancel.Enabled = False
        If chkInsub.Value = Unchecked Then
            Call AddSingleFolder(Browser.Path)
        Else
            Call AddSubFolder(Browser.Path)
        End If
        Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo beep
    Me.Icon = LoadResPicture(112, vbResIcon)
    AlwaysOnTop Me, True
    
    frmMedia.Enabled = False
    Browser.Path = strLastDir
    cmdCancel.Enabled = True
beep:
    If Err.Number <> 0 Then
        Browser.Path = App.Path
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AlwaysOnTop Me, False
    strLastDir = Browser.Path
    Me.MousePointer = 0
    frmMedia.Enabled = True
End Sub

