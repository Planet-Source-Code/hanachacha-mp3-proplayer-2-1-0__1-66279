VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmInfor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTrans 
      Interval        =   50
      Left            =   240
      Top             =   5160
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Flame"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Form_Load()
    On Error Resume Next
    If FileExists(App.path & "\Image\Splash.bmp") Then
        Me.Picture = LoadPicture(App.path & "\Image\Splash.bmp")
    Else
        Me.BackColor = RGB(0, 0, 0)
    End If
    Call CenterForm(frmSplash)
    Call AlwaysOnTop(Me, True)
    Call WinAPI.Make_TransPercent(Me.hwnd, 0)
    bolLoading = True
    i = 0
    Me.MousePointer = 11
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Load frmMenu
End Sub

Private Sub tmrTrans_Timer()
    i = i + 10
    If i <= 100 Then
        Call WinAPI.Make_TransPercent(Me.hwnd, i)
    Else
        tmrTrans.Enabled = False
        Unload Me
    End If
End Sub
