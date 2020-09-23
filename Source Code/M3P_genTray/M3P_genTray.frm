VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTray 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M3P_genTray"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "M3P_genTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgTray 
      Left            =   2760
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M3P_genTray.frx":23D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M3P_genTray.frx":2626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M3P_genTray.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M3P_genTray.frx":2ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M3P_genTray.frx":2D22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Stop"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Pause"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Next"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Play"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Previous"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1560
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   840
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP           As Long = &H202


Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                        (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                        (ByVal HWND As Long, ByVal wMsg As Long, _
                                        ByVal wParam As Long, lParam As Any) As Long
Private Type COPYDATASTRUCT
            dwData As Long
            cbData As Long
            lpData As Long
End Type
Private Const GWL_WNDPROC = (-4)
Private Const WM_COPYDATA = &H4A


Private Sub SendData(ByVal sData As String)
    Dim cdCopyData As COPYDATASTRUCT
    Dim ThWnd As Long
    Dim bytBuffer(1 To 255) As Byte
      
    ThWnd = FindWindow(vbNullString, "MP3_ProPlayer 2.1.0")
    
    CopyMemory bytBuffer(1), ByVal sData, Len(sData)
    cdCopyData.dwData = 3
    cdCopyData.cbData = Len(sData) + 1
    cdCopyData.lpData = VarPtr(bytBuffer(1))
    Call SendMessage(ThWnd, WM_COPYDATA, frmTray.HWND, cdCopyData)

End Sub

Private Sub chkTray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Select Case Index
            Case 0
                tTray.bolPrevious = CBool(chkTray(0).Value)
                If tTray.bolPrevious Then
                    Call AddIcon(picTray(0).HWND, imgTray.ListImages(2).ExtractIcon.Handle)
                    Call SysTip(picTray(0).HWND, "Previous")
                Else
                    Call RemoveIcon(picTray(0).HWND)
                End If
            Case 1
                tTray.bolPlay = CBool(chkTray(1).Value)
                If tTray.bolPlay Then
                    Call AddIcon(picTray(1).HWND, imgTray.ListImages(2).ExtractIcon.Handle)
                    Call SysTip(picTray(1).HWND, "Play")
                Else
                    Call RemoveIcon(picTray(1).HWND)
                End If
            Case 2
                tTray.bolNext = CBool(chkTray(2).Value)
                If tTray.bolNext Then
                    Call AddIcon(picTray(2).HWND, imgTray.ListImages(3).ExtractIcon.Handle)
                    Call SysTip(picTray(2).HWND, "Next")
                Else
                    Call RemoveIcon(picTray(2).HWND)
                End If
            Case 3
                tTray.bolPause = CBool(chkTray(3).Value)
                If tTray.bolPause Then
                    Call AddIcon(picTray(3).HWND, imgTray.ListImages(4).ExtractIcon.Handle)
                    Call SysTip(picTray(3).HWND, "Pause")
                Else
                    Call RemoveIcon(picTray(3).HWND)
                End If
            Case 4
                tTray.bolStop = CBool(chkTray(4).Value)
                If tTray.bolStop Then
                    Call AddIcon(picTray(4).HWND, imgTray.ListImages(5).ExtractIcon.Handle)
                    Call SysTip(picTray(4).HWND, "Stop")
                Else
                    Call RemoveIcon(picTray(4).HWND)
                End If
        End Select
    End If
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    If tTray.bolPrevious Then
        chkTray(0).Value = vbChecked
    End If
    If tTray.bolPlay Then
        chkTray(1).Value = vbChecked
    End If
    If tTray.bolNext Then
        chkTray(2).Value = vbChecked
    End If
    If tTray.bolPause Then
        chkTray(3).Value = vbChecked
    End If
    If tTray.bolStop Then
        chkTray(4).Value = vbChecked
    End If
End Sub

Private Sub Form_Terminate()
    RemoveIcon (picTray(0).HWND)
    RemoveIcon (picTray(1).HWND)
    RemoveIcon (picTray(2).HWND)
    RemoveIcon (picTray(3).HWND)
    RemoveIcon (picTray(4).HWND)
End Sub

Private Sub picTray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    Err.Clear
    Static bolBusy As Boolean
        If bolBusy = False Then
            bolBusy = True
            Select Case CLng(X)
                Case WM_LBUTTONUP
                    Select Case Index
                        Case 0  'Previous
                            Call SendData("/back")
                        Case 1  'Play
                            Call SendData("/play")
                        Case 2  'Next
                            Call SendData("/next")
                        Case 3  'Pause
                            Call SendData("/pause")
                        Case 4  'Stop
                            Call SendData("/stop")
                    End Select
            End Select
            bolBusy = False
        End If
End Sub
