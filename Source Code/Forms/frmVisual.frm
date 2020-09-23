VERSION 5.00
Begin VB.Form frmVisual 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "M3P _ Visualization"
   ClientHeight    =   5490
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   6720
   FillColor       =   &H008080FF&
   Icon            =   "frmVisual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   3
      Top             =   5280
      Width           =   6720
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Visualization Plugins :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   30
         Width           =   1515
      End
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picVis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Timer tmrVisual 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   6840
   End
End
Attribute VB_Name = "frmVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bolShow As Boolean

Dim fft(4096) As Single
Dim Sample(4000) As Integer
Dim oPlugin As Object ' the object to hold the reference to the ActiveX Dll
Dim strVis As String
Dim strClass As String
Dim channel As Long

Private Sub Form_Load()
    On Error GoTo noPlugIn
    
    Me.Icon = LoadResPicture(112, vbResIcon)
    Me.height = ReadINI("Demension", "VisualHeight", strFileconfig, 400) * Screen.TwipsPerPixelY
    Me.width = ReadINI("Demension", "VisualWidth", strFileconfig, 400) * Screen.TwipsPerPixelX
    Me.Left = ReadINI("Demension", "VisualLeft", strFileconfig, 400)
    Me.Top = ReadINI("Demension", "VisualTop", strFileconfig, 400)
    
    bolShow = True
    
    frmMedia.tmrVisual.Enabled = False
    frmMedia.Vis.doStop
    frmMedia.prgVU(0).value = 32768
    frmMedia.prgVU(1).value = 0
    picVis.BackColor = tMainWin.BackColor
    lblTitle.Visible = tMainWin.bolShowTitle
    lblTitle.ForeColor = tMainWin.FontColor
    If FileExists(tMainWin.BackGround) Then
        picBG.Picture = LoadPicture(tMainWin.BackGround)
    End If
    frmMenu.mnuVisualC(0).Checked = True
    
    If FileExists(App.path & "\Plugins\" & tMainWin.plugin) Then
        strVis = Mid(tMainWin.plugin, 1, Len(tMainWin.plugin) - 4)
        strClass = Mid(tMainWin.plugin, 5, Len(tMainWin.plugin) - 8)
        lblStatus.Caption = "Configure Plugins : " & Mid(tMainWin.plugin, 1, Len(tMainWin.plugin) - 4)
        Set oPlugin = CreateObject(strVis & "." & strClass)
        tmrVisual.Interval = tMainWin.TimeDisplay
        tmrVisual.Enabled = True
    End If
    Exit Sub
    
noPlugIn:
    ' plug in could no be created, usually happens if the class was not found
    MsgBox "Could not load Plugin, please register your plugin", vbOKOnly + vbExclamation, "M3P -- PlugIn Error"
    Shell "regsvr32 /s " & App.path & "\Plugins\" & tMainWin.plugin
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteINI "Demension", "VisualHeight", Me.ScaleHeight, strFileconfig
    WriteINI "Demension", "VisualWidth", Me.ScaleWidth, strFileconfig
    WriteINI "Demension", "VisualLeft", Me.Left, strFileconfig
    WriteINI "Demension", "VisualTop", Me.Top, strFileconfig
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleHeight < 117 Then Me.height = 117 * Screen.TwipsPerPixelY
    If Me.ScaleWidth < 274 Then Me.height = 274 * Screen.TwipsPerPixelX
    picVis.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 18
    picStatus.Move 0, Me.ScaleHeight - 18, Me.ScaleWidth, 18
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    tmrVisual.Enabled = False
    bolShow = False
    frmMenu.mnuVisualC(0).Checked = False
    frmMedia.tmrVisual.Enabled = True
    If ObjPtr(oPlugin) > 0 Then
        Set oPlugin = Nothing
    End If
End Sub

Private Sub lblStatus_Click()
    On Error Resume Next
    oPlugin.doConfig ' show the config form
End Sub

Private Sub picVis_Resize()
    On Error Resume Next
    If tMainWin.bolUsePic Then
        If FileExists(tMainWin.BackGround) Then
            picVis.Cls
            picVis.PaintPicture picBG, 0, 0, picVis.ScaleWidth, picVis.ScaleHeight, 0, 0
            picVis.Picture = picVis.Image
        End If
    End If
End Sub

Private Sub tmrVisual_Timer()
    On Error GoTo handle
    Dim lRslt As Long
    
    ' first, clear the picture box
    picVis.Cls
    Select Case tMainWin.Style
        Case 0
            Select Case tMainWin.Data
                Case 0
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, fft(0), BASS_DATA_FFT512)
                Case 1
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, fft(0), BASS_DATA_FFT1024)
                Case 2
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, fft(0), BASS_DATA_FFT2048)
                Case 3
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, fft(0), BASS_DATA_FFT4096)
            End Select
            If lRslt <> BASSFALSE And lRslt > BASSFALSE Then
                oPlugin.drawVis picVis.hDC, fft, picVis.ScaleHeight, picVis.ScaleWidth
            End If
        Case 1
            Select Case tMainWin.Data
                Case 0
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, Sample(0), 500)
                Case 1
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, Sample(0), 1000)
                Case 2
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, Sample(0), 2000)
                Case 3
                    lRslt = BASS_ChannelGetData(frmMedia.Player.handle, Sample(0), 4000)
            End Select
            If lRslt <> BASSFALSE And lRslt > BASSFALSE Then
                oPlugin.drawVis picVis.hDC, Sample, picVis.ScaleHeight, picVis.ScaleWidth
            End If
    End Select

    Exit Sub
handle:
    Exit Sub
End Sub
