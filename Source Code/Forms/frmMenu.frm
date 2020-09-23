VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   ClientHeight    =   30
   ClientLeft      =   60
   ClientTop       =   150
   ClientWidth     =   9150
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   Begin MSComDlg.CommonDialog cdloColor 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuMain 
      Caption         =   "MP3_ProPlayer"
      Begin VB.Menu mnuAbout 
         Caption         =   "MP3_ProPlayer ..."
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "AlwaysOnTop           A"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Open"
         Begin VB.Menu mnuPlayC 
            Caption         =   "Files ..."
            Index           =   0
         End
         Begin VB.Menu mnuPlayC 
            Caption         =   "Folder ..."
            Index           =   1
         End
         Begin VB.Menu mnuPlayC 
            Caption         =   "Playlist ..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuCDDrive 
         Caption         =   "CD Drive ..."
         Begin VB.Menu mnuCDRom 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPL 
         Caption         =   "Playlist                      W"
      End
      Begin VB.Menu mnuShowEQ 
         Caption         =   "Equalizer                   E"
      End
      Begin VB.Menu mnuDSP 
         Caption         =   "DSP Studio ...           D"
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "M3P Library               L"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "Video"
         Begin VB.Menu mnuScreenC 
            Caption         =   "50 %"
            Index           =   0
         End
         Begin VB.Menu mnuScreenC 
            Caption         =   "100 %"
            Index           =   1
         End
         Begin VB.Menu mnuScreenC 
            Caption         =   "200 %"
            Index           =   2
         End
         Begin VB.Menu mnuScreenC 
            Caption         =   "Full Screen"
            Index           =   3
         End
         Begin VB.Menu mnuScreenC 
            Caption         =   "Exit FullScreen"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu mnuVideoBar 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVideoMode 
            Caption         =   "View Mode"
            Begin VB.Menu mnuVideoModeC 
               Caption         =   "Desktop"
               Index           =   0
            End
            Begin VB.Menu mnuVideoModeC 
               Caption         =   "Window"
               Checked         =   -1  'True
               Index           =   1
            End
         End
         Begin VB.Menu mnuShowDesktop 
            Caption         =   "Show Desktop"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuPb 
         Caption         =   "Playback"
         Begin VB.Menu mnuVolume 
            Caption         =   "Volume"
            Begin VB.Menu mnuVol 
               Caption         =   "Vol inc       +"
               Index           =   0
            End
            Begin VB.Menu mnuVol 
               Caption         =   "Vol dec      -"
               Index           =   1
            End
         End
         Begin VB.Menu mnuMute 
            Caption         =   "Mute"
         End
         Begin VB.Menu mnuBar4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Top                     Home"
            Index           =   0
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "PreviousTrack            B"
            Index           =   1
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Prev 5 seconds        <-"
            Index           =   2
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Bottom                   End"
            Index           =   3
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "NextTrack                  N"
            Index           =   4
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Next 5 seconds        ->"
            Index           =   5
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Play                            L"
            Index           =   6
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Pause                         P"
            Index           =   7
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Stop                          O"
            Index           =   8
         End
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Option"
         Begin VB.Menu mnuOptionC 
            Caption         =   "Preferences ..."
            Index           =   0
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "Time elapsed"
            Index           =   2
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "Time remaining"
            Index           =   3
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "Shuffe                  S"
            Index           =   5
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "Repeat                 R"
            Index           =   6
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "Loop"
            Index           =   7
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuOptionC 
            Caption         =   "Play speed ..."
            Index           =   9
         End
      End
      Begin VB.Menu mnuSkin 
         Caption         =   "Skins"
         Begin VB.Menu mnuSkinC 
            Caption         =   "Skin Browse ..."
            Index           =   0
         End
         Begin VB.Menu mnuSkinC 
            Caption         =   "<<<Get more skin ...>>>"
            Index           =   1
         End
         Begin VB.Menu mnuSkinC 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuSkinC 
            Caption         =   "Default"
            Index           =   3
         End
      End
      Begin VB.Menu mnuVisual 
         Caption         =   "Visualization"
         Begin VB.Menu mnuVisualC 
            Caption         =   "Show Visualization"
            Index           =   0
         End
         Begin VB.Menu mnuVisualC 
            Caption         =   "Configure ..."
            Index           =   1
         End
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit                         X"
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "EditPlaylist"
      Begin VB.Menu mnuPl 
         Caption         =   "Play item"
         Index           =   0
      End
      Begin VB.Menu mnuPl 
         Caption         =   "View item infor ..."
         Index           =   1
      End
      Begin VB.Menu mnuPl 
         Caption         =   "Refresh item infor from file"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuAddC 
         Caption         =   "Add file ..."
         Index           =   0
      End
      Begin VB.Menu mnuAddC 
         Caption         =   "Add folder ..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuSub 
      Caption         =   "Sub"
      Begin VB.Menu mnuSubC 
         Caption         =   "Remove selected file"
         Index           =   0
      End
      Begin VB.Menu mnuSubC 
         Caption         =   "Remove duplicate file"
         Index           =   1
      End
      Begin VB.Menu mnuSubC 
         Caption         =   "Remove missing files"
         Index           =   2
      End
      Begin VB.Menu mnuSubC 
         Caption         =   "Remove All"
         Index           =   3
      End
      Begin VB.Menu mnuSubC 
         Caption         =   "Delete selected file"
         Index           =   4
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Select"
      Begin VB.Menu mnuSel 
         Caption         =   "Select all"
         Index           =   0
      End
      Begin VB.Menu mnuSel 
         Caption         =   "Select none"
         Index           =   1
      End
      Begin VB.Menu mnuSel 
         Caption         =   "Select invert"
         Index           =   2
      End
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Begin VB.Menu mnuSortC 
         Caption         =   "Sort by Artist"
         Index           =   0
      End
      Begin VB.Menu mnuSortC 
         Caption         =   "Sort by Title"
         Index           =   1
      End
      Begin VB.Menu mnuSortC 
         Caption         =   "Sort by Album"
         Index           =   2
      End
      Begin VB.Menu mnuSortC 
         Caption         =   "Sort by Filename"
         Index           =   3
      End
      Begin VB.Menu mnuSortC 
         Caption         =   "Sort by Path and Filename"
         Index           =   4
      End
      Begin VB.Menu mnuSortC 
         Caption         =   "Sort Advance ..."
         Index           =   5
      End
      Begin VB.Menu mnuSortC 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSortStyle 
         Caption         =   "Sort order"
         Begin VB.Menu mnuSortStyleC 
            Caption         =   "Ascending"
            Index           =   0
         End
         Begin VB.Menu mnuSortStyleC 
            Caption         =   "Descending"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuManaList 
      Caption         =   "Manager"
      Begin VB.Menu mnuManList 
         Caption         =   "Load list"
         Index           =   0
      End
      Begin VB.Menu mnuManList 
         Caption         =   "Save list"
         Index           =   1
      End
   End
   Begin VB.Menu mnuEqua 
      Caption         =   "Equalizer"
      Begin VB.Menu mnuEquaLoad 
         Caption         =   "Load"
         Begin VB.Menu mnuEquaLoadC 
            Caption         =   "Preset ..."
            Index           =   0
         End
         Begin VB.Menu mnuEquaLoadC 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuEquaLoadC 
            Caption         =   "From EQF ..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuEquaSave 
         Caption         =   "Save"
         Begin VB.Menu mnuEquaSaveC 
            Caption         =   "Preset ..."
            Index           =   0
         End
         Begin VB.Menu mnuEquaSaveC 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuEquaSaveC 
            Caption         =   "To EQF ..."
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuMainSpec 
      Caption         =   "Spectrum"
      Begin VB.Menu mnuMainSpecC 
         Caption         =   "Oscilliscope"
         Index           =   0
      End
      Begin VB.Menu mnuMainSpecC 
         Caption         =   "Spectrum"
         Index           =   1
      End
      Begin VB.Menu mnuMainSpecC 
         Caption         =   "None"
         Index           =   2
      End
   End
   Begin VB.Menu mnuLib 
      Caption         =   "Library"
      Begin VB.Menu mnuLibOpt 
         Caption         =   "Option"
         Begin VB.Menu mnuDbl 
            Caption         =   "Add to Now Playing on DblClick"
            Index           =   0
         End
         Begin VB.Menu mnuDbl 
            Caption         =   "Add to Now Playing && Play on DblClick"
            Index           =   1
         End
         Begin VB.Menu mnuDbl 
            Caption         =   "Play selected item on DblClick"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLibAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuLibAddC 
            Caption         =   "Add file ..."
            Index           =   0
         End
         Begin VB.Menu mnuLibAddC 
            Caption         =   "Add folder ..."
            Index           =   1
         End
      End
      Begin VB.Menu mnuLibSub 
         Caption         =   "Sub"
         Begin VB.Menu mnuLibSubC 
            Caption         =   "Current item"
            Index           =   0
         End
         Begin VB.Menu mnuLibSubC 
            Caption         =   "All selected item"
            Index           =   1
         End
         Begin VB.Menu mnuLibSubC 
            Caption         =   "All item"
            Index           =   2
         End
         Begin VB.Menu mnuLibSubC 
            Caption         =   "All dead item"
            Index           =   3
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "View"
         Begin VB.Menu mnuViewC 
            Caption         =   "Artist"
            Index           =   0
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Album"
            Index           =   1
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Title"
            Index           =   2
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Genre"
            Index           =   3
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Year"
            Index           =   4
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Bitrate"
            Index           =   5
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Duration"
            Index           =   6
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Frequency"
            Index           =   7
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Rate"
            Index           =   8
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Date add"
            Index           =   9
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Date update"
            Index           =   10
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Play count"
            Index           =   11
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Name"
            Index           =   12
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Path"
            Index           =   13
         End
         Begin VB.Menu mnuViewC 
            Caption         =   "Size"
            Index           =   14
         End
      End
      Begin VB.Menu mnuLibEdit 
         Caption         =   "Edit Library"
         Begin VB.Menu mnuEditLibC 
            Caption         =   "Edit File(s)"
            Index           =   0
         End
         Begin VB.Menu mnuEditLibC 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuEditLibC 
            Caption         =   "Add to Now Playling"
            Index           =   2
         End
         Begin VB.Menu mnuEditLibC 
            Caption         =   "Add to Now Playling && Play"
            Index           =   3
         End
         Begin VB.Menu mnuEditLibC 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Rating"
            Begin VB.Menu mnuRateC 
               Caption         =   "* * * * *"
               Index           =   0
            End
            Begin VB.Menu mnuRateC 
               Caption         =   "* * * *"
               Index           =   1
            End
            Begin VB.Menu mnuRateC 
               Caption         =   "* * *"
               Index           =   2
            End
            Begin VB.Menu mnuRateC 
               Caption         =   "* *"
               Index           =   3
            End
            Begin VB.Menu mnuRateC 
               Caption         =   "*"
               Index           =   4
            End
            Begin VB.Menu mnuRateC 
               Caption         =   "No rating"
               Index           =   5
            End
         End
         Begin VB.Menu mnuLibSplash 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnuSend 
            Caption         =   "Send To"
            Begin VB.Menu mnuSendC 
               Caption         =   "Now playing"
               Index           =   0
            End
            Begin VB.Menu mnuSendC 
               Caption         =   "-"
               Index           =   1
            End
            Begin VB.Menu mnuSendC 
               Caption         =   "My Playlist"
               Index           =   2
               Begin VB.Menu mnuSendCC 
                  Caption         =   ""
                  Index           =   0
               End
            End
            Begin VB.Menu mnuSendC 
               Caption         =   "New playlist"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuTreeView 
         Caption         =   "Edit TreeView"
         Begin VB.Menu mnuTreeViewC 
            Caption         =   "Play"
            Index           =   0
         End
         Begin VB.Menu mnuTreeViewC 
            Caption         =   "Add to Now Playling"
            Index           =   1
         End
         Begin VB.Menu mnuTreeViewC 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuTreeViewC 
            Caption         =   "Delete"
            Index           =   3
         End
         Begin VB.Menu mnuTreeViewC 
            Caption         =   "Rename"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long



Private Sub Form_Load()
        
    Me.Top = -10000
    Me.Left = -10000
    Me.Visible = tAppConfig.bolTaskbar
    Call Add_SkinInstall(tSkinOption.SkinDir)
    Call UpdateCDInfo
    Call LoadSkin(tCurrentSkin.Infor.Name & ".skn", tCurrentSkin.mini)
    
    startODMenus frmMenu, True
    
    With CustomMenu
        .Texture = True
        .UseCustomFonts = False
        .FontBold = False
    End With
    
    
    Load frmMedia
    frmMedia.Init

    mnuOptionC(5).Checked = tPlayerConfig.bolShuffe
    mnuOptionC(6).Checked = tPlayerConfig.bolRepeat
    mnuShowPL.Checked = Not tPlaylistConfig.bolHidePL
    mnuOptionC(2).Checked = Not tPlayerConfig.bolTimer
    mnuOptionC(3).Checked = tPlayerConfig.bolTimer
    mnuShowEQ.Checked = tPlayerConfig.bolShowEQ
    mnuMainSpecC(tSkinVis.intStyle).Checked = True
    mnuMediaControl(2).Caption = "Prev " & tPlayerConfig.intTime & " seconds" & "        <-"
    mnuMediaControl(5).Caption = "Next " & tPlayerConfig.intTime & " seconds" & "        ->"
    mnuDbl(LibOption.intDblClick).Checked = True
    If LibOption.bolUse Then
        mnuLibrary.Enabled = True
    Else
        mnuLibrary.Enabled = False
    End If
    

    
    LoadLang CurrentLang

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    stopODMenus Me
End Sub

Private Sub Form_Resize()
    If bolLoading Then Exit Sub
    If Me.WindowState = vbMinimized Then
        frmMedia.Visible = False
        If tPlaylistConfig.bolHidePL = False Then frmPlayList.Visible = False
        frmMedia.tmrTestPic.Enabled = False
        frmMedia.tmrVisual.Enabled = False
    End If
    If Me.WindowState = vbNormal Then
        frmMedia.Visible = True
        If tPlaylistConfig.bolHidePL = False Then frmPlayList.Visible = True
        frmMedia.tmrTestPic.Enabled = True
        If mnuVisualC(0).Checked = False Then
            frmMedia.tmrVisual.Enabled = True
        End If
    End If
End Sub


Private Sub mnuAbout_Click()
    On Error Resume Next
    frmOption.CallTab (10)
    frmOption.Show
End Sub

Private Sub mnuAddC_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            Call OpenGetFile
        Case 1
            frmBrowse.Show
    End Select
End Sub

Private Sub mnuCDRom_Click(Index As Integer)
    Select Case Index
        Case 0
            Exit Sub
        Case Else
            If BASS_CD_DoorIsOpen(Index - 1) Then
                frmMedia.Player.OpenCDDrive CLng(Index - 1), "Close"
                If BASS_CD_IsReady(Index - 1) Then
                    Dim fso As New FileSystemObject
                    Dim strVCD As String
                    'MPEGAV
                    strVCD = Mid(frmMedia.Player.GetCDDecs(Index - 1), 1, 1)
                    strVCD = strVCD & ":\MPEGAV"
                    If fso.FolderExists(strVCD) Then
                        frmPlayList.List.ClearItem
                        AddSingleFolder (strVCD)
                    Else
                        bolCDPlay = AddCDTrack(Index - 1)
                    End If
                End If
            Else
                If BASS_CD_IsReady(Index - 1) And Not bolCDPlay Then
                    strVCD = Mid(frmMedia.Player.GetCDDecs(Index - 1), 1, 1)
                    strVCD = strVCD & ":\MPEGAV"
                    If fso.FolderExists(strVCD) Then
                        frmPlayList.List.ClearItem
                        Call AddSingleFolder(strVCD)
                    Else
                        bolCDPlay = AddCDTrack(Index - 1)
                    End If
                Else
                    If bolCDPlay Or bolVideoOn Then
                        StopPlayer
                        frmPlayList.List.ClearItem
                        frmMedia.Player.OpenCDDrive CLng(Index - 1), "Open"
                        bolCDPlay = False
                        bolVideoOn = False
                    Else
                        frmMedia.Player.OpenCDDrive CLng(Index - 1), "Open"
                    End If
                End If
            End If
    End Select
End Sub

Private Sub mnuDbl_Click(Index As Integer)
    For i = 0 To 2
        mnuDbl(i).Checked = False
    Next i
    mnuDbl(Index).Checked = True
    LibOption.intDblClick = Index
End Sub

Private Sub mnuDSP_Click()
    mnuDSP.Checked = Not mnuDSP.Checked
    If mnuDSP.Checked Then
        frmDSP.Show
    Else
        Unload frmDSP
    End If
End Sub


Private Sub mnuEditLibC_Click(Index As Integer)
    On Error Resume Next
    Dim tStr As String
    Select Case Index
        Case 0
            frmEdit.Show
        Case 2
            Dim z As Long
            For z = 1 To frmLibrary.lvwFilter.ListItems.Count
                If frmLibrary.lvwFilter.ListItems(z).Selected Then
                    tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                    Call AddToNowPlaying(tStr)
                End If
            Next z
        Case 3
            For z = 1 To frmLibrary.lvwFilter.ListItems.Count
                If frmLibrary.lvwFilter.ListItems(z).Selected Then
                    tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                    Call AddToNowPlaying(tStr)
                End If
            Next z
            For z = 1 To frmPlayList.List.ListItemCount
                If Library(CurrentTrack).Infor.FullName = NowPlaying(frmPlayList.List.Key(z)).Infor.FullName Then
                    Play (z)
                    Exit For
                End If
            Next z
    End Select
End Sub



Private Sub mnuEquaLoadC_Click(Index As Integer)
    On Error GoTo handle
    Select Case Index
        Case 0
            frmEQLoadPreset.Show
        Case 1
        Case 2
            Dim strEQ As String
            With frmMenu.cdloColor
                .DialogTitle = "Load Equalizer"
                .CancelError = True
                .ShowOpen
                strEQ = .Filename
            End With
            Call LoadEQ(strEQ)
    End Select
handle:
        If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub mnuEquaSaveC_Click(Index As Integer)
    Select Case Index
        Case 0
            frmEQSavePreset.Show
        Case 1
        Case 2
            On Error Resume Next
            Dim strFileEQ As String
            With frmMenu.cdloColor
                        .InitDir = App.path
                        .CancelError = True
                        .DialogTitle = "MP3_ProPlayer - Save Equalizer"
                        .DefaultExt = "eqf"
                        .Filter = "MP3_ProPlayer Equalizer Preset (*.eqf)|*.eqf"
                        .ShowSave
                        strFileEQ = .Filename
            End With
            Call SaveEQ(strFileEQ)
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload frmMedia
End Sub

Private Sub mnuLibAddC_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim strFilePath As String
            Dim vFileSelected As Variant
            Dim lngNumber As Long
            With frmMenu.cdloColor
                .InitDir = strLastDir
                .Filename = ""
                .DialogTitle = "MP3_ProPlayer - Open Media"
                .Filter = "M3P - able |*.mp3;*.wma;*.wav;*.aif;*.avi;*.wmv;*.asf;*.mov;*.mpeg;*.mpg;*.mpe;*.mp2;*.ogg" & _
                            "|All Supported Audio Files |*.mp3;*.wma;*.wav;*.ogg" & _
                            "|All Supported Video Files |*.avi;*.wmv;*.asf;*.mov;*.mpeg;*.mpg;*.mpe;*.mp2" & _
                            "|MP3 Files (*.mp3)|*.mp3" & _
                            "|Mpeg File (*.mp3,*.mpeg,*.mpg,*.mpe,*.mp2)|*.mp3;*.mpeg;*.mpg;*.mpe;*.mp2" & _
                            "|Ogg Vorbis Files (*.ogg)" & _
                            "|Wave Files (*.wav)|*.wav" & _
                            "|Windows Media File (*.wma,*.asf,*.avi,*.wmv)|*.wma;*.asf;*.avi;*.wmv" & _
                            "|All File (*.*)|*.*"
                .MaxFileSize = 16384
                .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
                .ShowOpen
                .CancelError = True
                vFileSelected = Split(.Filename, Chr(0))
                If UBound(vFileSelected) = 0 Then
                    strFilePath = frmMenu.cdloColor.Filename
                    If FileLen(strFilePath) > 0 Then
                        Select Case LCase(Right(strFilePath, 4))
                            Case ".mp3", ".wav", ".wma", ".aif", ".asf", ".wmv", ".mpg", ".mpe", ".mp2", "mpeg", ".ogg", ".avi", ".mov"
                                Call AddNewFile(strFilePath)
                        End Select
                    End If
                Else
                    For lngNumber = 1 To UBound(vFileSelected)
                        strFilePath = vFileSelected(0) + "\" & vFileSelected(lngNumber)
                        Select Case LCase(Right(strFilePath, 4))
                            Case ".mp3", ".wav", ".wma", ".aif", ".asf", ".wmv", ".mpg", ".mpe", ".mp2", "mpeg", ".ogg", ".avi", ".mov"
                                Call AddNewFile(strFilePath)
                        End Select
                    Next
                End If
            End With
            Call frmLibrary.ReadLibrary
        Case 1
            On Error GoTo beep
            Dim Browse As Shell
            Dim strFolder As String
            Set Browse = New Shell
            strFolder = Browse.BrowseForFolder(Me.hwnd, "Select a Folder", 0).Items.Item.path
            Call SearchMedia(strFolder, True)
    End Select
beep:
    Set Browse = Nothing
End Sub


Private Sub mnuLibrary_Click()
    On Error Resume Next
    mnuLibrary.Checked = Not mnuLibrary.Checked
    If mnuLibrary.Checked Then
        Load frmLibrary
        frmLibrary.Visible = True
    Else
        Unload frmLibrary
    End If
End Sub


Private Sub mnuLibSubC_Click(Index As Integer)
    Dim z As Long
    Dim y As Long
    Dim tStr As String
    Select Case Index
        Case 0
            Call RemoveTrack(CurrentTrack)
        Case 1
            y = frmLibrary.lvwFilter.ListItems.Count
            For z = 1 To y
                If frmLibrary.lvwFilter.ListItems(z).Selected Then
                    tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                    Call RemoveTrack(TrackIndex(tStr))
                End If
            Next z
        Case 2
            Call RemoveAllTrack
        Case 3
            For z = UBound(Library) - 1 To LBound(Library) Step -1
                If FileExists(Library(z).Infor.FullName) = False Then
                    Call RemoveTrack(z)
                End If
            Next z
    End Select
    Call frmLibrary.ReadLibrary
End Sub

Private Sub mnuMainSpecC_Click(Index As Integer)
    For i = 0 To 2
        mnuMainSpecC(i).Checked = False
    Next i
    tSkinVis.intStyle = Index
    frmMedia.Vis.StyleVis = Index
    mnuMainSpecC(Index).Checked = True
    frmOption.optSkinVis(Index).value = True
End Sub


Private Sub mnuManList_Click(Index As Integer)
    Select Case Index
        Case 0
            Call AddPl
        Case 1
            On Error Resume Next
            Dim strFilePath As String
            'Open Save
            With frmMenu.cdloColor
                .InitDir = strLastDir
                .CancelError = True
                .DialogTitle = "MP3_ProPlayer - Save playlist"
                .DefaultExt = "m3u"
                .Filter = "Media Playlist (*.m3u)|*.m3u|Winamp playlist ( *.pls)|*.pls"
                .ShowSave
                strFilePath = .Filename
            End With
            If strFilePath <> "" Then
                If UCase(Right(strFilePath, 3)) = "M3U" Then Call SaveM3U(strFilePath)
                If UCase(Right(strFilePath, 3)) = "PLS" Then Call SavePLS(strFilePath)
            End If
    End Select
End Sub


Private Sub mnuMediaControl_Click(Index As Integer)
    Select Case Index
        Case 0
            Call GoTop
        Case 1
           Call BackTrack
        Case 2
            Call Prev
        Case 3
            Call GoBottom
        Case 4
           Call NextTrack
        Case 5
            Call Forw
        Case 6
            If frmPlayList.List.ListItemCount Then
                If currentIndex <> 0 Then
                    Call Play(CLng(currentIndex))
                End If
            End If
        Case 7
            With frmMedia
                On Error Resume Next
                Call frmMedia.btnPauseClick
            End With
        Case 8
            If frmMedia.Player.PlayState = 1 Or 2 Then
                Call StopPlayer
            End If
    End Select
End Sub


Private Sub mnuMute_Click()
    mnuMute.Checked = Not mnuMute.Checked
    tPlayerConfig.bolMute = mnuMute.Checked
    If tPlayerConfig.bolMute Then
        Call BASS_SetVolume(0)
    Else
        Call BASS_SetVolume(tPlayerConfig.lngMasterVol)
    End If
    
End Sub

Private Sub mnuOnTop_Click()
    mnuOnTop.Checked = Not mnuOnTop.Checked
    tAppConfig.bolOnTop = mnuOnTop.Checked
    Call AlwaysOnTop(frmMedia, tAppConfig.bolOnTop)
End Sub

Private Sub mnuOptionC_Click(Index As Integer)
    Select Case Index
        Case 0
            frmOption.Show
        Case 2
            If mnuOptionC(2).Checked = False Then
                mnuOptionC(2).Checked = True
                mnuOptionC(3).Checked = False
            Else
                mnuOptionC(2).Checked = False
                mnuOptionC(3).Checked = True
            End If
                tPlayerConfig.bolTimer = mnuOptionC(3).Checked
            Exit Sub
        Case 3
            If mnuOptionC(3).Checked = False Then
                mnuOptionC(3).Checked = True
                mnuOptionC(2).Checked = False
            Else
                mnuOptionC(3).Checked = False
                mnuOptionC(2).Checked = True
            End If
                tPlayerConfig.bolTimer = mnuOptionC(3).Checked
            Exit Sub
        Case 5
            If mnuOptionC(5).Checked Then
                mnuOptionC(5).Checked = False
                tPlayerConfig.bolShuffe = False
            Else
                mnuOptionC(5).Checked = True
                tPlayerConfig.bolShuffe = True
            End If
                frmMedia.btnMedia(5).bolOn = tPlayerConfig.bolShuffe
                frmMedia.btnMedia(5).Refresh
                frmMedia.btnMiniMedia(5).bolOn = tPlayerConfig.bolShuffe
                frmMedia.btnMiniMedia(5).Refresh
        Case 6
            If mnuOptionC(6).Checked = False Then
                mnuOptionC(6).Checked = True
                tPlayerConfig.bolRepeat = True
            Else
                mnuOptionC(6).Checked = False
                tPlayerConfig.bolRepeat = False
            End If
                frmMedia.btnMedia(6).bolOn = tPlayerConfig.bolRepeat
                frmMedia.btnMedia(6).Refresh
                frmMedia.btnMiniMedia(6).bolOn = tPlayerConfig.bolRepeat
                frmMedia.btnMiniMedia(6).Refresh
        Case 7
            If mnuOptionC(7).Checked = False Then
                mnuOptionC(7).Checked = True
                tPlayerConfig.bolLoop = True
            Else
                mnuOptionC(7).Checked = False
                tPlayerConfig.bolLoop = False
            End If
        Case 9
            frmRate.Show
        End Select
End Sub

Private Sub mnuPL_Click(Index As Integer)
    Dim x As Long
    Dim y As Long
    Dim strFile As String
    
    y = frmPlayList.List.Key(currentRIndex)
    strFile = NowPlaying(y).Infor.FullName
    Select Case Index
        Case 0
            Call Play(currentRIndex)
        Case 1
            If currentRIndex <> currentPlayIndex Then
                Select Case LCase(Right(strFile, 4))
                    Case ".mp3"
                        frmTagID.Show
                    Case ".wav"
                        Dim strMsg As String
                        Dim wav As WAVE
                        Call ReadWave(strFile, wav)
                        strMsg = strFile & vbCrLf
                        strMsg = strMsg & IIf(wav.Format = 1, "PCM", "Unknow Type") & ", "
                        strMsg = strMsg & wav.Freq & " Khz, "
                        strMsg = strMsg & wav.Bit & " bit per sample, "
                        strMsg = strMsg & wav.Mode & vbCrLf
                        strMsg = strMsg & "Lenght : " & wav.Lenght & " seconds = " & Time2String(wav.Lenght)
                        MsgBox strMsg, vbOKOnly, "Wave property"
                    Case ".wma"
                        frmWMA.Show
                    Case ".ogg"
                        frmOGG.Show
                    Case ".wmv", ".avi", ".mpg", ".mpe", ".asf"
                        frmVideo_Infor.Show
                    Case Else
                        Exit Sub
                End Select
            End If
        Case 2
            Dim strText As String
            Dim tmpTrack As track
            Call OpenTrack(strFile, tmpTrack)
            strText = ""
            strText = tPlaylistConfig.strDisplay
            strText = Replace(strText, "%1", tmpTrack.Artist)
            strText = Replace(strText, "%2", tmpTrack.Title)
            strText = Replace(strText, "%3", tmpTrack.Album)
            strText = Replace(strText, "%4", tmpTrack.Genre)
            strText = Replace(strText, "%5", tmpTrack.Year)
            strText = Replace(strText, "%6", tmpTrack.Filename)
            strText = Replace(strText, "%7", tmpTrack.FullName)
            NowPlaying(y).strText = strText
            NowPlaying(y).Infor = tmpTrack
            frmPlayList.List.ListItemText(currentRIndex) = strText
            frmPlayList.List.ListItemTime(currentRIndex) = Time2String(tmpTrack.Duration)
            frmPlayList.List.Number = tPlaylistConfig.bolShowNumber
    End Select
End Sub

Private Sub mnuPlayC_Click(Index As Integer)
On Error Resume Next
Dim ret As Long
    Select Case Index
        Case 0
            Call OpenGetFile
        Case 1
            frmBrowse.Show
        Case 2
            Call AddPl
        End Select
End Sub




Private Sub mnuRateC_Click(Index As Integer)
    On Error Resume Next
    Dim i As Long
    Dim x As Long
    For i = 1 To frmLibrary.lvwFilter.ListItems.Count
        If frmLibrary.lvwFilter.ListItems(i).Selected Then
            x = TrackIndex(frmLibrary.lvwFilter.ListItems(i).SubItems(13))
            Library(x).intRate = 5 - Index
        End If
    Next i
    Call WriteDataFile(CurrentTrack)
    Call frmLibrary.ReadLibrary
    frmLibrary.lvwFilter.ListItems.Clear
    Call frmLibrary.DrawList
End Sub

Private Sub mnuScreenC_Click(Index As Integer)
    Select Case Index
        Case 0
            frmVD.Video.Display = HaftSize
            mnuScreenC(4).Enabled = False
        Case 1
            frmVD.Video.Display = DefaultSize
            mnuScreenC(4).Enabled = False
        Case 2
            frmVD.Video.Display = DoubleSize
            mnuScreenC(4).Enabled = False
        Case 3
            frmVD.Video.Display = Fullscreen
            mnuScreenC(4).Enabled = True
        Case 4
            mnuScreenC(4).Enabled = False
            frmVD.Video.Display = CustomizeSize
            frmVD.WindowState = vbNormal
    End Select
End Sub

Private Sub mnuSel_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPlayList.List.SelectItem ("all")
        Case 1
            frmPlayList.List.SelectItem ("none")
            currentIndex = 0
            currentRIndex = 0
        Case 2
            frmPlayList.List.SelectItem ("invert")
    End Select
End Sub


Private Sub mnuSendC_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim tStr As String
            Dim z As Long
            Call SubFolder
            For z = 1 To frmLibrary.lvwFilter.ListItems.Count
                If frmLibrary.lvwFilter.ListItems(z).Selected Then
                    tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                    Call AddFile(tStr)
                End If
            Next z
        Case 3
            frmLibrary.picInput.Left = (frmLibrary.picLibrary.ScaleWidth - frmLibrary.picInput.width) / 2
            frmLibrary.picInput.Visible = True
    End Select
End Sub

Private Sub mnuSendCC_Click(Index As Integer)
    On Error Resume Next
    Dim i As Long
    Dim x As Long
    Dim tStr As String
    Dim strName As String
    strName = LCase(mnuSendCC(Index).Caption)
    For i = 0 To UBound(Playlist) - 1
        If strName = LCase(Playlist(i).Name) Then
            x = i
            Exit For
        End If
    Next i
    
    Dim arrPL() As String
    Dim J As Long
    Dim z As Long
    arrPL = Split(Playlist(x).file & ",", ",")
    While i <= frmLibrary.lvwFilter.ListItems.Count
        If frmLibrary.lvwFilter.ListItems(i).Selected Then
            tStr = frmLibrary.lvwFilter.ListItems(i).SubItems(13)
            z = TrackIndex(tStr)
            Debug.Print z
            For J = 0 To UBound(arrPL) - 1
                If CStr(z) <> arrPL(J) Then
                    Playlist(x).file = Playlist(x).file & "," & z
                    Exit For
                End If
            Next J
        End If
        i = i + 1
    Wend
    WriteDataList x
End Sub

Private Sub mnuShowDesktop_Click()
    mnuShowDesktop.Checked = Not mnuShowDesktop.Checked
    If mnuShowDesktop.Checked Then
        WinAPI.HideDesktop False
    Else
        WinAPI.HideDesktop True
    End If
End Sub

Private Sub mnuShowEQ_Click()
    If mnuShowEQ.Checked = False Then
        mnuShowEQ.Checked = True
        tPlayerConfig.bolShowEQ = True
    Else
        mnuShowEQ.Checked = False
        tPlayerConfig.bolShowEQ = False
    End If
    Call ShowEQ
    frmMedia.btnMedia(7).bolOn = tPlayerConfig.bolShowEQ
    frmMedia.btnMedia(7).Refresh
    frmMedia.btnMiniMedia(7).bolOn = tPlayerConfig.bolShowEQ
    frmMedia.btnMiniMedia(7).Refresh
End Sub

Private Sub mnuShowPL_Click()
    If mnuShowPL.Checked = False Then
        mnuShowPL.Checked = True
        tPlaylistConfig.bolHidePL = False
        frmMedia.btnMedia(8).bolOn = True
        frmMedia.btnMiniMedia(8).bolOn = True
        frmPlayList.Show
    Else
        mnuShowPL.Checked = False
        frmPlayList.Hide
        tPlaylistConfig.bolHidePL = True
        frmMedia.btnMedia(8).bolOn = False
        frmMedia.btnMiniMedia(8).bolOn = False
    End If
    frmMedia.btnMedia(8).Refresh
    frmMedia.btnMiniMedia(8).Refresh
End Sub

Private Sub mnuSkinC_Click(Index As Integer)
    On Error Resume Next
    Dim flag As Integer
    Dim strName As String
    Dim strNewSkin As String
    
    If mnuSkinC(Index).Checked = True Then
        MsgBox "Current Skin already load", vbOKOnly, "Chose skin"
        Exit Sub
    End If
    For i = 0 To mnuSkinC.Count - 1
        mnuSkinC(i).Checked = False
    Next i
    

    Select Case Index
        Case 0
            frmOption.CallTab (6)
            frmOption.Show
        Case 1
            ShellExecute Me.hwnd, vbNullString, "www.freewebs\Hanachacha.com", vbNullString, vbNullString, SW_SHOWNORMAL
        Case 3
            Call LoadSkin("Default.skn", tCurrentSkin.mini)
        Case Else
            WriteINI "Demension", "PlaylistTop", frmPlayList.Top, strFileconfig
            WriteINI "Demension", "PlaylistLeft", frmPlayList.Left, strFileconfig
            WriteINI "Demension", "PlaylistWidth", frmPlayList.width, strFileconfig
            WriteINI "Demension", "PlaylistHeight", frmPlayList.height, strFileconfig
            
            Call LoadSkin(mnuSkinC(Index).Caption & ".skn", tCurrentSkin.mini)
    End Select
End Sub

Private Sub mnuSortC_Click(Index As Integer)
    On Error Resume Next
    Dim strSorted As String
    
    For i = 0 To mnuSortC.Count - 1
        mnuSortC(i).Checked = False
    Next i
    mnuSortC(Index).Checked = True
    If Index <> 5 Then
        strSorted = Mid(frmMenu.mnuSortC(Index).Caption, Len("Sort by") + 1)
    Else
        strSorted = tPlaylistConfig.strSortString
    End If
    tPlaylistConfig.strSortString = strSorted
    Select Case Index
        Case 0
            tPlaylistConfig.intSortKey = 0
            Call SetSortString("Artist")
        Case 1
            tPlaylistConfig.intSortKey = 1
            Call SetSortString("Title")
        Case 2
            tPlaylistConfig.intSortKey = 2
            Call SetSortString("Album")
        Case 3
            tPlaylistConfig.intSortKey = 3
            Call SetSortString("Filename")
        Case 4
            tPlaylistConfig.intSortKey = 4
            Call SetSortString("Pathname")
        Case 5
            frmCustomSort.Show
            Exit Sub
    End Select
    
    frmPlayList.List.IsSort = True
    
    If currentPlayIndex <> 0 Then
        currentPlayIndex = frmPlayList.List.CurrentPlayItem
    End If
End Sub

Private Sub mnuSortStyleC_Click(Index As Integer)
    For i = 0 To 1
        mnuSortStyleC(i).Checked = False
    Next i
    mnuSortStyleC(Index).Checked = True
    Select Case Index
        Case 0
            frmPlayList.List.IsSortStyle = True
        Case 1
            frmPlayList.List.IsSortStyle = False
    End Select
End Sub

Private Sub mnuSubC_Click(Index As Integer)
    Dim x As Long
    Dim y As Long

    Select Case Index
        Case 4
            frmPlayList.MousePointer = 11
            For i = frmPlayList.List.ListItemCount To 1 Step -1
                If frmPlayList.List.ListItemSelect(i) = True Then Call KillFile(i)
            Next i
            frmPlayList.MousePointer = 0
        Case 3
            frmPlayList.MousePointer = 11
            Call SubFolder
            frmPlayList.MousePointer = 0
        Case 2
            frmPlayList.MousePointer = 11
            For x = frmPlayList.List.ListItemCount To 1 Step -1
                If Not FileExists(NowPlaying(frmPlayList.List.Key(x)).Infor.FullName) Then
                    Call SubFile(x)
                End If
            Next x
            frmPlayList.MousePointer = 0
        Case 1
            frmPlayList.MousePointer = 11
            For x = 1 To UBound(NowPlaying)
                For y = frmPlayList.List.ListItemCount To x + 1 Step -1
                    If NowPlaying(frmPlayList.List.Key(y)).Infor.FullName = NowPlaying(frmPlayList.List.Key(x)).Infor.FullName Then
                        Call SubFile(y)
                    End If
                Next y
            Next x
            frmPlayList.MousePointer = 0
        Case 0
            frmPlayList.MousePointer = 11
            For x = frmPlayList.List.ListItemCount To 1 Step -1
                If frmPlayList.List.ListItemSelect(x) = True Then
                    SubFile (x)
                End If
            Next x
            frmPlayList.MousePointer = 0
    End Select
End Sub


Private Sub mnuTreeViewC_Click(Index As Integer)
    On Error Resume Next
    Dim z As Long
    Dim x As Long
    Dim tStr As String
    Select Case Index
        Case 0
            For z = 1 To frmLibrary.lvwFilter.ListItems.Count
                tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                Call AddFile(tStr)
            Next z
            For z = 1 To frmPlayList.List.ListItemCount
                If Library(CurrentTrack).Infor.FullName = NowPlaying(frmPlayList.List.Key(z)).Infor.FullName Then
                    Play (z)
                    Exit For
                End If
            Next z
        Case 1
            For z = 1 To frmLibrary.lvwFilter.ListItems.Count
                tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                Call AddFile(tStr)
            Next z
        Case 3
            Dim y As Long
            Dim tKey As String
            tKey = LCase(Mid(frmLibrary.CurrentNode.Key, 1, InStr(1, frmLibrary.CurrentNode.Key, " ", vbBinaryCompare) - 1))
            y = frmLibrary.lvwFilter.ListItems.Count
            Select Case tKey
                Case "audio", "video", "auto"
                    For z = 1 To y
                        tStr = frmLibrary.lvwFilter.ListItems(z).SubItems(13)
                        Call RemoveTrack(TrackIndex(tStr))
                    Next z
                Case "my"
                    Dim s As Long
                    s = -1
                    If LCase(frmLibrary.CurrentNode.Parent.text) = "my playlist" Then
                        For x = 0 To UBound(Playlist) - 1
                            If Len(frmLibrary.CurrentNode.text) = Len(Playlist(x).Name) Then
                                If LCase(frmLibrary.CurrentNode.text) = LCase(Playlist(x).Name) Then
                                    s = x
                                    Exit For
                                End If
                            End If
                        Next x
                        If s = -1 Then Exit Sub
                        If UBound(Playlist) = 1 Then
                            ReDim Playlist(0)
                        Else
                            For x = s To UBound(Playlist) - 1
                                Playlist(x) = Playlist(x + 1)
                            Next x
                            ReDim Preserve Playlist(UBound(Playlist) - 1)
                        End If
                    End If
                Case "now"
                    Call SubFolder
                    frmLibrary.lvwFilter.ListItems.Clear
                    frmLibrary.DrawList
            End Select
            Call frmLibrary.ReadLibrary
        Case 4
            Call frmLibrary.tvwLibrary.StartLabelEdit
    End Select
End Sub

Private Sub mnuVideoModeC_Click(Index As Integer)
    If mnuVideoModeC(Index).Checked = False Then
        Select Case Index
            Case 0
                mnuVideoModeC(0).Checked = True
                mnuVideoModeC(1).Checked = False
                frmVD.Video.VideoMode = Desktop
                frmVD.Visible = False
                mnuShowDesktop.Enabled = True
                mnuShowDesktop.Checked = False
            Case 1
                mnuVideoModeC(0).Checked = False
                mnuVideoModeC(1).Checked = True
                frmVD.Video.VideoMode = Window
                frmVD.Visible = True
                mnuShowDesktop.Enabled = False
        End Select
    End If
End Sub

Private Sub mnuViewC_Click(Index As Integer)
    On Error Resume Next
    Dim tmp As Long
    mnuViewC(Index).Checked = Not mnuViewC(Index).Checked
    LibOption.ColView(Index) = frmMenu.mnuViewC(Index).Checked
    If mnuViewC(Index).Checked = False Then
        frmLibrary.lvwFilter.ColumnHeaders(Index + 1).width = 0
    Else
        frmLibrary.lvwFilter.ColumnHeaders(Index + 1).width = LibOption.ColWidth(Index)
    End If
    Call frmLibrary.DrawList
    Dim r As Long, z As Long
    For r = z To z + frmLibrary.ItemPage
        If i > frmLibrary.lvwFilter.ListItems.Count Then Exit For
        If frmLibrary.lvwFilter.ListItems(r).Selected = True Then
            Call frmLibrary.DrawSelected(r)
        End If
    Next r
End Sub

Private Sub mnuVisualC_Click(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        mnuVisualC(0).Checked = Not mnuVisualC(0).Checked
        If mnuVisualC(0).Checked = False Then
            Unload frmVisual
        Else
            frmVisual.Show
        End If
    End If
    If Index = 1 Then
        frmOption.CallTab (7)
        frmOption.Show
    End If
End Sub


Private Sub mnuVol_Click(Index As Integer)
    Dim lngVol As Long
    Select Case Index
        Case 0
            If frmMedia.sldVolume.value < 100 Then
                lngVol = frmMedia.sldVolume.value + 1
                Call frmMedia.volume(lngVol)
            End If
        Case 1
            If frmMedia.sldVolume.value > 1 Then
                lngVol = frmMedia.sldVolume.value - 1
                Call frmMedia.volume(lngVol)
            End If
    End Select
End Sub

Public Sub Add_SkinInstall(strFolder As String)
    On Error Resume Next
    Dim DirList As New Collection
    Dim temp As String
    Dim str As String
    Dim flag As Long
    
    
    '
    For i = 4 To mnuSkinC.Count - 1
        Unload mnuSkinC(CInt(i))
    Next
    DirList.Add RepairPath(strFolder, "")
    flag = 3
    Do While DirList.Count
        temp = Dir$(DirList(1), vbDirectory)
        Do Until temp = ""
            If temp = "." Or temp = ".." Then
            ElseIf (GetAttr(RepairPath(DirList(1), temp)) And vbDirectory) = vbDirectory Then
            ElseIf InStr(temp, ".skn") Then
                If temp <> "Default.skn" Then
                    DirList.Add RepairPath(DirList(1), temp) & "\"
                    flag = flag + 1
                    Load mnuSkinC(flag)
                    mnuSkinC(flag).Caption = Mid(temp, 1, Len(temp) - 4)
                    mnuSkinC(flag).Enabled = True
                End If
            End If
            temp = Dir$
        Loop
        DirList.Remove 1
    Loop
End Sub

