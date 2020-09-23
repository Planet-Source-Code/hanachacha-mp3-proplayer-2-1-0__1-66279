Attribute VB_Name = "modM3P"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit


' MEMORY
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

' FILE
Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1

Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' WAV Header
Private Type WAVEHEADER_RIFF    ' == 12 bytes ==
    RIFF As Long                ' "RIFF" = &H46464952
    riffBlockSize As Long       ' reclen - 8
    riffBlockType As Long       ' "WAVE" = &H45564157
End Type

Private Type WAVEFORMAT         ' == 24 bytes ==
    wfBlockType As Long         ' "fmt " = &H20746D66
    wfBlockSize As Long
    ' == block size begins from here = 16 bytes
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type

Private Type WAVEHEADER_data    ' == 8 bytes ==
   dataBlockType As Long        ' "data" = &H61746164
   dataBlockSize As Long        ' reclen - 44
End Type

Dim wr As WAVEHEADER_RIFF
Dim wf As WAVEFORMAT
Dim wd As WAVEHEADER_data


Public BUFSTEP As Long        ' memory allocation unit
Public input_ As Long         ' current input source
Public recPTR As Long         ' a recording pointer to a memory location
Public reclen As Long         ' buffer length

Public rchan As Long          ' recording channel


Dim i As Long
Dim strFilePath As String

'Variables
Public currentIndex As Long 'Index of current selected item by left click
Public currentRIndex As Long 'Index of current selected item by right click
Public currentPlayIndex As Long 'Index of current item is playing
'Play media and update stream information
Public Sub Play(Index As Long)
    On Error GoTo beep
    Dim x As Long
    Dim NowIndex As Long 'index pointer to NowPlaying()
    
    If Index = 0 Or Index > frmPlayList.List.ListItemCount Then StopPlayer: Exit Sub
    
    frmMedia.PlayCrossFade
    
    NowIndex = frmPlayList.List.Key(Index) 'Pointer in NowPlaying()
    
    frmPlayList.List.CurrentPlayItem = Index
    frmPlayList.List.DisplayList
    frmPlayList.lblCurrentperTotal.Caption = Index & "/" & frmPlayList.List.ListItemCount
    currentPlayIndex = Index
    
    
    strFilePath = NowPlaying(NowIndex).Infor.FullName
    With frmMedia
        .btnMedia(4).Enabled = True
        .btnMiniMedia(4).Enabled = True
        .srlInfor.Title = "Now loading ..."
        .srlMiniInfor.Title = "Now loading ..."
        
        If LibOption.bolUse Then
            x = TrackIndex(strFilePath)
            If x = -1 Then
                Call OpenTrack(strFilePath, tCurrentTrack)
            Else
                tCurrentTrack = Library(x).Infor
            End If
        End If
        
        If (tPlaylistConfig.intLoadID = 1 And NowPlaying(NowIndex).strText <> GetShortName(NowPlaying(NowIndex).Infor.FullName)) Or tPlaylistConfig.intLoadID = 2 Then
            Call OpenTrack(strFilePath, tCurrentTrack)
            NowPlaying(NowIndex).Infor = tCurrentTrack
            Dim strText As String
            strText = ""
            strText = tPlaylistConfig.strDisplay
            strText = Replace(strText, "%1", NowPlaying(NowIndex).Infor.Artist)
            strText = Replace(strText, "%2", NowPlaying(NowIndex).Infor.Title)
            strText = Replace(strText, "%3", NowPlaying(NowIndex).Infor.Album)
            strText = Replace(strText, "%4", NowPlaying(NowIndex).Infor.Genre)
            strText = Replace(strText, "%5", NowPlaying(NowIndex).Infor.Year)
            strText = Replace(strText, "%6", NowPlaying(NowIndex).Infor.Filename)
            strText = Replace(strText, "%7", NowPlaying(NowIndex).Infor.FullName)
            frmPlayList.List.ListItemText(x) = strText
            frmPlayList.List.ListItemTime(x) = Time2String(NowPlaying(NowIndex).Infor.Duration)
            frmPlayList.List.Number = tPlaylistConfig.bolShowNumber
        Else
            tCurrentTrack = NowPlaying(NowIndex).Infor
        End If
        
        
        .srlInfor.Title = NowPlaying(NowIndex).Infor.Artist & "--" & NowPlaying(NowIndex).Infor.Title
        .srlMiniInfor.Title = NowPlaying(NowIndex).Infor.Artist & "--" & NowPlaying(NowIndex).Infor.Title
        
        Select Case LCase(NowPlaying(NowIndex).Infor.ExtType)
            Case ".mp3", ".wma", ".wav", ".cda", ".mid", ".ogg"
                If bolVideoOn Then
                    frmVD.Video.StopVideo
                    frmVD.Hide: frmVD.bolShow = False
                End If
                .Player.OpenFile strFilePath, tDevice.SoundD.WaveWrite
                .lblMpgInfo(0).Caption = Int(((FileLen(strFilePath) * 0.8) / .Player.Duration) / 100)
                .lblMpgInfo(1).Caption = .Player.GetInfo(1) / 1000
                'Start Equal
                Call SetEqual(.Player.handle)
                If bolEQEnabled Then
                    For i = 0 To 9
                        Call UpdateEqual(i, .sldEqua(i).value)
                    Next i
                Else
                    For i = 0 To 9
                        Call BASS_ChannelRemoveFX(.Player.handle, EQ(i))
                    Next i
                End If
                'testDX
                Call SetFX(.Player.handle)
                For i = 0 To 7
                    Call UpdateFX(i)
                Next i
                'test WinampDSP
                If tWinampDSP.bolEnabled Then
                    If frmOption.lvwDSP.ListItems.Count > 0 Then
                        Call StartDSP
                        RestartDsp
                    End If
                End If
                Wait 10
                'Play
                .sldPosition.max = .Player.Duration
                .lblDuration.Caption = Time2String(.Player.Duration)
                NowPlaying(NowIndex).Infor.Duration = .Player.Duration
                If tDevice.SoundD.WaveWrite = False Then
                    .Player.PlayStream
                    If frmMenu.mnuVisualC(0).Checked = False Then
                        .tmrVisual.Enabled = True
                    Else
                        frmVisual.Show
                    End If
                    If frmOption.lstPlugins.ListCount > 0 Then
                        If frmOption.cboPlugins.ListCount > 0 Then
                            Call Stop_VisPlg
                            Call Start_VisPlg
                        End If
                    End If
                Else
                    Dim strWaveOut As String
                    If tDevice.WaveD.bolAutoFilename Then
                        MkDir tDevice.WaveD.strWaveOutput & "\" & tCurrentTrack.Artist
                        MkDir tDevice.WaveD.strWaveOutput & "\" & tCurrentTrack.Artist & "\" & tCurrentTrack.Album
                        strWaveOut = tDevice.WaveD.strWaveOutput & "\" & tCurrentTrack.Artist & "\" & tCurrentTrack.Album & "\" & Mid(tCurrentTrack.Filename, 1, Len(tCurrentTrack.Filename) - Len(tCurrentTrack.ExtType)) & ".wav"
                    Else
                        strWaveOut = tDevice.WaveD.strWaveOutput & "\" & Mid(tCurrentTrack.Filename, 1, Len(tCurrentTrack.Filename) - Len(tCurrentTrack.ExtType)) & ".wav"
                    End If
                    If frmMenu.mnuVisualC(0).Checked = False Then
                        .tmrVisual.Enabled = False
                    Else
                        frmVisual.Hide
                        .tmrVisual.Enabled = False
                    End If
                    Call .Player.WriteWave(strWaveOut)
                End If
                bolVideoOn = False
            Case ".aif", ".avi", ".wmv", ".mpe", ".mpg", ".asf", ".dat", ".mpeg", ".mov", ".mp1", ".mp2", ".mp4", ".ogm"
                If .Player.PlayState = 1 Then
                    .Player.StopStream
                    .Vis.doStop
                End If
                .lblMpgInfo(0).Caption = tCurrentTrack.bitrate
                .lblMpgInfo(1).Caption = tCurrentTrack.Frequency / 1000
                With frmVD
                    .Show
                    .bolShow = True
                    .Video.OpenVideo strFilePath
                    .Video.volume = frmMedia.sldVolume.value
                    NowPlaying(NowIndex).Infor.Duration = .Video.Duration
                    .prgStatus.max = NowPlaying(NowIndex).Infor.Duration
                    frmPlayList.List.ListItemTime(frmPlayList.List.CurrentPlayItem) = Time2String(NowPlaying(NowIndex).Infor.Duration)
                End With
                
                .sldPosition.max = NowPlaying(NowIndex).Infor.Duration
                .lblDuration.Caption = Time2String(NowPlaying(NowIndex).Infor.Duration)
                
                If frmMenu.mnuVideoModeC(0).Checked Then
                    frmVD.Video.VideoMode = Desktop
                    frmVD.Visible = False
                    InvalidateRect 0&, ByVal 0, 1&
                    WinAPI.SettingWallpaper ClearWall
                    frmMenu.mnuShowDesktop.Enabled = True
                    frmMenu.mnuShowDesktop.Checked = False
                Else
                    If frmMenu.mnuVideoModeC(1).Checked Then
                        frmVD.Video.VideoMode = Window
                        frmVD.Video.Display = tDevice.VideoD.intDefaultScreen
                        frmVD.Visible = True
                        frmMenu.mnuShowDesktop.Enabled = False
                    End If
                End If
                
                frmVD.Video.PlayVideo
                bolVideoOn = True
                WinAPI.ScreenSaverActive False
                
                If tDevice.SoundD.WaveWrite = True Then
                    StartRecording
                End If
            Case Else
                
                MsgBox "Not supported this type !!!", vbInformation, "Error File type"
                StopPlayer
                Exit Sub
        End Select
        .sldMiniPosition.max = .sldPosition.max
        .lblMiniDur.Caption = .lblDuration.Caption
        .lblMiniMpgInfo(0).Caption = .lblMpgInfo(0).Caption
        .lblMiniMpgInfo(1).Caption = .lblMpgInfo(1).Caption
        .btnMedia(2).Enabled = True
        .btnMiniMedia(2).Enabled = True
    End With
    
    If LibOption.bolUse Then
        If x <> -1 Then Library(x).intPlaycount = Library(x).intPlaycount + 1
    End If
    
    With frmPlayList
        .lblTotalTime = Time2String(CalTime)
        .List.ItemVisible (Index)
        .sldPl.value = Index
    End With
    
    If frmMenu.mnuVisualC(0).Checked Then frmVisual.lblTitle.Caption = frmMedia.srlInfor.Title
    If tAppConfig.bolSysTray Then sysTray.SysTip frmMedia.hwnd, "MP3_ProPlayer - " & tCurrentTrack.Artist & "--" & tCurrentTrack.Title
    Exit Sub
beep:
    If Err.Number <> 0 Then
        Call frmMedia.PlayNext
    End If
End Sub
Public Sub Forw()
    On Error Resume Next
    With frmMedia
        .sldPosition.value = .sldPosition.value + tPlayerConfig.intTime ' .Player.Position
        If Not bolVideoOn Then
            .Player.Position = .sldPosition.value
        Else
            frmVD.Video.Position = .sldPosition.value
        End If
    End With
End Sub
Public Sub Prev()
    On Error Resume Next
    With frmMedia
        .sldPosition.value = .sldPosition.value - tPlayerConfig.intTime ' .Player.Position
        If Not bolVideoOn Then
            .Player.Position = .sldPosition.value
        Else
            frmVD.Video.Position = .sldPosition.value
        End If
    End With
End Sub
Public Sub NextTrack()
On Error Resume Next
    With frmPlayList
        Dim NextIndex As Long
        If Not bolCDPlay Then
            If tPlayerConfig.bolShuffe = False Then
                If currentPlayIndex = .List.ListItemCount Then
                    Exit Sub
                Else
                    NextIndex = currentPlayIndex + 1
                End If
            Else
                Randomize Timer
                NextIndex = .List.ListItemCount * Rnd
            End If
            Call Play(NextIndex)
        Else
            If tPlayerConfig.bolShuffe = False Then
                If tCDplay.CurrentTrack = tCDplay.TotalTrack Then
                    Exit Sub
                Else
                    NextIndex = tCDplay.CurrentTrack + 1
                End If
            Else
                Randomize Timer
                NextIndex = tCDplay.TotalTrack * Rnd
            End If
            Call PlayCD(tCDplay.CurrentDrive, NextIndex)
        End If
    End With
End Sub
Public Sub BackTrack()
    With frmPlayList
        Dim NextIndex As Long
        If Not bolCDPlay Then
            If tPlayerConfig.bolShuffe = False Then
                If currentPlayIndex = 1 Then
                    Exit Sub
                Else
                    NextIndex = currentPlayIndex - 1
                End If
            Else
                Randomize Timer
                NextIndex = .List.ListItemCount * Rnd
            End If
            Call Play(NextIndex)
        Else
            If tPlayerConfig.bolShuffe = False Then
                If tCDplay.CurrentTrack = 0 Then
                    Exit Sub
                Else
                    NextIndex = tCDplay.CurrentTrack - 1
                End If
            Else
                Randomize Timer
                NextIndex = tCDplay.TotalTrack * Rnd
            End If
            Call PlayCD(tCDplay.CurrentDrive, NextIndex)
        End If
    End With
End Sub
Public Sub StopPlayer()
    With frmMedia
        If Not bolVideoOn Then
            If tDevice.SoundD.WaveWrite Then .Player.ExitWaveWrite
            .Player.StopStream
        Else
            frmVD.Video.StopVideo
            frmVD.Video.CloseVideo
            frmVD.Video.VideoMode = Window
            frmVD.Hide: frmVD.bolShow = False
            bolVideoOn = False
            WinAPI.ScreenSaverActive True
            WinAPI.SettingWallpaper RestoreWall
            If tDevice.SoundD.WaveWrite = True Then
                Dim strWaveOut As String
                If tDevice.WaveD.bolAutoFilename Then
                    MkDir tDevice.WaveD.strWaveOutput & "\" & tCurrentTrack.Artist
                    MkDir tDevice.WaveD.strWaveOutput & "\" & tCurrentTrack.Artist & "\" & tCurrentTrack.Album
                    strWaveOut = tDevice.WaveD.strWaveOutput & "\" & tCurrentTrack.Artist & "\" & tCurrentTrack.Album & "\" & Mid(tCurrentTrack.Filename, 1, Len(tCurrentTrack.Filename) - Len(tCurrentTrack.ExtType)) & ".wav"
                Else
                    strWaveOut = tDevice.WaveD.strWaveOutput & "\" & Mid(tCurrentTrack.Filename, 1, Len(tCurrentTrack.Filename) - Len(tCurrentTrack.ExtType)) & ".wav"
                End If
                StopRecording
                WriteToDisk strWaveOut
            End If
        End If
        .prgVU(0).value = 32768
        .prgVU(1).value = 0
        .Vis.doStop
        .sldPosition.value = 0
        .lblDuration.Caption = "00:00"
        .lblPosition.Caption = "00:00"
        .btnMedia(2).Enabled = False
        .btnMiniMedia(2).Enabled = False
        .btnMedia(4).Enabled = False
        .btnMiniMedia(4).Enabled = False
        tPlaylistConfig.intLastFile = currentPlayIndex
        currentPlayIndex = 0
        .srlInfor.Title = strTitle
        .srlMiniInfor.Title = strTitle
    End With
End Sub
Public Sub PlayCD(ByVal nDrive As Long, ByVal nTrack As Long)
    With frmMedia
        Wait 10
        Call StopPlayer
        currentPlayIndex = nTrack + 1
        Call SetEqual(.Player)
        
        If bolEQEnabled Then
            For i = 0 To 9
                Call UpdateEqual(i, .sldEqua(i).value)
            Next i
        Else
            For i = 0 To 9
                Call BASS_ChannelRemoveFX(.Player.handle, EQ(i))
            Next i
        End If
        
        'testDX
        Call SetFX(.Player)
        For i = 0 To 7
            Call UpdateFX(i)
        Next i
        'test WinampDSP
        If tWinampDSP.bolEnabled Then
            If frmOption.lvwDSP.ListItems.Count > 0 Then
                Call StartDSP
                RestartDsp
            End If
        End If
        
        'Play
        tCDplay.CurrentTrack = nTrack
        .Player.ReadCDTrack nDrive, nTrack, False
        .btnMedia(1).bolOn = True
        .btnMedia(2).bolOn = False
        .btnMedia(2).Enabled = True
        .btnMedia(4).Enabled = True
        .btnMiniMedia(4).Enabled = True
        .btnMedia(1).Refresh
        .btnMedia(2).Refresh
        .sldPosition.max = .Player.Duration
        .sldPosition.min = 0
        .sldMiniPosition.max = .sldPosition.max
        .sldMiniPosition.min = 0
        .lblDuration.Caption = Time2String(.Player.Duration)
        .lblMiniDur.Caption = .lblDuration.Caption
        .srlInfor.Title = frmPlayList.List.ListItemText(nTrack + 1)
        .srlMiniInfor.Title = .srlInfor.Title
        frmPlayList.List.CurrentPlayItem = nTrack + 1
        frmPlayList.List.DisplayList
        frmPlayList.lblCurrentperTotal.Caption = nTrack + 1 & "/" & frmPlayList.List.ListItemCount
        If Not tDevice.SoundD.WaveWrite Then
            .Player.PlayStream
        Else
            frmPlayList.List.ClearItem
            Dim strPath As String
            strPath = Mid(.Player.GetCDDecs(tCDplay.CurrentDrive), 1, 1)
            strPath = strPath & ":"
            AddSingleFolder (strPath)
            Play (nTrack + 1)
        End If
    End With
End Sub
Public Sub GoBottom()
    On Error Resume Next
    If bolCDPlay = False Then
        Call Play(frmPlayList.List.ListItemCount)
    Else
        Call PlayCD(tCDplay.CurrentDrive, tCDplay.TotalTrack - 1)
    End If
End Sub
Public Sub GoTop()
    On Error Resume Next
    If bolCDPlay = False Then
        Call Play(1)
    Else
        Call PlayCD(tCDplay.CurrentDrive, 0)
    End If
End Sub



' buffer the recorded data
Public Function RecordingCallback(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
    ' increase buffer size if needed
    If ((reclen Mod BUFSTEP) + length >= BUFSTEP) Then
        recPTR = GlobalReAlloc(ByVal recPTR, ((reclen + length) / BUFSTEP + 1) * BUFSTEP, GMEM_MOVEABLE)
        If recPTR = 0 Then
            rchan = 0
            RecordingCallback = BASSFALSE ' stop recording
            Exit Function
        End If
    End If
    ' buffer the data
    Call CopyMemory(ByVal recPTR + reclen, ByVal buffer, length)
    reclen = reclen + length
    RecordingCallback = BASSTRUE ' continue recording
End Function

Public Sub StartRecording()
    ' free old recording
    If (recPTR) Then
        Call GlobalFree(ByVal recPTR)
        recPTR = 0
    End If

    ' allocate initial buffer and make space for WAVE header
    recPTR = GlobalAlloc(GMEM_FIXED, BUFSTEP)
    reclen = 44

    ' fill the WAVE header
    wf.wFormatTag = 1
    wf.nChannels = 2
    wf.wBitsPerSample = 16
    wf.nSamplesPerSec = 44100
    wf.nBlockAlign = wf.nChannels * wf.wBitsPerSample / 8
    wf.nAvgBytesPerSec = wf.nSamplesPerSec * wf.nBlockAlign

    ' Set WAV "fmt " header
    wf.wfBlockType = &H20746D66      ' "fmt "
    wf.wfBlockSize = 16

    ' Set WAV "RIFF" header
    wr.RIFF = &H46464952             ' "RIFF"
    wr.riffBlockSize = 0             ' after recording
    wr.riffBlockType = &H45564157    ' "WAVE"

    ' set WAV "data" header
    wd.dataBlockType = &H61746164    ' "data"
    wd.dataBlockSize = 0             ' after recording

    ' copy WAV Header to Memory
    Call CopyMemory(ByVal recPTR, wr, LenB(wr))        ' "RIFF"
    Call CopyMemory(ByVal recPTR + 12, wf, LenB(wf))   ' "fmt "
    Call CopyMemory(ByVal recPTR + 36, wd, LenB(wd))   ' "data"

    ' start recording @ 44100hz 16-bit stereo
    rchan = BASS_RecordStart(44100, 2, 0, AddressOf RecordingCallback, 0)
    
    If (rchan = 0) Then
        Call GlobalFree(ByVal recPTR)
        recPTR = 0
        Exit Sub
    End If
End Sub

Public Sub StopRecording()
    Call BASS_ChannelStop(rchan)
    rchan = 0
    
    ' complete the WAVE header
    wr.riffBlockSize = reclen - 8
    wd.dataBlockSize = reclen - 44

    Call CopyMemory(ByVal recPTR + 4, wr.riffBlockSize, LenB(wr.riffBlockSize))
    Call CopyMemory(ByVal recPTR + 40, wd.dataBlockSize, LenB(wd.dataBlockSize))

End Sub

' write the recorded data to disk
Public Sub WriteToDisk(strFile As String)
    Dim FileHandle As Long, ret As Long, OF As OFSTRUCT

    FileHandle = OpenFile(strFile, OF, OF_CREATE)

    If (FileHandle = 0) Then
        Exit Sub
    End If

    Call WriteFile(FileHandle, ByVal recPTR, reclen, ret, ByVal 0&)
    Call CloseHandle(FileHandle)
    
End Sub


