Attribute VB_Name = "modBoot"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Sub SetNumber(NumberText As TextBox, flag As Boolean)
    Dim curstyle As Long
    Dim newstyle As Long
    
    curstyle = GetWindowLong(NumberText.hwnd, GWL_STYLE)
    If flag Then
       curstyle = curstyle Or ES_NUMBER
    Else
       curstyle = curstyle And (Not ES_NUMBER)
    End If
    newstyle = SetWindowLong(NumberText.hwnd, GWL_STYLE, curstyle)
    NumberText.Refresh
End Sub

Public Function Hex2VB(strHexColor As String) As String
  Hex2VB = "&H00" & Right(strHexColor, 2) & Mid(strHexColor, 3, 2) & Left(strHexColor, 2)
End Function
Public Sub SaveConfig()
    Dim i As Integer
    On Error GoTo beep
        '[Application]
        WriteINI "Application", "AutoStart", tAppConfig.bolAutoStart, strFileconfig
        WriteINI "Application", "AlwaysOnTop", tAppConfig.bolOnTop, strFileconfig
        WriteINI "Application", "ShowMenu", tAppConfig.bolMenu, strFileconfig
        WriteINI "Application", "ShowSplash", tAppConfig.bolShowSplash, strFileconfig
        WriteINI "Application", "SystemTray", tAppConfig.bolSysTray, strFileconfig
        WriteINI "Application", "Taskbar", tAppConfig.bolTaskbar, strFileconfig
        WriteINI "Application", "TaskbarScroll", tAppConfig.bolTaskbarScroll, strFileconfig
        WriteINI "Application", "TrayIcon", tAppConfig.intIcon, strFileconfig
        WriteINI "Application", "FileIcon", tAppConfig.intFileIcon, strFileconfig
        WriteINI "Application", "PlaylistIcon", tAppConfig.intPLIcon, strFileconfig
        WriteINI "Application", "Language", CurrentLang, strFileconfig
        
        '[Folder]
        WriteINI "Demension", "LastDir", strLastDir, strFileconfig
        
        '[Skin]
        WriteINI "Skins", "Skin", tCurrentSkin.Infor.Name & ".skn", strFileconfig
        WriteINI "Skins", "EQSlide", tSkinOption.bolEQSlide, strFileconfig
        WriteINI "Skins", "Mini", tCurrentSkin.mini, strFileconfig
        
        '[Device]
        WriteINI "Device", "OutputDevice", tDevice.SoundD.OutputDevice, strFileconfig
        WriteINI "Device", "SampleRate", tDevice.SoundD.Freq, strFileconfig
        WriteINI "Device", "WaveWrite", tDevice.SoundD.WaveWrite, strFileconfig
        WriteINI "Device", "LockVideoRatio", tDevice.VideoD.bolLockRatio, strFileconfig
        WriteINI "Device", "RatioHeight", tDevice.VideoD.intRatioHeight, strFileconfig
        WriteINI "Device", "RatioWidth", tDevice.VideoD.intRatioWidth, strFileconfig
        WriteINI "Device", "DefaultSrceen", tDevice.VideoD.intDefaultScreen, strFileconfig
        WriteINI "Device", "WaveAutoFileName", tDevice.WaveD.bolAutoFilename, strFileconfig
        WriteINI "Device", "WaveOutput", tDevice.WaveD.strWaveOutput, strFileconfig


        '[Player]
        WriteINI "Player", "AutoPlay", tPlayerConfig.bolAutoPlay, strFileconfig
        WriteINI "Player", "Mute", tPlayerConfig.bolMute, strFileconfig
        WriteINI "Player", "Timer", tPlayerConfig.bolTimer, strFileconfig
        WriteINI "Player", "ShowEQ", tPlayerConfig.bolShowEQ, strFileconfig
        WriteINI "Player", "AutoExit", tPlayerConfig.bolAutoExit, strFileconfig
        WriteINI "Player", "AutoRemove", tPlayerConfig.bolAutoRemove, strFileconfig
        WriteINI "Player", "AutoShutdown", tPlayerConfig.bolAutoShutdow, strFileconfig
        WriteINI "Player", "Crossfade", tPlayerConfig.bolCrossfade, strFileconfig
        WriteINI "Player", "CrossfadeTime", tPlayerConfig.intCrossfade, strFileconfig
        WriteINI "Player", "Balance", tPlayerConfig.intBalance, strFileconfig
        WriteINI "Player", "RepeatAll", tPlayerConfig.bolLoop, strFileconfig
        WriteINI "Player", "Repeat", tPlayerConfig.bolRepeat, strFileconfig
        WriteINI "Player", "Shuffe", tPlayerConfig.bolShuffe, strFileconfig
        WriteINI "Player", "Scroll", tPlayerConfig.bolScroll, strFileconfig
        WriteINI "Player", "TimeFast", tPlayerConfig.intTime, strFileconfig
        WriteINI "Player", "Volume", tPlayerConfig.intVolume, strFileconfig
        WriteINI "Player", "FileType", tPlayerConfig.strFileType, strFileconfig
        WriteINI "Player", "ShowList", tPlayerConfig.bolShowList, strFileconfig

        '[Playlist]
        WriteINI "Playlist", "HidePL", tPlaylistConfig.bolHidePL, strFileconfig
        WriteINI "Playlist", "ShowNumber", tPlaylistConfig.bolShowNumber, strFileconfig
        WriteINI "Playlist", "LastFile", tPlaylistConfig.intLastFile, strFileconfig
        WriteINI "Playlist", "RowScroll", tPlaylistConfig.intRowS, strFileconfig
        WriteINI "Playlist", "LoadID", tPlaylistConfig.intLoadID, strFileconfig
        WriteINI "Playlist", "SortBy", tPlaylistConfig.intSortKey, strFileconfig
        WriteINI "Playlist", "SortString", tPlaylistConfig.strSortString, strFileconfig
        WriteINI "Playlist", "DisplayText", tPlaylistConfig.strDisplay, strFileconfig
       
        '[Library]
        WriteINI "Library", "DblClickAction", LibOption.intDblClick, strFileconfig
        WriteINI "Library", "AudioSkip", LibOption.lngAudioSkip, strFileconfig
        WriteINI "Library", "VideoSkip", LibOption.lngVideoSkip, strFileconfig
        WriteINI "Library", "Enable", LibOption.bolUse, strFileconfig
        
        '[Visualization]
        WriteINI "Visualization", "Style", tSkinVis.intStyle, strFileconfig
        WriteINI "Visualization", "Refresh", tSkinVis.intRefresh, strFileconfig
        WriteINI "Visualization", "SpecDraw", tSkinVis.intSpecDraw, strFileconfig
        WriteINI "Visualization", "SpecFill", tSkinVis.intSpecFill, strFileconfig
        WriteINI "Visualization", "SpecPeak", tSkinVis.bolSpecPeak, strFileconfig
        WriteINI "Visualization", "SpecPeakPause", tSkinVis.intSpecPeakPause, strFileconfig
        WriteINI "Visualization", "SpecPeakDrop", tSkinVis.intSpecPeakDrop, strFileconfig
        WriteINI "Visualization", "Osc", tSkinVis.intOsc, strFileconfig
        
        WriteINI "Visualization", "GetStyle", tMainWin.Style, strFileconfig
        WriteINI "Visualization", "Data", tMainWin.Data, strFileconfig
        WriteINI "Visualization", "Plugin", tMainWin.plugin, strFileconfig
        WriteINI "Visualization", "FontColor", tMainWin.FontColor, strFileconfig
        WriteINI "Visualization", "BackColor", tMainWin.BackColor, strFileconfig
        WriteINI "Visualization", "BackGround", tMainWin.BackGround, strFileconfig
        WriteINI "Visualization", "Interval", tMainWin.TimeDisplay, strFileconfig
        WriteINI "Visualization", "UsePictureBG", tMainWin.bolUsePic, strFileconfig
        WriteINI "Visualization", "ShowTitle", tMainWin.bolShowTitle, strFileconfig
        
        WriteINI "WinampPlugins", "Enabled", tWinamp.bolEnabled, strFileconfig
        WriteINI "WinampPlugins", "LastPlugins", tWinamp.intCurrentPlugin, strFileconfig
        WriteINI "WinampPlugins", "LastSubPlugins", tWinamp.intSubPlugin, strFileconfig
        WriteINI "WinampPlugins", "DSPEnabled", tWinampDSP.bolEnabled, strFileconfig
        WriteINI "WinampPlugins", "LastDSPPlugins", tWinampDSP.intCurrentPlugin, strFileconfig
        
beep:
        Exit Sub
End Sub

Sub Main()
    On Error Resume Next
    Dim bolInstance As Boolean
    
    
    strFileconfig = App.path & "\M3P.ini"
    
    If App.PrevInstance Then
        Call SendData(Command)
        End
    End If
        
    InitCommonControls
    XPStyle False
    
    Dim reg As New clsRegistry
    Dim RunCount As Long
    
    reg.ClassKey = HKEY_CURRENT_USER
    reg.SectionKey = "Software\HanaSoft\MP3_proPlayer"
    reg.ValueKey = "RunCount"
    reg.ValueType = REG_DWORD
    If reg.KeyExists = False Then reg.CreateKey
    RunCount = reg.value
    reg.ValueKey = "RunCount"
    reg.value = RunCount + 1
    
    Call reg.CreateEXEAssociation(GetProperPath(App.path) & "MP3_proPlayer.exe", "M3P.Skin", "M3P Skin Zip", "skn", "&Apply this skin for M3P", "/open", False, , , , , , 0)
    
    If Not FileExists(strFileconfig) Then
        Call LoadConfig(True)
        Open strFileconfig For Output As #1
        Close #1
        Call SaveConfig
    Else
        Call LoadConfig(False)
    End If
    
    GenreArray = Split(strGenreMatrix, "|")
    
    strLibrary = App.path & "\M3P.M3PData"
    If Not FileExists(strLibrary) Then
        Open strLibrary For Output As #1
            Print #1, "[M3P_Library]"
            Print #1, "FileLen=0"
        Close #1
    End If
    
    If Dir(App.path & "\Skins\Default", vbDirectory) = "" Then
        Call UnZipSkin(App.path & "\Skins\Default.skn")
    End If
    
    If tAppConfig.bolShowSplash Then
        Load frmSplash
    Else
        bolLoading = True
        Load frmMenu
    End If
    
    DoCommand Command
    
End Sub
Public Sub LoadConfig(Optional Default As Boolean)
    On Error GoTo beep
    Dim i As Integer
    
    If Default = vbNull Or Default = False Then
    '[Application]
        tAppConfig.bolAutoStart = ReadINI("Application", "AutoStart", strFileconfig, False)
        tAppConfig.bolMenu = ReadINI("Application", "ShowMenu", strFileconfig, False)
        tAppConfig.bolOnTop = ReadINI("Application", "AlwaysOnTop", strFileconfig, False)
        tAppConfig.bolShowSplash = ReadINI("Application", "ShowSplash", strFileconfig, True)
        tAppConfig.bolSysTray = ReadINI("Application", "SystemTray", strFileconfig, True)
        tAppConfig.bolTaskbar = ReadINI("Application", "Taskbar", strFileconfig, True)
        tAppConfig.bolTaskbarScroll = ReadINI("Application", "TaskbarScroll", strFileconfig, True)
        tAppConfig.intIcon = ReadINI("Application", "TrayIcon", strFileconfig, 1)
        tAppConfig.intFileIcon = ReadINI("Application", "FileIcon", strFileconfig, 1)
        tAppConfig.intPLIcon = ReadINI("Application", "PlaylistIcon", strFileconfig, 1)
        CurrentLang = ReadINI("Application", "Language", strFileconfig, "English")
        
    '[Folder]
        strLastDir = ReadINI("Demension", "LastDir", strFileconfig, WinAPI.GetWinInfor(eDIR_USER_MY_DOCUMENTS))
    
    '[Device]
        tDevice.SoundD.OutputDevice = ReadINI("Device", "OutputDevice", strFileconfig, 1)
        tDevice.SoundD.Freq = ReadINI("Device", "SampleRate", strFileconfig, 44100)
        tDevice.SoundD.WaveWrite = ReadINI("Device", "WaveWrite", strFileconfig, False)
        tDevice.VideoD.bolLockRatio = ReadINI("Device", "LockVideoRatio", strFileconfig, False)
        tDevice.VideoD.intDefaultScreen = ReadINI("Device", "DefaultScreen", strFileconfig, 1)
        tDevice.VideoD.intRatioHeight = ReadINI("Device", "RatioHeight", strFileconfig, 3)
        tDevice.VideoD.intRatioWidth = ReadINI("Device", "RatioWidth", strFileconfig, 4)
        tDevice.WaveD.bolAutoFilename = ReadINI("Device", "WaveAutoFileName", strFileconfig, False)
        tDevice.WaveD.strWaveOutput = ReadINI("Device", "WaveOutput", strFileconfig, WinAPI.GetWinInfor(eDIR_USER_MY_DOCUMENTS))
        
    '[Player]
        tPlayerConfig.bolAutoPlay = ReadINI("Player", "AutoPlay", strFileconfig, False)
        tPlayerConfig.bolAutoExit = ReadINI("Player", "AutoExit", strFileconfig, False)
        tPlayerConfig.bolAutoRemove = ReadINI("Player", "AutoRemove", strFileconfig, False)
        tPlayerConfig.bolAutoShutdow = ReadINI("Player", "AutoShutdown", strFileconfig, False)
        tPlayerConfig.bolCrossfade = ReadINI("Player", "Crossfade", strFileconfig, False)
        tPlayerConfig.bolMute = ReadINI("Player", "Mute", strFileconfig, False)
        tPlayerConfig.bolLoop = ReadINI("Player", "RepeatAll", strFileconfig, False)
        tPlayerConfig.bolRepeat = ReadINI("Player", "Repeat", strFileconfig, False)
        tPlayerConfig.bolScroll = ReadINI("Player", "Scroll", strFileconfig, True)
        tPlayerConfig.bolShowEQ = ReadINI("Player", "ShowEQ", strFileconfig, True)
        tPlayerConfig.bolShowList = ReadINI("Player", "ShowList", strFileconfig, False)
        tPlayerConfig.bolShuffe = ReadINI("Player", "Shuffe", strFileconfig, False)
        tPlayerConfig.bolTimer = ReadINI("Player", "Timer", strFileconfig, True)
        tPlayerConfig.intBalance = ReadINI("Player", "Balance", strFileconfig, 0)
        tPlayerConfig.intCrossfade = ReadINI("Player", "CrossfadeTime", strFileconfig, 5)
        tPlayerConfig.intTime = ReadINI("Player", "TimeFast", strFileconfig, 10)
        tPlayerConfig.intVolume = ReadINI("Player", "Volume", strFileconfig, 100)
        tPlayerConfig.strFileType = ReadINI("Player", "FileType", strFileconfig, ".MP1,.MP2,.MP3,.OGG")
        
    '[Playlist]
        tPlaylistConfig.bolHidePL = ReadINI("Playlist", "HidePL", strFileconfig, False)
        tPlaylistConfig.bolShowNumber = ReadINI("Playlist", "ShowNumber", strFileconfig, True)
        tPlaylistConfig.intLoadID = ReadINI("Playlist", "LoadID", strFileconfig, 0)
        tPlaylistConfig.intLastFile = ReadINI("Playlist", "LastFile", strFileconfig, 1)
        tPlaylistConfig.intRowS = ReadINI("Playlist", "RowScroll", strFileconfig, 1)
        tPlaylistConfig.intSortKey = ReadINI("Playlist", "SortBy", strFileconfig)
        tPlaylistConfig.strSortString = ReadINI("Playlist", "SortString", strFileconfig)
        tPlaylistConfig.strDisplay = ReadINI("Playlist", "DisplayText", strFileconfig, "%1 - %2")
        
    '[Skin]
        tSkinOption.SkinDir = App.path & "\Skins"
        tCurrentSkin.Infor.Name = Mid(ReadINI("Skins", "Skin", strFileconfig, "Default"), 1, Len(ReadINI("Skins", "Skin", strFileconfig)) - 4)
        tCurrentSkin.mini = ReadINI("Skins", "Mini", strFileconfig, False)
        tSkinOption.bolEQSlide = ReadINI("Skins", "EQSlide", strFileconfig, False)
        
    '[Library]
        LibOption.intDblClick = CInt(ReadINI("Library", "DblClickAction", strFileconfig, 0))
        LibOption.lngAudioSkip = CLng(ReadINI("Library", "AudioSkip", strFileconfig, 500))
        LibOption.lngVideoSkip = CLng(ReadINI("Library", "VideoSkip", strFileconfig, 1000))
        LibOption.bolUse = ReadINI("Library", "Enable", strFileconfig, True)
    '[DirectX 8 Effect]
        bolEQEnabled = ReadINI("Equalizer", "EqualEnabled", strFileconfig, True)
        strCurrentEQPreset = ReadINI("Equalizer", "LastPreset", strFileconfig, "")
        For i = 0 To 9
               intEQ(i) = ReadINI("Equalizer", "Equa_" & i, strFileconfig, 0)
        Next i
        'intAmp = ReadINI("Equalizer", "Ampli", strFileconfig)

        bolUseDirectX = ReadINI("DSP Effect", "UseDirectX", strFileconfig, False)
        bolDSP(0) = ReadINI("DSP Effect", "Chorus", strFileconfig, False)
        bolDSP(1) = ReadINI("DSP Effect", "Compressor", strFileconfig, False)
        bolDSP(2) = ReadINI("DSP Effect", "Distortion", strFileconfig, False)
        bolDSP(3) = ReadINI("DSP Effect", "Echo", strFileconfig, False)
        bolDSP(4) = ReadINI("DSP Effect", "Flanger", strFileconfig, False)
        bolDSP(5) = ReadINI("DSP Effect", "Gargle", strFileconfig, False)
        bolDSP(6) = ReadINI("DSP Effect", "I3DL2Reverb", strFileconfig, False)
        bolDSP(7) = ReadINI("DSP Effect", "Reverb", strFileconfig, False)
        For i = 0 To 6
            intChorus(i) = ReadINI("DSP Effect", "Chorus_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 5
            intCompressor(i) = ReadINI("DSP Effect", "Compressor_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 4
            intDistortion(i) = ReadINI("DSP Effect", "Distortion_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 4
            intEcho(i) = ReadINI("DSP Effect", "Echo_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 6
            intFlanger(i) = ReadINI("DSP Effect", "Flanger_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 1
            intGargle(i) = ReadINI("DSP Effect", "Gargle_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 11
            intI3DL2Reverb(i) = ReadINI("DSP Effect", "I3DL2Reverb_" & i, strFileconfig, 0)
        Next i
        For i = 0 To 3
            intReverb(i) = ReadINI("DSP Effect", "Reverb_" & i, strFileconfig, 0)
        Next i
   
    '[Visualization]
        tSkinVis.intStyle = ReadINI("Visualization", "Style", strFileconfig, 0)
        tSkinVis.intRefresh = ReadINI("Visualization", "Refresh", strFileconfig, 0)
        tSkinVis.intSpecDraw = ReadINI("Visualization", "SpecDraw", strFileconfig, 0)
        tSkinVis.intSpecFill = ReadINI("Visualization", "SpecFill", strFileconfig, 0)
        tSkinVis.bolSpecPeak = ReadINI("Visualization", "SpecPeak", strFileconfig, True)
        tSkinVis.intSpecPeakPause = ReadINI("Visualization", "SpecPeakPause", strFileconfig, 10)
        tSkinVis.intSpecPeakDrop = ReadINI("Visualization", "SpecPeakDrop", strFileconfig, 3)
        tSkinVis.intOsc = ReadINI("Visualization", "Osc", strFileconfig, 0)
        
        tMainWin.Style = ReadINI("Visualization", "GetStyle", strFileconfig)
        tMainWin.Data = ReadINI("Visualization", "Data", strFileconfig)
        tMainWin.plugin = ReadINI("Visualization", "Plugin", strFileconfig)
        tMainWin.BackColor = CLng(ReadINI("Visualization", "BackColor", strFileconfig, 0))
        tMainWin.FontColor = CLng(ReadINI("Visualization", "FontColor", strFileconfig, &HFFFFFF))
        tMainWin.BackGround = (ReadINI("Visualization", "BackGround", strFileconfig, ""))
        tMainWin.TimeDisplay = ReadINI("Visualization", "Inteval", strFileconfig, 10)
        tMainWin.bolShowTitle = ReadINI("Visualization", "ShowTitle", strFileconfig, True)
        tMainWin.bolUsePic = ReadINI("Visualization", "UsePictureBG", strFileconfig, False)
   '[Winamp Vis]
        tWinamp.bolEnabled = ReadINI("WinampPlugins", "Enabled", strFileconfig, False)
        tWinamp.intSubPlugin = ReadINI("WinampPlugins", "LastSubPlugins", strFileconfig, 0)
        tWinamp.intCurrentPlugin = ReadINI("WinampPlugins", "LastPlugins", strFileconfig, 0)
        tWinampDSP.bolEnabled = ReadINI("WinampPlugins", "DSPEnabled", strFileconfig, False)
        tWinampDSP.intCurrentPlugin = ReadINI("WinampPlugins", "LastDSPPlugins", strFileconfig, 0)
    Else ' Load Default Value
    '[Demension]
        WriteINI "Demension", "VisualHeight", 400, strFileconfig
        WriteINI "Demension", "VisualWidth", 400, strFileconfig
        WriteINI "Demension", "VisualLeft", 400, strFileconfig
        WriteINI "Demension", "VisualTop", 400, strFileconfig
    '[Application]
        tAppConfig.bolAutoStart = False
        tAppConfig.bolMenu = False
        tAppConfig.bolOnTop = False
        tAppConfig.bolShowSplash = True
        tAppConfig.bolSysTray = False
        tAppConfig.bolTaskbar = True
        tAppConfig.bolTaskbarScroll = True
        tAppConfig.intIcon = 1
        tAppConfig.intFileIcon = 1
        tAppConfig.intPLIcon = 1
        CurrentLang = "English"
    '[Folder]
        strLastDir = WinAPI.GetWinInfor(eDIR_USER_MY_DOCUMENTS)
    '[Device]
        tDevice.SoundD.OutputDevice = 1
        tDevice.SoundD.Freq = 44100
        tDevice.SoundD.WaveWrite = False
        tDevice.VideoD.bolLockRatio = False
        tDevice.VideoD.intDefaultScreen = 1
        tDevice.VideoD.intRatioHeight = 3
        tDevice.VideoD.intRatioWidth = 4
        tDevice.WaveD.bolAutoFilename = False
        tDevice.WaveD.strWaveOutput = WinAPI.GetWinInfor(eDIR_USER_MY_DOCUMENTS)
    '[Player]
        tPlayerConfig.bolAutoPlay = False
        tPlayerConfig.bolAutoExit = False
        tPlayerConfig.bolAutoRemove = False
        tPlayerConfig.bolAutoShutdow = False
        tPlayerConfig.bolCrossfade = False
        tPlayerConfig.bolMute = False
        tPlayerConfig.bolLoop = False
        tPlayerConfig.bolRepeat = False
        tPlayerConfig.bolScroll = True
        tPlayerConfig.bolShowEQ = True
        tPlayerConfig.bolShowList = False
        tPlayerConfig.bolShuffe = False
        tPlayerConfig.bolTimer = True
        tPlayerConfig.intBalance = 0
        tPlayerConfig.intCrossfade = 0
        tPlayerConfig.intTime = 10
        tPlayerConfig.intVolume = 100
        tPlayerConfig.strFileType = ".MP1,.MP2,.MP3,.OGG"
    '[Playlist]
        tPlaylistConfig.bolHidePL = False
        tPlaylistConfig.bolShowNumber = True
        tPlaylistConfig.intLoadID = 0
        tPlaylistConfig.intLastFile = 0
        tPlaylistConfig.intRowS = 10
        tPlaylistConfig.intSortKey = 2
        tPlaylistConfig.strSortString = "Filename"
        tPlaylistConfig.strDisplay = "%1 - %2"
    '[Skin]
        tSkinOption.SkinDir = App.path & "\Skins"
        tCurrentSkin.Infor.Name = "Default"
        tCurrentSkin.mini = False
        tSkinOption.bolEQSlide = False
    '[Library]
        LibOption.intDblClick = 0
        LibOption.lngAudioSkip = 500
        LibOption.lngVideoSkip = 1000
        LibOption.bolUse = False
        For i = 0 To 14
            LibOption.ColView(i) = True
            LibOption.ColWidth(i) = 100
        Next i
    '[DirectX 8 Effect]
        bolEQEnabled = False
        strCurrentEQPreset = ""
        For i = 0 To 9
            intEQ(i) = 0
        Next i
        'intAmp = ReadINI("Equalizer", "Ampli", strFileconfig)
        bolUseDirectX = False
        For i = 0 To 7
            bolDSP(i) = False
        Next i
    '[Visualization]
        tSkinVis.intStyle = 1
        tSkinVis.intOsc = 0
        tSkinVis.bolSpecPeak = True
        tSkinVis.intRefresh = 75
        tSkinVis.intSpecDraw = 0
        tSkinVis.intSpecFill = 0
        tSkinVis.intSpecPeakDrop = 4
        tSkinVis.intSpecPeakPause = 5
        tMainWin.Style = 0
        tMainWin.Data = 3
        tMainWin.plugin = ""
        tMainWin.BackColor = &H0
        tMainWin.FontColor = &HFFFFFF
        tMainWin.BackGround = ""
        tMainWin.TimeDisplay = 25
        tMainWin.bolShowTitle = False
        tMainWin.bolUsePic = False
   '[Winamp Vis]
        tWinamp.bolEnabled = False
        tWinamp.intSubPlugin = 0
        tWinamp.intCurrentPlugin = 0
        tWinampDSP.bolEnabled = False
        tWinampDSP.intCurrentPlugin = 0
    End If
    Exit Sub
beep:
    If Err.Number <> 0 Then
        LoadConfig True
        Exit Sub
    End If
End Sub

Public Function GetFullName(strFile As String) As String
    Dim ret As Long
    Dim buffer As String
    Dim tmp As String
    buffer = String(4096, 0)
    ret = GetFullPathName(strFile, 4096, buffer, tmp)
    GetFullName = Left(buffer, ret)
End Function
Public Sub Wait(milliseconds As Long)
    DoEvents
    Sleep milliseconds
    DoEvents
End Sub

Public Function GetProperPath(ByVal fullPath As String) As String
    GetProperPath = IIf(Mid(fullPath, Len(fullPath), 1) = "\", fullPath, fullPath & "\")
End Function
Public Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Dir(fullPath) <> "")
End Function

Public Function ReadINI(strSection As String, strKey As String, strFileINI As String, Optional strDefault As Variant) As String
    On Error GoTo beep
    Dim StrTemp As String * 255
    GetPrivateProfileString strSection, strKey, vbNull, StrTemp, Len(StrTemp), strFileINI
    ReadINI = Mid(StrTemp, 1, InStr(1, StrTemp, vbNullChar) - 1)
    Exit Function
beep:
    ReadINI = strDefault
End Function
Public Function WriteINI(strSection As String, strKey As String, KeyValue As Variant, strFileINI As String) As Boolean
    Dim ret As Long
    ret = WritePrivateProfileString(strSection, strKey, CStr(KeyValue), strFileINI)
    If ret = 0 Then
        WriteINI = True
    Else
        WriteINI = False
    End If
End Function


Public Sub AlwaysOnTop(frm As Form, bolOnTop As Boolean)
    Dim lFlag
    If bolOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos frm.hwnd, lFlag, frm.Left / Screen.TwipsPerPixelX, frm.Top / Screen.TwipsPerPixelY, frm.width / Screen.TwipsPerPixelX, frm.height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Public Sub DoCommand(strData As String)
    On Error Resume Next
    Dim sw As String
    Dim Com As String
    Dim FileOpen As String
    
    Com = Trim(strData)
    If InStr(1, Com, " ") <> 0 Then
      sw = Left(Com, InStr(1, Com, " ") - 1)
    Else
      sw = Com
    End If
    Select Case sw
        Case "/play"
            frmMedia.btnPlayClick
        Case "/pause"
            frmMedia.btnPauseClick
        Case "/stop"
            StopPlayer
        Case "/back"
            BackTrack
        Case "/next"
            NextTrack
        Case "/exit"
            StopPlayer
            Unload frmMedia
        Case "/top"
            GoTop
        Case "/bottom"
            GoBottom
        Case "/shuffe"
            frmMedia.btnShuffeClick
        Case "/repeat"
            frmMedia.btnRepeatClick
        Case "/showeq"
            frmMedia.btnShowEQClick
        Case "/showpl"
            frmMedia.btnShowPLClick
        Case "/open"
            FileOpen = Right(Com, Len(Com) - InStr(1, Com, " "))
            Select Case LCase(Right(FileOpen, InStrRev(FileOpen, ".", -1, vbBinaryCompare)))
                Case "m3u"
                    Call LoadPlaylistM3U(FileOpen)
                    Play (1)
                Case "pls"
                    Call LoadPlaylistPLS(FileOpen)
                    Play (1)
                Case "skn"
                    Call LoadSkin(Mid(GetShortName(FileOpen), 1, Len(GetShortName(FileOpen - 1)) - 4), tCurrentSkin.mini)
                Case Else
                    Call AddFile(FileOpen)
                    With frmPlayList
                        .List.DisplayList
                    End With
            End Select
        Case "/add"
            FileOpen = Right(Com, Len(Com) - InStr(1, Com, " "))
            AddSingleFolder (FileOpen)
        Case "/vold"
            If tPlayerConfig.intVolume > 0 Then
                Call frmMedia.volume(tPlayerConfig.intVolume - 10)
            End If
        Case "/volu"
            If tPlayerConfig.intVolume < 100 Then
                Call frmMedia.volume(tPlayerConfig.intVolume + 10)
            End If
    End Select
End Sub



