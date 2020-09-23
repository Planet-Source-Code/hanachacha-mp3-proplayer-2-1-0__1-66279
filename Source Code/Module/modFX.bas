Attribute VB_Name = "modFX"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" _
        Alias "FindFirstFileA" (ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindNextFile Lib "kernel32" _
        Alias "FindNextFileA" (ByVal hFindFile As Long, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindClose Lib "kernel32" (ByVal _
        hFindFile As Long) As Long

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Const MAX_PATH = 259

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public EQ(9) As Long        '10 equalizer band
Public FX(7) As Long
Dim i As Integer
Public DSPPlugins() As String
Public Function LoadDsp(file As String) As Boolean
    Dim plugin As Long
    Dim module As Long
    Dim DSPTemp() As String
    Dim chan As Long
    Dim i As Integer
    
    
    'load winamp plugin from the plugin dir
    plugin = BASS_WADSP_Load(App.path & "\plugins\" & file, 5, 5, 100, 100, 0)
    If plugin = 0 Then
        LoadDsp = False
        Exit Function
    End If
    chan = frmMedia.Player.handle
    Call BASS_WADSP_Start(plugin, 1, chan) 'start plugin
    BASS_WADSP_ChannelSetDSP plugin, chan, 1 'and set channel
    LoadDsp = True
    ReDim DSPTemp(UBound(DSPPlugins))
    'copy DSPPlugins() to local var
    For i = 0 To UBound(DSPPlugins)
        DSPTemp(i) = DSPPlugins(i)
    Next
    ReDim DSPPlugins(plugin) 'allocate more mem for the new plugins
    'copy var back
    For i = 0 To UBound(DSPTemp)
        DSPPlugins(i) = DSPTemp(i)
    Next
    DSPPlugins(plugin) = file 'add new plugin to var
End Function


Public Function UnloadDsp(file As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(DSPPlugins) 'search for plugin
        If DSPPlugins(i) = file Then 'found...
            BASS_WADSP_Stop CLng(i) 'stop and
            BASS_WADSP_FreeDSP CLng(i) 'free it.
            DSPPlugins(i) = "" 'reset var
            UnloadDsp = True 'return true
            Exit Function
        End If
    Next
    UnloadDsp = False 'return false
End Function


Public Function DspIsLoad(file As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(DSPPlugins) ' search for plugins
        If DSPPlugins(i) = file Then 'found
            DspIsLoad = True 'return true
            Exit Function
        End If
    Next
    DspIsLoad = False 'return false
End Function


Public Sub DspOpenConfig(file As String)
    Dim i As Integer
    For i = 0 To UBound(DSPPlugins) 'search for plugin
        If DSPPlugins(i) = file Then 'found
            BASS_WADSP_Config CLng(i) 'open config
        End If
    Next
End Sub

Public Sub RestartDsp()
    Dim i As Integer
    Dim chan As Long
    
    chan = frmMedia.Player.handle
    For i = 0 To UBound(DSPPlugins) 'search for plugin
        If Len(DSPPlugins(i)) > 0 Then 'found
            BASS_WADSP_ChannelSetDSP CLng(i), chan, 1
        End If
    Next
End Sub


Public Sub SetFX(lngHandle As Long)
    If bolDSP(0) = True Then
        FX(0) = BASS_ChannelSetFX(lngHandle, BASS_FX_CHORUS, 1)
    End If
    If bolDSP(1) = True Then
        FX(1) = BASS_ChannelSetFX(lngHandle, BASS_FX_COMPRESSOR, 1)
    End If
    If bolDSP(2) = True Then
        FX(2) = BASS_ChannelSetFX(lngHandle, BASS_FX_DISTORTION, 1)
    End If
    If bolDSP(3) = True Then
        FX(3) = BASS_ChannelSetFX(lngHandle, BASS_FX_ECHO, 1)
    End If
    If bolDSP(4) = True Then
        FX(4) = BASS_ChannelSetFX(lngHandle, BASS_FX_FLANGER, 1)
    End If
    If bolDSP(5) = True Then
        FX(5) = BASS_ChannelSetFX(lngHandle, BASS_FX_GARGLE, 1)
    End If
    If bolDSP(6) = True Then
        FX(6) = BASS_ChannelSetFX(lngHandle, BASS_FX_I3DL2REVERB, 1)
    End If
    If bolDSP(7) = True Then
        FX(7) = BASS_ChannelSetFX(lngHandle, BASS_FX_REVERB, 1)
    End If
End Sub
Public Sub UpdateFX(FXhandle As Long)
    With frmDSP
        Select Case FXhandle
            Case 0
                If bolDSP(0) Then
                    Dim p As BASS_FXCHORUS
                    Call BASS_FXGetParameters(FX(0), p)
                    p.fWetDryMix = .prgChorus(0).value
                    p.fDepth = .prgChorus(1).value
                    p.fFeedback = .prgChorus(2).value
                    p.fFrequency = .prgChorus(3).value
                    p.fDelay = .prgChorus(4).value
                    p.lWaveform = .cboChorus(0).ListIndex
                    p.lPhase = .cboChorus(1).ListIndex
                    Call BASS_FXSetParameters(FX(0), p)
                End If
            Case 1
                If bolDSP(1) Then
                    Dim p1 As BASS_FXCOMPRESSOR
                    Call BASS_FXGetParameters(FX(1), p1)
                    p1.fAttack = .prgCompressor(1).value / 100
                    p1.fGain = .prgCompressor(0).value
                    p1.fPredelay = .prgCompressor(5).value
                    p1.fRatio = .prgCompressor(4).value
                    p1.fRelease = .prgCompressor(2).value
                    p1.fThreshold = .prgCompressor(3).value
                    Call BASS_FXSetParameters(FX(1), p1)
                End If
            Case 2
                If bolDSP(2) Then
                    Dim p2 As BASS_FXDISTORTION
                    Call BASS_FXGetParameters(FX(2), p2)
                    p2.fEdge = .prgDistortion(1).value
                    p2.fGain = .prgDistortion(0).value
                    p2.fPostEQBandwidth = .prgDistortion(3).value
                    p2.fPostEQCenterFrequency = .prgDistortion(2).value
                    p2.fPreLowpassCutoff = .prgDistortion(4).value
                    Call BASS_FXSetParameters(FX(2), p2)
                End If
            Case 3
                If bolDSP(3) Then
                    Dim p3 As BASS_FXECHO
                    Call BASS_FXGetParameters(FX(3), p3)
                    p3.fFeedback = .prgEcho(1).value
                    p3.fLeftDelay = .prgEcho(2).value
                    p3.fRightDelay = .prgEcho(3).value
                    p3.fWetDryMix = .prgEcho(0).value
                    p3.lPanDelay = CBool(.chkEcho.value)
                    Call BASS_FXSetParameters(FX(3), p3)
                End If
            Case 4
                If bolDSP(4) Then
                    Dim p4 As BASS_FXFLANGER
                    Call BASS_FXGetParameters(FX(4), p4)
                    p4.fDelay = .prgFlanger(4).value
                    p4.fDepth = .prgFlanger(1).value
                    p4.fFeedback = .prgFlanger(2).value
                    p4.fFrequency = .prgFlanger(3).value
                    p4.fWetDryMix = .prgFlanger(0).value
                    p4.lPhase = .cboFlanger(1).ListIndex
                    p4.lWaveform = .cboFlanger(0).ListIndex
                    Call BASS_FXSetParameters(FX(4), p4)
                End If
            Case 5
                If bolDSP(5) Then
                    Dim p5 As BASS_FXGARGLE
                    Call BASS_FXGetParameters(FX(5), p5)
                    p5.dwRateHz = .prgGargle.value
                    p5.dwWaveShape = .cboGargle.ListIndex
                    Call BASS_FXSetParameters(FX(5), p5)
                End If
            Case 6
                If bolDSP(6) Then
                    Dim p6 As BASS_FXI3DL2REVERB
                    Call BASS_FXGetParameters(FX(6), p6)
                    p6.flDecayHFRatio = .prgLReverb(4).value / 100
                    p6.flDecayTime = .prgLReverb(3).value / 100
                    p6.flDensity = .prgLReverb(10).value
                    p6.flDiffusion = .prgLReverb(9).value
                    p6.flHFReference = .prgLReverb(11).value
                    p6.flReflectionsDelay = .prgLReverb(6).value / 1000
                    p6.flReverbDelay = .prgLReverb(8).value / 1000
                    p6.flRoomRolloffFactor = .prgLReverb(2).value
                    p6.lReflections = .prgLReverb(5).value
                    p6.lReverb = .prgLReverb(7).value
                    p6.lRoom = .prgLReverb(0).value
                    p6.lRoomHF = .prgLReverb(0).value
                    Call BASS_FXSetParameters(FX(6), p6)
                End If
            Case 7
                If bolDSP(7) Then
                    Dim p7 As BASS_FXREVERB
                    Call BASS_FXGetParameters(FX(7), p7)
                    p7.fInGain = .prgReverb(0).value
                    p7.fReverbMix = .prgReverb(1).value
                    p7.fReverbTime = .prgReverb(2).value
                    p7.fHighFreqRTRatio = .prgReverb(3).value / 1000
                    Call BASS_FXSetParameters(FX(7), p7)
                End If
        End Select
    End With
End Sub
Public Sub SetEqual(lngHandle As Long)
    Dim p As BASS_FXPARAMEQ
    Dim i As Integer
    EQ(0) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(1) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(2) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(3) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(4) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(5) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(6) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(7) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(8) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    EQ(9) = BASS_ChannelSetFX(lngHandle, BASS_FX_PARAMEQ, 0)
    p.fGain = 0
    p.fBandwidth = 12
    'setup equalizer
    p.fCenter = 80 '[80hz]
    Call BASS_FXSetParameters(EQ(0), p)
        
    p.fCenter = 180 '[180hz]
    Call BASS_FXSetParameters(EQ(1), p)
        
    p.fCenter = 340 '[340hz]
    Call BASS_FXSetParameters(EQ(2), p)
             
    p.fCenter = 650 '[650hz]
    Call BASS_FXSetParameters(EQ(3), p)
        
    p.fCenter = 1000 '[1khz]
    Call BASS_FXSetParameters(EQ(4), p)
        
    p.fCenter = 3000 '[2khz]
    Call BASS_FXSetParameters(EQ(5), p)
            
    p.fCenter = 6000 '[6khz]
    Call BASS_FXSetParameters(EQ(6), p)
        
    p.fCenter = 12000 '[12khz]
    Call BASS_FXSetParameters(EQ(7), p)
        
    p.fCenter = 14000 '[14khz]
    Call BASS_FXSetParameters(EQ(8), p)
   
    p.fCenter = 16000 '[16khz]
    Call BASS_FXSetParameters(EQ(9), p)
            
End Sub
Public Sub UpdateEqual(ByVal IndexBand As Long, ByVal intGain As Long)
    If bolEQEnabled = True Then
        Dim p As BASS_FXPARAMEQ
        Call BASS_FXGetParameters(EQ(IndexBand), p)
        p.fGain = intGain
        p.fBandwidth = 12
        Call BASS_FXSetParameters(EQ(IndexBand), p)
    End If
End Sub
Public Sub SaveEQ(strFile As String)
    On Error Resume Next
    Dim intEQpreset(9) As Integer
        Call WriteINI("MP3_ProPlayer", "EqualizerVerSion", "2.0.0", strFile)
        Call WriteINI("Equalizer", "Name", Mid(GetShortName(strFile), 1, Len(GetShortName(strFile)) - 4), strFile)
        With frmMedia
            For i = 0 To 9
                intEQpreset(i) = .sldEqua(i).value
                Call WriteINI("Equalizer", "Equa_" & i, intEQpreset(i), strFile)
            Next i
        End With
End Sub
Public Sub SaveEQPreset(strPreset As String)
    On Error Resume Next
        Dim strEqualPreset As String
        Dim intEQpreset(9) As Integer
        strEqualPreset = App.path & "\EQ\EqualizerPreset.epr"
        With frmMedia
            For i = 0 To 9
                intEQpreset(i) = .sldEqua(i).value
                Call WriteINI(strPreset, "Equa_" & i, intEQpreset(i), strEqualPreset)
            Next i
        End With
End Sub
Public Sub LoadEQ(strFileEQ As String)
    If FileExists(strFileEQ) Then
        Dim strTest As String
        strTest = ReadINI("MP3_ProPlayer", "EqualizerVerSion", strFileEQ)
        If strTest = "2.0.0" Then
            With frmMedia
                For i = 0 To 9
                    .sldEqua(i).value = ReadINI("Equalizer", "Equa_" & i, strFileEQ)
                    .sldMiniEqua(i).value = ReadINI("Equalizer", "Equa_" & i, strFileEQ)
                Next i
                If bolEQEnabled = True Then
                    For i = 0 To 9
                        Call UpdateEqual(i, .sldEqua(i).value)
                    Next i
                End If
            End With
            strCurrentEQPreset = ""
        Else
            MsgBox "This file isn't MP3_proPlayer 2.0.0 Equalizer file", vbCritical, "MP3_proPlayer"
        End If
    End If
End Sub
Public Sub LoadEQPreset(strPreset As String)
    On Error Resume Next
    Dim strEqualPreset As String
    
    strEqualPreset = App.path & "\EQ\EqualizerPreset.epr"
    If FileExists(strEqualPreset) Then
        With frmMedia
            For i = 0 To 9
                .sldEqua(i).value = ReadINI(strPreset, "Equa_" & i, strEqualPreset)
                .sldMiniEqua(i).value = ReadINI(strPreset, "Equa_" & i, strEqualPreset)
            Next i
            If bolEQEnabled = True Then
                For i = 0 To 9
                    Call UpdateEqual(i, .sldEqua(i).value)
                Next i
            End If
            strCurrentEQPreset = strPreset
            WriteINI "Equalizer", "LastPreset", strCurrentEQPreset, strFileconfig
        End With
    Else
        MsgBox "MP3_proPlayer 2.0.0 Equalizer preset 'EqualizerPreset.epr'file coudn 't found", vbCritical, "MP3_proPlayer"
    End If
End Sub
'It crash my computer so you don't use it
Public Sub PreAmp(ByVal handle As Long, ByVal channel As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
    Dim d() As Single, a As Long
    ReDim d(length / 4) As Single
    
    Call CopyMemory(d(0), ByVal buffer, length)
    
    For a = 0 To (length / 4) - 1 Step 2
        d(a) = d(a) * lngPreAmpVal
        d(a + 1) = d(a + 1) * lngPreAmpVal
    Next a
    
    Call CopyMemory(ByVal buffer, d(0), length)
End Sub

Public Sub SetupDSPPlugins()
    Dim file$, hFile&, FD As WIN32_FIND_DATA, Patt$, Root$
    Dim plugin As Long
    
    Patt = "dsp_*.dll"
    Root = App.path & "\Plugins\"
    'read plugin dir and add to list
    hFile = FindFirstFile(Root & Patt, FD)
    If hFile = 0 Or hFile = -1 Then Exit Sub
    With frmOption
        Do
           file = Left(FD.cFileName, InStr(FD.cFileName, Chr(0)) - 1)
           If Not (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
             = FILE_ATTRIBUTE_DIRECTORY Then
             If (file <> ".") And (file <> "..") Then
               plugin = BASS_WADSP_Load(App.path & "\Plugins\" & file, 5, 5, 100, 100, CLng(0))
               If plugin = 0 Then
                    .lvwDSP.ListItems.Add , , App.path & "\Plugins\" & file & " - Invalid Plugin"
                    .lvwDSP.ListItems(.lvwDSP.ListItems.Count).SubItems(1) = "False"
               Else
                    
                    .lvwDSP.ListItems.Add , , VBStrFromAnsiPtr(BASS_WADSP_GetName(plugin))
                    .lvwDSP.ListItems(.lvwDSP.ListItems.Count).Key = file
                    .lvwDSP.ListItems(.lvwDSP.ListItems.Count).SubItems(1) = IIf(DspIsLoad(.lvwDSP.ListItems(.lvwDSP.ListItems.Count).Key), "True", "False")
                    
                    BASS_WADSP_FreeDSP plugin
               End If
             End If
           End If
        Loop While FindNextFile(hFile, FD)
        Call FindClose(hFile)
    End With
End Sub
Public Sub StartDSP()
    With frmOption
        If (.lvwDSP.ListItems.Count = 0) Or (tWinampDSP.intCurrentPlugin = 0) Or (.lvwDSP.ListItems.Count < tWinampDSP.intCurrentPlugin) Then Exit Sub
        If Len(.lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Key) > 0 Then
            If DspIsLoad(.lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Key) Then Exit Sub
            If LoadDsp(.lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Key) Then
                .lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).SubItems(1) = "True"
            Else
                .lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).SubItems(1) = "Error"
            End If
        End If
    End With
End Sub
Public Sub StopDSP()
    With frmOption
        If (.lvwDSP.ListItems.Count = 0) Or (tWinampDSP.intCurrentPlugin = 0) Or (.lvwDSP.ListItems.Count < tWinampDSP.intCurrentPlugin) Then Exit Sub
        If Len(.lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Key) > 0 Then
            If Not DspIsLoad(.lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Key) Then Exit Sub
            If UnloadDsp(.lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).Key) Then
                .lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).SubItems(1) = "False"
            Else
                .lvwDSP.ListItems(tWinampDSP.intCurrentPlugin).SubItems(1) = "Error"
            End If
        End If
    End With
End Sub
Public Sub InitGenPlugin()
    Dim str As String
    
    str = App.path & "\Plugins"
    With frmOption
        .File1.Pattern = "*.dll"
        .File1.path = str
        For i = 0 To frmOption.File1.ListCount - 1
            .File1.ListIndex = i
            If Mid(.File1.List(i), 1, 7) = "M3P_gen" Then
                .lstGenPlugins.AddItem .File1.Filename
            End If
        Next i
        ReDim genPlugin(.lstGenPlugins.ListCount)
    
        On Error GoTo handle
        Dim strGen As String
        Dim strClass As String
        
        For i = 0 To UBound(genPlugin) - 1
            If FileExists(App.path & "\Plugins\" & .lstGenPlugins.List(i)) Then
                strGen = Mid(.lstGenPlugins.List(i), 1, Len(.lstGenPlugins.List(i)) - 4)
                strClass = Mid(.lstGenPlugins.List(i), 5, Len(.lstGenPlugins.List(i)) - 8)
                Set genPlugin(i) = CreateObject(strGen & "." & strClass)
                genPlugin(i).Run
            End If
        Next i
    End With
handle:
    If Err.Number <> 0 Then
        For i = 0 To UBound(genPlugin) - 1
            If ObjPtr(genPlugin(i)) > 0 Then
                Set genPlugin(i) = Nothing
            End If
        Next i
    End If
End Sub

Public Sub InitLang()
    Dim str As String
    
    str = App.path & "\Lang"
    With frmOption
        .File1.Pattern = "*.lng"
        .File1.path = str
        For i = 0 To frmOption.File1.ListCount - 1
            .File1.ListIndex = i
            .cboLang.AddItem Mid(.File1.Filename, 1, Len(.File1.Filename) - 4)
        Next i
    End With
End Sub

