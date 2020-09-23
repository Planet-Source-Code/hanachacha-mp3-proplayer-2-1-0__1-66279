VERSION 5.00
Begin VB.UserControl Player 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "Player.ctx":0000
   Begin VB.Timer tmrState 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   2400
   End
End
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type WAVEHEADER_RIFF  '12 bytes
    RIFF As Long                '"RIFF" = &H46464952
    riffBlockSize As Long       'pos + 44 - 8
    riffBlockType As Long       '"WAVE" = &H45564157
End Type

Private Type WAVEHEADER_data  '8 bytes
   dataBlockType As Long        '"data" = &H61746164
   dataBlockSize As Long        'pos
End Type

Private Type WAVEFORMAT     '24 bytes
    wfBlockType As Long         '"fmt " = &H20746D66
    wfBlockSize As Long
    '--- block size begins from here = 16 bytes
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type



Dim wr As WAVEHEADER_RIFF
Dim wf As WAVEFORMAT
Dim wd As WAVEHEADER_data

Dim lngPlayState As Long
Dim bolAutoPlay As Boolean
Dim bolRecordAble As Boolean

Dim dblVol As Double
Dim lngFreq As Long
Dim filename As String
Dim Stream As Long
Dim fn As Long
Dim pos As Long
Dim buf() As Byte
Dim info As BASS_CHANNELINFO

Event PlayStateChange()
Event EndOfStream()
'Dim InputWin As New Collection
Public Property Get PlayState() As Long
    PlayState = lngPlayState
End Property

Private Sub tmrState_Timer()
    If BASS_ChannelGetLength(Stream) = BASS_ChannelGetPosition(Stream) Then
        RaiseEvent EndOfStream
    End If
End Sub

Private Sub UserControl_Initialize()
    ChDrive App.path
    ChDir App.path & "\"
    'Check Dll files is exists
    If Not FileExists(GetProperPath(App.path & "\") & "bass.dll") Then
        Call MsgBox("BASS.DLL does not exists _ Please reinstall MP3_proPlayer", vbCritical, "MP3_proPlayer")
    End If
    'Check that BASS current version was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("Mp3_proPlayer use Bass.dll version 2.3", vbCritical, "MP3_proPlayer")
    End If
    
    'Now load all addon plugins
    Dim fh As String
    fh = Dir("bass*.dll")   ' find 1st file

    Do While (fh <> "")
        Dim plug As Long
        plug = BASS_PluginLoad(fh, 0)   ' plugin loaded...
        fh = Dir()  ' get next file
    Loop
    
    UserControl.Height = 255
    UserControl.Width = 255
End Sub

Public Sub StopStream()
    On Error Resume Next
    BASS_ChannelStop Stream
    BASS_ChannelSetPosition Stream, 0
    lngPlayState = 0 'Stopped
    tmrState.Enabled = False
    RaiseEvent PlayStateChange
End Sub

Public Sub PlayStream()
    On Error Resume Next
    BASS_ChannelPlay Stream, BASSFALSE
    lngPlayState = 1 'Playing
    tmrState.Enabled = True
    RaiseEvent PlayStateChange
End Sub

Public Sub OpenFile(strMedia As String, Decode As Boolean)
    On Error Resume Next
    Dim StreamHandle As Long
    ChDrive App.path
    ChDir App.path & "\"
    filename = strMedia
    lngPlayState = 0
    RaiseEvent PlayStateChange
    Call BASS_StreamFree(Stream)
    Call BASS_MusicFree(Stream)
    If Not Decode Then
        StreamHandle = BASS_StreamCreateFile(BASSFALSE, filename, 0, 0, BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
        If StreamHandle = 0 Then StreamHandle = BASS_StreamCreateURL(filename, 0, 0, Null, 0)
        If StreamHandle = 0 Then
            Exit Sub
        Else
            Stream = StreamHandle
        End If
        If bolAutoPlay Then
            PlayStream
        End If
    Else
        StreamHandle = BASS_StreamCreateFile(BASSFALSE, filename, 0, 0, BASS_STREAM_DECODE)
        If StreamHandle = 0 Then StreamHandle = BASS_StreamCreateURL(filename, 0, BASS_STREAM_DECODE, 0, 0)
        If StreamHandle = 0 Then
            Exit Sub
        Else
            Stream = StreamHandle
        End If
   End If
End Sub

Public Sub PauseStream()
    BASS_ChannelPause Stream
    lngPlayState = 2 'Paused
    tmrState.Enabled = False
    RaiseEvent PlayStateChange
End Sub

Public Property Get handle() As Double
    handle = Stream
End Property


Public Property Let Position(val As Double)
    BASS_ChannelSetPosition Stream, BASS_ChannelSeconds2Bytes(Stream, val)
End Property

Public Property Get Position() As Double
        Position = BASS_ChannelBytes2Seconds(Stream, BASS_ChannelGetPosition(Stream))
End Property
Public Property Let Balance(New_Val As Double)
        BASS_ChannelSetAttributes Stream, 0, dblVol, New_Val
End Property
Public Property Let Rate(New_Val As Double)
        BASS_ChannelSetAttributes Stream, New_Val, -1, -101
End Property
Public Property Get Rate() As Double
        Call BASS_ChannelGetInfo(Stream, info)
        Rate = info.freq
End Property

Public Property Let volume(New_Vol As Double)
    BASS_ChannelSetAttributes Stream, 0, New_Vol, -1
    dblVol = New_Vol
End Property

Public Property Get volume() As Double
    volume = dblVol
End Property

Public Property Get Duration() As Double
    Duration = modBass.BASS_ChannelBytes2Seconds(Stream, BASS_ChannelGetLength(Stream))
End Property
Public Property Let AutoPlay(bolVal As Boolean)
    bolAutoPlay = bolVal
End Property
Public Property Get AutoPlay() As Boolean
    AutoPlay = bolAutoPlay
End Property
Public Property Get hwnd()
    hwnd = UserControl.hwnd
End Property
Public Function InitBass(ByVal device As Long, ByVal freq As Long, ByVal flags As Long, ByVal clsid As Long) As Boolean
    If (BASS_Init(device, freq, flags, UserControl.hwnd, clsid)) = BASSFALSE Then
        InitBass = False
    Else
        InitBass = True
    End If
End Function
Public Function SetDevice(device As Long) As Boolean
    If BASS_SetDevice(device) = BASSFALSE Then
        SetDevice = False
    Else
        SetDevice = True
    End If
End Function

Public Function GetTotalCDDrive() As Long
    ChDrive App.path
    ChDir App.path & "\"
    Dim intCount As Long
    intCount = 0
    While (BASS_CD_GetDriveDescription(intCount))
        intCount = intCount + 1
    Wend
    GetTotalCDDrive = intCount
End Function
Public Function GetCDDecs(nDrive As Long) As String
    On Error Resume Next
    Dim lngDecs  As Long
    lngDecs = BASS_CD_GetDriveDescription(nDrive)
    GetCDDecs = Chr$(65 + BASS_CD_GetDriveLetter(nDrive)) & ": " & VBStrFromAnsiPtr(lngDecs)
End Function
Public Sub OpenCDDrive(nDrive As Long, strAction As String)
    On Error Resume Next
    If LCase(strAction) = "open" Then
        BASS_CD_Door nDrive, BASS_CD_DOOR_OPEN
    Else
        BASS_CD_Door nDrive, BASS_CD_DOOR_CLOSE
    End If
End Sub
Public Function GetTotalTrack(ByVal nDrive As Long) As Long
    ChDrive App.path
    ChDir App.path & "\"
    If BASS_CD_IsReady(nDrive) Then
        GetTotalTrack = BASS_CD_GetTracks(nDrive)
    End If
End Function
Public Sub ReadCDTrack(ByVal nDrive As Long, ByVal nTrack As Long, Decode As Boolean)
    On Error Resume Next
    If Not Decode Then
        Stream = BASS_CD_StreamCreate(nDrive, nTrack, BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX) ' create stream
    Else
        Stream = BASS_CD_StreamCreate(nDrive, nTrack, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    End If
End Sub
Public Function GetInfo(intInfor As Integer) As Long
    On Error Resume Next
    Call BASS_ChannelGetInfo(Stream, info)
    Select Case LCase(intInfor)
        Case 0
            GetInfo = info.chans
        Case 1
            GetInfo = info.freq
        Case 2
            GetInfo = info.flags
        Case 3
            GetInfo = info.origres
        Case 4
            GetInfo = info.ctype
        Case 5
            GetInfo = info.ctype
    End Select
End Function
Public Sub ExitWaveWrite()
    On Local Error Resume Next

    'complete WAVE header
    wr.riffBlockSize = pos + 44 - 8
    wd.dataBlockSize = pos
    Put #fn, 5, wr.riffBlockSize
    Put #fn, 41, wd.dataBlockSize
    Close #fn
    lngPlayState = 0
    RaiseEvent PlayStateChange
End Sub

Public Sub WriteWave(strFileOut As String)
    On Local Error GoTo ErrHandle
    
        Call BASS_ChannelGetInfo(Stream, info)
    
        wf.wFormatTag = 1
        wf.nChannels = info.chans
        Call BASS_ChannelGetAttributes(Stream, wf.nSamplesPerSec, -1, -1)
        wf.wBitsPerSample = IIf(info.flags And BASS_SAMPLE_8BITS, 8, 16)
        wf.nBlockAlign = wf.nChannels * wf.wBitsPerSample / 8
        wf.nAvgBytesPerSec = wf.nSamplesPerSec * wf.nBlockAlign
    
        'Set WAV "fmt " header
        wf.wfBlockType = &H20746D66      '"fmt "
        wf.wfBlockSize = 16
    
        'Set WAV "RIFF" header
        wr.RIFF = &H46464952             '"RIFF"
        wr.riffBlockSize = 0             'after conversion
        wr.riffBlockType = &H45564157    '"WAVE"
    
        'set WAV "data" header
        wd.dataBlockType = &H61746164    '"data"
        wd.dataBlockSize = 0             'after conversion
        
        fn = FreeFile
        'create a file WAV
        If (FileExists(strFileOut)) Then
            Call Kill(strFileOut)  'delete if already created and create a new one
        End If
        Open strFileOut For Binary Lock Read Write As #fn
        
            'Write WAV Header to file
            Put #fn, , wr    'RIFF
            Put #fn, , wf    'Format
            Put #fn, , wd    'data
        
            ReDim buf(19999) As Byte
            While BASS_ChannelIsActive(Stream) = BASS_ACTIVE_PLAYING
                lngPlayState = 1
                RaiseEvent PlayStateChange
                ReDim Preserve buf(BASS_ChannelGetData(Stream, buf(0), 20000) - 1) As Byte
                'write data to WAV file
                Put #fn, , buf
                pos = BASS_ChannelGetPosition(Stream)
                Sleep 1
                DoEvents        'in case you want to exit...
            Wend
            
            'complete WAV header
            wr.riffBlockSize = pos + 44 - 8
            wd.dataBlockSize = pos
        
            Put #fn, 5, wr.riffBlockSize
            Put #fn, 41, wd.dataBlockSize
        Close #fn
        lngPlayState = 0
        RaiseEvent PlayStateChange
        RaiseEvent EndOfStream
    
ErrHandle:
    Call ExitWaveWrite
End Sub

Private Sub UserControl_Terminate()
    BASS_Stop
    BASS_Free
    BASS_PluginFree (0)
End Sub
