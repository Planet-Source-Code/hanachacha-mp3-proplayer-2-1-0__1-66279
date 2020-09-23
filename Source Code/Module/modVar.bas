Attribute VB_Name = "modVariables"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit


'Variables
Public strFileconfig As String
Public strTitle As String
Public bolLoading As Boolean
'[App]
Public Type AppGeneral
    bolShowSplash As Boolean 'Show splash screen
    bolAutoStart As Boolean 'Auto open when windows start
    bolSysTray As Boolean 'Show app in system tray
    bolTaskbar As Boolean 'show app in taskbar
    bolTaskbarScroll As Boolean
    bolOnTop As Boolean ' ? AlwaysOnTop
    bolMenu As Boolean
    intIcon As Integer
    intFileIcon As Integer
    intPLIcon As Integer
End Type

'[Skin]
Private Type SkinInfor
    Name As String
    Author As String
    Comment As String
    Location As String
End Type

Public Type M3PSkin
    Infor As SkinInfor
    strConfig  As String
    mini As Boolean
End Type

Public Type SkinOption
    SkinDir As String
    bolEQSlide As Boolean
End Type

'[Device]
Private Type SoundDevice
    OutputDevice As Long
    Freq As Long
    WaveWrite As Boolean
End Type

Private Type VideoDevice
    intDefaultScreen As Integer
    bolLockRatio As Boolean
    intRatioHeight As Integer
    intRatioWidth As Integer
End Type

Private Type WaveDevice
    strWaveOutput As String
    bolAutoFilename As String
End Type

Public Type Device
    SoundD As SoundDevice
    VideoD As VideoDevice
    WaveD As WaveDevice
End Type

'[Playlist]
Public Type PlaylistConfig
    bolHidePL As Boolean '? Show playlist window
    bolShowNumber As Boolean 'Show ? number in playlist
    intLoadID As Integer ' When program load media information ?
    intSortKey As Integer 'Sortkey of playlist
    intLastFile As Integer 'Last file play
    intRowS As Integer 'Mow many row scroll ?
    strDisplay As String 'Format text display
    strSortString As String
    'Skin Variables
    bolPlayBold As Boolean 'Play Text Bold
    ePAlign As PicAlignment
    lngForeColor As Long
    lngBackColor As Long
    lngFontSize As Long 'Fontsize of playlist
    lngPlayColor As Long 'Palying format color
    lngPlayBackColor As Long
    lngSelectedColor As Long 'Back color of selectted Item
    lngSelectedBorderColor As Long
    strFontName As String 'Fontname of playlist
End Type

'[Player]
Public Type PlayerConfig
    strFileType As String
    bolAutoPlay As Boolean ' ? Auto play on start
    bolAutoRemove As Boolean
    bolAutoExit As Boolean
    bolAutoShutdow As Boolean
    bolCrossfade As Boolean
    bolMute As Boolean 'Mute ???
    bolLoop As Boolean ' ? Loop playlist
    bolRepeat As Boolean ' ? Repeat one
    bolShuffe As Boolean ' ? Shuffe
    bolShowEQ As Boolean 'Equalizer is active
    bolShowList As Boolean ' ? Auto load last playlist
    bolTimer As Boolean ' ? remaining or elapsed time
    bolScroll As Boolean
    intCrossfade As Integer
    intVolume As Integer ' volume
    intBalance As Integer 'balance
    intTime As Integer ' ? jump time
    lngMasterVol As Long 'System volume
End Type

Public Type CDplay
   CurrentDrive As Long
   CurrentTrack As Long
   TotalTrack As Long
End Type

'[Library]
Public Type track
    ExtType As String '*.??? Ex: .mp3,.mpeg, ....
    Artist As String
    Title As String
    Album As String
    Year As String
    Genre As String
    bitrate As Integer
    Frequency As Long
    Duration As Long
    Filename As String
    FullName As String
    Size As Long 'In byte
End Type

Public Type LibraryOption
    bolUse As Boolean
    intDblClick As Integer
    lngAudioSkip As Long
    lngVideoSkip As Long
    ColView(14) As Boolean
    ColWidth(14) As Long
End Type

Public Type MediaData
    Infor As track
    intPlaycount As Integer
    intRate As Integer
    strDay As String
    strDayUpdate As String
    strType As String * 5
End Type

Public Type Playlist
    Infor As track
    strText As String
End Type

Public Type LibPlaylist
    Name As String
    file As String
End Type

'[Winamp visualization]
Public Type WinampVisual
    bolEnabled As Boolean
    bolInit As Boolean
    intCurrentPlugin As Integer
    intSubPlugin As Integer
End Type

Public Type WinampDSP
    bolEnabled As Boolean
    intCurrentPlugin As Integer
End Type

'[DSP DirectX]
Public bolEQEnabled As Boolean
Public strCurrentEQPreset As String
Public bolUseDirectX As Boolean
Public bolDSP(7) As Boolean

Public lngPreAmpVal As Long
Public intEQ(9) As Long

Public intChorus(6) As Long 'Chorus value
Public intCompressor(5) As Long 'Compressor value
Public intDistortion(4) As Long 'Distortion value
Public intEcho(4) As Long 'Echo value
Public intFlanger(6) As Long
Public intGargle(1) As Long
Public intI3DL2Reverb(11) As Long
Public intReverb(3) As Long 'Revreb value

'[Visualization]

'Window extend
Public Type MainVis
    BackColor As Long
    BackGround As String
    bolUsePic As Boolean
    bolShowTitle As Boolean
    FontColor As Long
    plugin As String
    TimeDisplay As Long
    Data As Integer
    Style As Integer
End Type


Public sysTray As New clsSysTray
Public WinAPI As New clsWindow

'Main visualization
Public Type SkinVis
    intStyle As Integer
    intRefresh As Integer
    intSpecDraw As Integer
    intSpecFill As Integer
    intSpecPeakPause As Integer
    intSpecPeakDrop As Integer
    bolSpecPeak As Boolean
    intOsc As Integer
End Type

'[Public Variables]
Public tSkinVis As SkinVis
Public tMainWin As MainVis

Public tCurrentTrack As track
Public CurrentTrack As Long
Public CurrentPlaylist As Long

Public tAppConfig As AppGeneral
Public tPlaylistConfig As PlaylistConfig
Public tPlayerConfig As PlayerConfig
Public tSkinOption As SkinOption
Public tDevice As Device
Public tWinamp As WinampVisual
Public tWinampDSP As WinampDSP
Public tCurrentSkin As M3PSkin
Public tCDplay As CDplay
Public bolCDPlay As Boolean 'CD is playing
Public bolVideoOn As Boolean 'Video is playing
Public LibOption As LibraryOption
Public strLastDir As String
Public strLibrary As String
Public intTitleScroll As Integer

Public Library() As MediaData
Public Playlist() As LibPlaylist
Public NowPlaying() As Playlist

Public genPlugin() As Object
