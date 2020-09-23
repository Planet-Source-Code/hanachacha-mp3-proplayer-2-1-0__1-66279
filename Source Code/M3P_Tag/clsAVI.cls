VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private AVIFile3 As IWMPMedia3
Private AVIFile2 As IWMPMedia2
Private AVIFile As IWMPMedia3
Private AVIReader As WMPCore
Private AVIPlayer As Object

Private Type VideoTag
    Artist As String
    Title As String
    Bitrate As Long
    Lenght As Long
    FileSize As Long
End Type
Private VDT As VideoTag
Public Sub ReadVideo(strVideo As String)
    Set AVIFile = AVIReader.newMedia(strVideo)
    Set AVIFile2 = AVIFile
    Set AVIFile3 = AVIFile2
    
    VDT.Artist = AVIFile3.getItemInfo("Author")
    VDT.Title = AVIFile3.getItemInfo("Title")
    VDT.Bitrate = AVIFile3.getItemInfo("Bitrate")
    VDT.FileSize = AVIFile3.getItemInfo("Filesize")
    VDT.Lenght = AVIFile3.getItemInfo("Duration")
    
    Set AVIFile = Nothing
    Set AVIFile2 = Nothing
    Set AVIFile3 = Nothing
End Sub
Public Property Get Artist() As String
    Artist = VDT.Artist
End Property
Public Property Get Title() As String
    Title = VDT.Title
End Property
Public Property Get Bitrate() As Long
    Bitrate = VDT.Bitrate
End Property
Public Property Get Lenght() As Long
    Lenght = VDT.Lenght
End Property
Public Property Get FileSize() As Long
    FileSize = VDT.FileSize
End Property
Public Property Let Artist(NewArtist As String)
    VDT.Artist = NewArtist
End Property
Public Property Let Title(NewTitle As String)
    VDT.Title = NewTitle
End Property
Public Sub WriteVideoTag(strVideo As String)
    Set AVIFile = AVIReader.newMedia(strVideo)
    Set AVIFile2 = AVIFile
    Set AVIFile3 = AVIFile2
    
    Call AVIFile.setItemInfo("Author", VDT.Artist)
    Call AVIFile.setItemInfo("Title", VDT.Title)
    
    Set AVIFile = Nothing
    Set AVIFile2 = Nothing
    Set AVIFile3 = Nothing
End Sub
Private Sub Class_Initialize()
    Set AVIPlayer = New WindowsMediaPlayer
    Set AVIReader = AVIPlayer
End Sub

Private Sub Class_Terminate()
    Set AVIPlayer = Nothing
    Set AVIReader = Nothing
End Sub

