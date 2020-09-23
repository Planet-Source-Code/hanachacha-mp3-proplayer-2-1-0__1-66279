VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type ID3v2Tag
    Track As String
    Title As String
    Artist As String
    Album As String
    Year As String
    Genre As String
    Comments As String
    Composer As String
    OrigArtist As String
    Copyright As String
    url As String
    EncodedBy As String
End Type
Private MP3File3 As IWMPMedia3
Private MP3File2 As IWMPMedia2
Private MP3File As IWMPMedia
Private MP3Reader As WMPCore
Private MP3Player As Object
Private tagID3v2 As ID3v2Tag
Private Sub ClearTag()
    tagID3v2.Album = ""
    tagID3v2.Artist = ""
    tagID3v2.Comments = ""
    tagID3v2.Composer = ""
    tagID3v2.Copyright = ""
    tagID3v2.EncodedBy = ""
    tagID3v2.Genre = ""
    tagID3v2.OrigArtist = ""
    tagID3v2.Title = ""
    tagID3v2.Track = ""
    tagID3v2.url = ""
    tagID3v2.Year = ""
End Sub
Public Function ReadID3v2Tag(strMP3 As String) As Boolean
    On Error GoTo beep
    Set MP3File = MP3Reader.newMedia(strMP3)
    Set MP3File2 = MP3File
    Set MP3File3 = MP3File2
    Call ClearTag
    With tagID3v2
        .Album = MP3File3.getItemInfo("WM/AlbumTitle")
        .Artist = MP3File3.getItemInfo("WM/AlbumArtist")
        If .Artist = "" Then
            .Artist = MP3File3.getItemInfo("Author")
        End If
        .Comments = MP3File3.getItemInfo("Description")
        .Composer = MP3File3.getItemInfo("WM/Composer")
        .Copyright = MP3File3.getItemInfo("Copyright")
        .EncodedBy = MP3File3.getItemInfo("WM/EncodedBy")
        .Genre = MP3File3.getItemInfo("WM/Genre")
        .OrigArtist = MP3File3.getItemInfo("WM/OriginalArtist")
        .Title = MP3File3.getItemInfo("Title")
        .Track = MP3File3.getItemInfo("WM/TrackNumber")
        .url = MP3File3.getItemInfo("WM/UserWebURL")
        .Year = MP3File3.getItemInfo("WM/Year")
    End With
    ReadID3v2Tag = True
    Set MP3File = Nothing
    Set MP3File2 = Nothing
    Set MP3File3 = Nothing
beep:
    If Err.Number <> 0 Then
        ReadID3v2Tag = False
        Exit Function
    End If
End Function
Public Sub WriteID3v2Tag(strMP3 As String)
    On Error GoTo Errhandler
    
    Set MP3File = MP3Reader.newMedia(strMP3)
    Set MP3File2 = MP3File
    Set MP3File3 = MP3File2
    With MP3File3
        .setItemInfo "WM/TrackNumber", tagID3v2.Track
        .setItemInfo "Title", tagID3v2.Title
        .setItemInfo "WM/AlbumArtist", tagID3v2.Artist
        .setItemInfo "Author", tagID3v2.Artist
        .setItemInfo "WM/AlbumTitle", tagID3v2.Album
        .setItemInfo "WM/Year", tagID3v2.Year
        .setItemInfo "WM/Genre", tagID3v2.Genre
        .setItemInfo "Description", tagID3v2.Comments
        .setItemInfo "WM/Composer", tagID3v2.Composer
        .setItemInfo "WM/OriginalArtist", tagID3v2.OrigArtist
    End With
    Set MP3File = Nothing
    Set MP3File2 = Nothing
    Set MP3File3 = Nothing
Errhandler:
    If Err.Number <> 0 Then
        Exit Sub
    End If
End Sub

Public Sub RemoveID3v2tag(strMP3 As String)
    On Error Resume Next
    ClearTag
    WriteID3v2Tag (strMP3)
End Sub

Public Property Get Artist() As String
    Artist = tagID3v2.Artist
End Property
Public Property Get Track() As String
    Track = tagID3v2.Track
End Property
Public Property Get Title() As String
    Title = tagID3v2.Title
End Property
Public Property Get Album() As String
    Album = tagID3v2.Album
End Property
Public Property Get Year() As String
    Year = tagID3v2.Year
End Property
Public Property Get Genre() As String
    Genre = tagID3v2.Genre
End Property
Public Property Get Comments() As String
    Comments = tagID3v2.Comments
End Property
Public Property Get Composer() As String
    Composer = tagID3v2.Composer
End Property
Public Property Get OrigArtist() As String
    OrigArtist = tagID3v2.OrigArtist
End Property
Public Property Get Copyright() As String
    Copyright = tagID3v2.Copyright
End Property
Public Property Get url() As String
    url = tagID3v2.url
End Property
Public Property Get EncodedBy() As String
    EncodedBy = tagID3v2.EncodedBy
End Property
Public Property Let Artist(NewVal As String)
    tagID3v2.Artist = NewVal
End Property
Public Property Let Track(NewVal As String)
    tagID3v2.Track = NewVal
End Property
Public Property Let Title(NewVal As String)
    tagID3v2.Title = NewVal
End Property
Public Property Let Album(NewVal As String)
    tagID3v2.Album = NewVal
End Property
Public Property Let Year(NewVal As String)
    tagID3v2.Year = NewVal
End Property
Public Property Let Genre(NewVal As String)
    tagID3v2.Genre = NewVal
End Property
Public Property Let Comments(NewVal As String)
    tagID3v2.Comments = NewVal
End Property
Public Property Let Composer(NewVal As String)
    tagID3v2.Composer = NewVal
End Property
Public Property Let OrigArtist(NewVal As String)
    tagID3v2.OrigArtist = NewVal
End Property
Public Property Let Copyright(NewVal As String)
    tagID3v2.Copyright = NewVal
End Property
Public Property Let url(NewVal As String)
    tagID3v2.url = NewVal
End Property
Public Property Let EncodedBy(NewVal As String)
    tagID3v2.EncodedBy = NewVal
End Property

Private Sub Class_Initialize()
    Set MP3Player = New WindowsMediaPlayer
    Set MP3Reader = MP3Player
End Sub

Private Sub Class_Terminate()
    Set MP3Player = Nothing
    Set MP3Reader = Nothing
End Sub
