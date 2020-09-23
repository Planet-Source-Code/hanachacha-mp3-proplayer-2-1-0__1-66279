VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type ID3v1Tag
    ID As String * 3
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 28
    isTrack  As Byte
    Track As Byte
    Genre As Byte
End Type
Dim tag As ID3v1Tag


Public Function ReadTag(strMP3 As String) As Boolean
    On Error Resume Next
    Dim lFileSize As Long
    Dim fn As Integer
    
    ClearData
    fn = FreeFile
    Open strMP3 For Binary As #fn
            lFileSize = LOF(fn)
            Get #fn, lFileSize - 127, tag.ID
            If tag.ID = "TAG" Then
                Get #fn, lFileSize - 127, tag
                ReadTag = True
            Else
                ReadTag = False
                ClearData
                Exit Function
            End If
    Close #fn
    tag.Title = Replace(tag.Title, Chr(0), "", , , vbBinaryCompare)
    tag.Artist = Replace(tag.Artist, Chr(0), "", , , vbBinaryCompare)
    tag.Album = Replace(tag.Album, Chr(0), "", , , vbBinaryCompare)
    tag.Year = Replace(tag.Year, Chr(0), "", , , vbBinaryCompare)
    tag.Comment = Replace(tag.Comment, Chr(0), "", , , vbBinaryCompare)
End Function
Public Function WriteTag(strMP3 As String)
On Error Resume Next
    Dim lFileSize As Long
    Dim fn As Integer
    fn = FreeFile
        Open strMP3 For Binary As #fn
        lFileSize = LOF(fn)
        Get #fn, lFileSize - 127, tag.ID
            If tag.ID = "TAG" Then
                Put #fn, , tag
            Else
                Put #fn, lFileSize - 127, "TAG"
                Close #fn
                Call WriteTag(strMP3)
            End If
        Close #fn
End Function
Public Function RemoveID3v1(strMP3 As String)
    Dim FileData() As Byte
    Dim fn As Integer
    fn = FreeFile
    
    Open strMP3 For Binary As #fn
    Get #fn, FileLen(strMP3) - 127, tag
    If tag.ID = "TAG" Then
        ReDim FileData(FileLen(strMP3) - 129)
        Open strMP3 & ".temp" For Binary As #2
        Get #1, fn, FileData
        Put #2, fn, FileData
        Close
        Kill strMP3
        Name strMP3 & ".temp" As strMP3
    Else
        Close
    End If
End Function
Private Sub ClearData()
    tag.Album = ""
    tag.Artist = ""
    tag.Comment = ""
    tag.Genre = 12
    tag.Title = ""
    tag.Track = 0
    tag.isTrack = 0
    tag.Year = ""
End Sub

Public Property Get Artist() As String
    Artist = tag.Artist
End Property
Public Property Let Artist(NewArtist As String)
    tag.Artist = NewArtist
End Property
Public Property Get Album() As String
    Album = tag.Artist
End Property
Public Property Let Album(NewAlbum As String)
    tag.Album = NewAlbum
End Property
Public Property Get Title() As String
    Title = tag.Title
End Property
Public Property Let Title(NewTitle As String)
    tag.Title = NewTitle
End Property
Public Property Get Track() As Byte
    Track = tag.Track
End Property
Public Property Let Track(NewTrack As Byte)
    tag.Track = NewTrack
End Property
Public Property Get Year() As String
    Year = tag.Year
End Property
Public Property Let Year(NewYear As String)
    tag.Year = NewYear
End Property
Public Property Get Genre() As Byte
    Genre = tag.Genre
End Property
Public Property Let Genre(NewGenre As Byte)
    tag.Genre = NewGenre
End Property
Public Property Get Comment() As String
    Comment = tag.Comment
End Property
Public Property Let Comment(NewComment As String)
    tag.Comment = NewComment
End Property

