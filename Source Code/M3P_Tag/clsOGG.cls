VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Original Source  :Simple OGG Vorbis tag-reader example by Intimidated of UniteTheCows.com
Option Explicit
Private Type VorbisTag
    Title As String
    Artist As String
    Album As String
    Genre As String
    Date As String
    Comment As String
    TrackNumber As Integer
    EncodedUsing As String
    Error As String
End Type
Private OGG As VorbisTag
Public Function GetOGGTag(strFileOGG As String) As Boolean
    Dim phase As Integer
    Dim filelength As Long
    Dim fileremaining As Long
    Dim fileopened As Long
    Dim filehandle As Integer
    Dim errmsg As String
    Dim tmp As Integer
    Dim tmp4 As Integer
    Dim tmp2 As Integer
    Dim tmp3 As Integer
    Dim foundatag As Boolean
    Dim s$
    OGG.Album = ""
    OGG.Artist = ""
    OGG.Comment = ""
    OGG.Date = ""
    OGG.EncodedUsing = ""
    OGG.Error = ""
    OGG.Genre = ""
    OGG.TrackNumber = 0
    OGG.Title = ""
    foundatag = False
    s$ = Space$(2048)
    'Phase is used so the error handler knows
    'what we are up to
    On Error GoTo Errhandler
    phase = 1
    'Phase 1: We're trying to open the file.
    filelength = FileLen(strFileOGG)
    fileremaining = filelength
    filehandle = FreeFile()
    Open strFileOGG For Binary Access Read As filehandle
    
    'Phase 2: File open.. so it exists..
    phase = 2
    Do
    Get filehandle, , s$
    fileopened = fileopened + Len(s$)
    fileremaining = fileremaining - Len(s$)
    If fileremaining < 2048 Then
    s$ = Space$(fileremaining)
    End If
    tmp = InStr(1, s$, "vorbis")
    tmp2 = InStr(1, s$, "vorbis ")
    
    
    If tmp <> 0 Then
    If tmp2 < tmp Then
    If tmp2 <> 0 Then
    tmp = tmp2
    End If
    End If
    Else
    If tmp2 <> 0 Then
    tmp = tmp2
    End If
    End If
    If tmp <> 0 Then Exit Do
    If fileremaining = 0 Then Exit Do
    Loop
    If tmp = 0 Then phase = 3: GoTo Errhandler
    'Ok, we've found the vorbis header.
    'Let's get a big chunk of data
    Get filehandle, tmp, s$
    'Ok, we've got 2kb of data after the header
    'lets find the header close
    tmp = InStr(7, s$, "vorbis")
    If tmp = 0 Then phase = 4: GoTo Errhandler
    
    'now we can get the required info
    
    tmp = InStr(1, s$, "TITLE=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.Title = Mid$(s$, tmp + 6, tmp2 - (tmp + 6))
    End If
    
    tmp = InStr(1, s$, "ARTIST=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.Artist = Mid$(s$, tmp + 7, tmp2 - (tmp + 7))
    End If
    
    tmp = InStr(1, s$, "COMMENT=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.Comment = Mid$(s$, tmp + 8, tmp2 - (tmp + 8))
    End If
    
    tmp = InStr(1, s$, "ALBUM=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.Album = Mid$(s$, tmp + 6, tmp2 - (tmp + 6))
    End If
    
    tmp = InStr(1, s$, "DATE=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.Date = Mid$(s$, tmp + 5, tmp2 - (tmp + 5))
    End If
    
    tmp = InStr(1, s$, "TRACKNUMBER=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.TrackNumber = Val(Mid$(s$, tmp + 12, tmp2 - (tmp + 12)))
    End If
    
    tmp = InStr(1, s$, "GENRE=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.Genre = Mid$(s$, tmp + 6, tmp2 - (tmp + 6))
    End If
    
    tmp = InStr(1, s$, "ENCODED_USING=")
    If tmp <> 0 Then
    foundatag = True
    tmp2 = InStr(tmp, s$, Chr$(0) + Chr$(0) + Chr$(0))
    tmp3 = InStr(tmp, s$, Chr$(1) + Chr$(5) + "vorbis")
    If tmp3 < tmp2 Then tmp2 = tmp3 + 1
    tmp2 = tmp2 - 1
    OGG.EncodedUsing = Mid$(s$, tmp + 14, tmp2 - (tmp + 14))
    End If
    If foundatag = False Then phase = 5: GoTo Errhandler
    Exit Function
Errhandler:
    Close filehandle
    If phase = 1 Then errmsg = "Error opening file! Not found? Already in use, perhaps?"
    If phase = 2 Then errmsg = "Error processing file. Ouchage."
    If phase = 3 Then errmsg = "Vorbis header not found?!"
    If phase = 4 Then errmsg = "Vorbis header not closed!"
    If phase = 5 Then errmsg = "I found a vorbis header, but no tag information seems to exist.."
    
    OGG.Error = "ERR:" + errmsg
    GetOGGTag = False
    Exit Function
End Function
Public Property Get oTitle() As String
    oTitle = OGG.Title
End Property
Public Property Get oArtist() As String
    oArtist = OGG.Artist
End Property
Public Property Get oAlbum() As String
    oAlbum = OGG.Album
End Property
Public Property Get oGenre() As String
    oGenre = OGG.Genre
End Property
Public Property Get oDate() As String
    oDate = OGG.Date
End Property
Public Property Get oEncodedUsing() As String
    oEncodedUsing = OGG.EncodedUsing
End Property
Public Property Get oComment() As String
    oComment = OGG.Comment
End Property
Public Property Get oError() As String
    oError = OGG.Error
End Property
Public Property Get oTrackNumber() As Integer
    oTrackNumber = OGG.TrackNumber
End Property

