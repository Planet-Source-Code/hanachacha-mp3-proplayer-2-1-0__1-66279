Attribute VB_Name = "modGenre"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Const strGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
Public GenreArray() As String
Public Function ReturnGenreID3v2(strGenre As String) As String
    On Error GoTo beep
    If Mid(strGenre, 1, 1) = "(" Then
        ReturnGenreID3v2 = GenreArray(Mid(strGenre, 2, Len(strGenre) - 2))
    Else
        ReturnGenreID3v2 = strGenre
    End If
    Exit Function
beep:
    If Err.Number <> 0 Then ReturnGenreID3v2 = "Other"
End Function
Public Function ReturnGenre(iGenre As Byte) As String
    On Error GoTo beep
    ReturnGenre = GenreArray(iGenre)
    Exit Function
beep:
    If Err.Number <> 0 Then ReturnGenre = "Other"
End Function
Public Function ReturnGenreID(strGenre As String) As Byte
    On Error GoTo beep
    Dim y As Byte
    
    For y = 0 To UBound(GenreArray) - 1
        If LCase(strGenre) = LCase(GenreArray(y)) Then
            ReturnGenreID = y
            Exit For
        End If
    Next y
    Exit Function
beep:
    If Err.Number <> 0 Then ReturnGenreID = 12
End Function

