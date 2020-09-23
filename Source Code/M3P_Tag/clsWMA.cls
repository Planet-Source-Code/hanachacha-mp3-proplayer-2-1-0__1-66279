VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type WMATag
    Album As String
    Artist As String
    Title As String
    Year As String
    Bitrate As Integer
    Copyright As Boolean
    Genre As String
    GenreID As String
    Frequency As Long
    Length As Long
    Mode As String
    Licensed As Boolean
    FileSize As Long
    Header As String
    SDKVer As String
    SDKNeeded As String
End Type

Private WMAFile3 As IWMPMedia3
Private WMAFile2 As IWMPMedia2
Private WMAFile As IWMPMedia3
Private WMAReader As WMPCore
Private WMAPlayer As Object
Private WMA As WMATag
Public Sub Get_WMA_Header(ByVal strFileName As String)
    On Error Resume Next
    Dim strData As String, strData2 As String
    Dim intGoto As Integer, i As Integer
    Dim intStart As String, intStop As String
    Dim intTitleLen As Integer, intArtistLen As String
    Dim intCopyLen As Integer
    Dim strData3 As String, intTempData As String
    
    Dim myFSO As New FileSystemObject
    Dim myTextStream As TextStream
    Set myTextStream = myFSO.OpenTextFile(strFileName, ForReading, False)
    
    intGoto = 5000
    myTextStream.Skip (10)
    
    For i = 1 To intGoto
        strData = myTextStream.Read(1)
        If Not Asc(strData) = 0 Then
        strData2 = strData2 & Chr(CLng(Asc(strData)))
        Else
        i = i - 1
        End If
        'myTextStream.Skip (1)
    Next i
    
    myTextStream.Close
    Set myFSO = Nothing
    Set myTextStream = Nothing
    
    WMA.Header = strData2
    WMA.FileSize = FileLen(strFileName)
    
    intStart = InStr(1, strData2, "3&²uŽfÏ¦ÙªbÎlŽ") + 15
    If intStart > 0 Then
        intStop = intStart + InStr(1, Mid(strData2, intStart + 1, Len(strData2) - intStart - 1), "@¤ÐÒ")
    End If
    
    If intStart > 0 And intStop > 0 Then
        strData3 = Mid(strData2, intStart, intStop - intStart)
        Dim intReadDat As Integer
        
        intReadDat = 2
        If Mid(strData3, 1, 1) = Chr(1) Or Mid(strData3, 1, 1) = Chr(2) Then
            intReadDat = 2
        End If
        intTitleLen = (Asc(Mid(strData3, intReadDat, 1)) / 2) - 1
        intArtistLen = (Asc(Mid(strData3, intReadDat + 1, 1)) / 2) - 1
        intCopyLen = (Asc(Mid(strData3, intReadDat + 2, 1)) / 2) - 1
      
        If intTitleLen + intArtistLen = Len(strData3) - 2 Then
        'artist and title
            WMA.Artist = Mid(strData3, intReadDat + 2 + intTitleLen, intArtistLen)
            WMA.Title = Mid(strData3, intReadDat + 2, intTitleLen)
        End If
        
        If intTitleLen + intArtistLen + intCopyLen = Len(strData3) - 2 - intReadDat Then
            WMA.Artist = Mid(strData3, intReadDat + 3 + intTitleLen, intArtistLen)
            WMA.Title = Mid(strData3, intReadDat + 3, intTitleLen)
            WMA.Copyright = True
        End If
    End If
    
    Dim intKBPS As Integer, strKBPS As String
    Dim intFoundSDKV As Integer, intFoundSDKN As Integer
    Dim intGenre As Integer, intGenreID As Integer, strGenre As String, strGenreID As String
    Dim lngLength As Long
    Dim intFoundAl As Integer
    Dim intFoundFreq As Integer
    Dim intFoundMode As Integer
        
        intFoundSDKV = InStr(1, UCase(strData2), "WMFSDKVERSION")
        If intFoundSDKV > 0 Then
            WMA.SDKVer = Trim(Mid(strData2, intFoundSDKV + 14, 12))
        End If
        
        intFoundSDKN = InStr(1, UCase(strData2), "WMFSDKNEEDED")
        If intFoundSDKN > 0 Then
            WMA.SDKNeeded = Trim(Mid(strData2, intFoundSDKN + 13, 10))
        End If
        
        intGenreID = InStr(1, UCase(strData2), "WM/GENREID")
        If intGenreID > 0 Then
            strGenreID = Trim(Mid(strData2, intGenreID + 11, 5))
            WMA.GenreID = Trim(Replace(strGenreID, "", ""))
            WMA.Genre = WMA.GenreID
        End If
        
        intKBPS = InStr(1, UCase(strData2), "KBPS")
        WMA.Bitrate = 0
        If intKBPS > 0 Then
        strKBPS = Trim(Mid(strData2, intKBPS - 4, 3))
            If Not strKBPS = "" And IsNumeric(strKBPS) Then
                WMA.Bitrate = CInt(strKBPS)
            End If
        End If
        If WMA.Bitrate > 0 Then
            lngLength = CLng(((FileLen(strFileName) * 8) - 217480) / (CLng(WMA.Bitrate) * 1000))
            WMA.Length = lngLength
        End If
        
        intFoundFreq = InStr(1, UCase(strData2), "KHZ")
        Dim strFreq As String
        
        If intFoundFreq > 0 Then
            strFreq = Mid(strData2, intFoundFreq - 5, 4)
            strFreq = Trim(Replace(strFreq, ",", ""))
            If IsNumeric(strFreq) Then
                WMA.Frequency = CLng(strFreq & "000")
            End If
        End If
        
        If InStr(1, UCase(strData2), "STEREO") > 0 Then
            WMA.Mode = "Stereo"
        End If
        
        If InStr(1, UCase(strData2), "MONO") > 0 Then
            WMA.Mode = "Mono"
        End If
        
        If InStr(1, UCase(strData2), "SIGNATURE") > 0 Then
            WMA.Licensed = True
        End If
        
        Dim intFoundYear As Integer
        intFoundYear = InStr(1, UCase(strData2), "WM/YEAR")
        If intFoundYear > 0 Then
            intFoundYear = intFoundYear + 7
            WMA.Year = Mid(strData2, intFoundYear + 1, CInt(Asc(Mid(strData2, intFoundYear, 1)) / 2) - 1)
        End If
        
        intFoundAl = InStr(1, UCase(strData2), "WM/ALBUMTITLE")
        If intFoundAl > 0 Then
            intFoundAl = intFoundAl + 13
            WMA.Album = Mid(strData2, intFoundAl + 1, CInt(Asc(Mid(strData2, intFoundAl, 1)) / 2) - 1)
        End If
        
        Set WMAFile = WMAReader.newMedia(strFileName)
        Set WMAFile2 = WMAFile
        Set WMAFile3 = WMAFile2
        
        If WMA.Artist = "" Then
            WMA.Artist = WMAFile3.getItemInfo("Author")
            If WMA.Artist = "" Then
                WMA.Artist = WMAFile3.getItemInfo("WM/AlbumArtist")
            End If
        End If
        If WMA.Title = "" Then
            WMA.Title = WMAFile3.getItemInfo("Title")
        End If
        If WMA.Album = "" Then
            WMA.Album = WMAFile3.getItemInfo("WM/AlbumTitle")
        End If
        If WMA.Genre = "" Then
            WMA.Genre = WMAFile3.getItemInfo("WM/Genre")
        End If
        If WMA.GenreID = "" Then
            WMA.GenreID = WMAFile3.getItemInfo("WM/GenreID")
        End If
        If WMA.Bitrate = 0 Then
            WMA.Bitrate = WMAFile3.getItemInfo("Bitrate")
        End If
        Set WMAFile = Nothing
        Set WMAFile2 = Nothing
        Set WMAFile3 = Nothing
End Sub

Public Property Get Artist() As String
    Artist = WMA.Artist
End Property
Public Property Get Album() As String
    Album = WMA.Album
End Property
Public Property Get Title() As String
    Title = WMA.Title
End Property
Public Property Get Year() As String
    Year = WMA.Year
End Property
Public Property Get Genre() As String
    Genre = WMA.Genre
End Property
Public Property Get GenreID() As String
    GenreID = WMA.GenreID
End Property
Public Property Get Mode() As String
    Mode = WMA.Mode
End Property
Public Property Get Header() As String
    Header = WMA.Header
End Property
Public Property Get SDKVer() As String
    SDKVer = WMA.SDKVer
End Property
Public Property Get SDKNeeded() As String
    SDKNeeded = WMA.SDKNeeded
End Property
Public Property Get Bitrate() As Integer
    Bitrate = WMA.Bitrate
End Property
Public Property Get Copyright() As Boolean
    Copyright = WMA.Copyright
End Property
Public Property Get Frequency() As Long
    Frequency = WMA.Frequency
End Property
Public Property Get Length() As Long
    Length = WMA.Length
End Property
Public Property Get Licensed() As Boolean
    Licensed = WMA.Licensed
End Property
Public Property Get FileSize() As Long
    FileSize = WMA.FileSize
End Property
Public Property Let Artist(NewVal As String)
    WMA.Artist = NewVal
End Property
Public Property Let Album(NewVal As String)
    WMA.Album = NewVal
End Property
Public Property Let Title(NewVal As String)
    WMA.Title = NewVal
End Property
Public Property Let Year(NewVal As String)
    WMA.Year = NewVal
End Property
Public Property Let Genre(NewVal As String)
    WMA.Genre = NewVal
End Property
Public Function WriteTag(strWma As String)
    Set WMAFile = WMAReader.newMedia(strWma)
    Set WMAFile2 = WMAFile
    Set WMAFile3 = WMAFile2
    
    Call WMAFile3.setItemInfo("Title", WMA.Title)
    Call WMAFile3.setItemInfo("WM/AlbumTitle", WMA.Album)
    Call WMAFile3.setItemInfo("WM/AlbumArtist", WMA.Artist)
    Call WMAFile3.setItemInfo("WM/Genre", WMA.Genre)
    Call WMAFile3.setItemInfo("WM/Year", WMA.Year)
    Set WMAFile = Nothing
    Set WMAFile2 = Nothing
    Set WMAFile3 = Nothing
End Function
Private Sub Class_Initialize()
    Set WMAPlayer = New WindowsMediaPlayer
    Set WMAReader = WMAPlayer
End Sub

Private Sub Class_Terminate()
    Set WMAPlayer = Nothing
    Set WMAReader = Nothing
End Sub

