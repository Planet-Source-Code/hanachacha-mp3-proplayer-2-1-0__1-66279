Attribute VB_Name = "modPlaylist"
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

Public Sub SetSortString(strSortedString As String)
    On Error Resume Next
    Dim J As Long
    Dim s As Long
    Dim str As String
    For J = 1 To frmPlayList.List.ListItemCount
        s = frmPlayList.List.Key(J)
        With frmPlayList
            str = strSortedString
            str = Replace(str, "Artist", NowPlaying(s).Infor.Artist)
            str = Replace(str, "Title", NowPlaying(s).Infor.Title)
            str = Replace(str, "Album", NowPlaying(s).Infor.Album)
            str = Replace(str, "Grene", NowPlaying(s).Infor.Genre)
            str = Replace(str, "Year", NowPlaying(s).Infor.Year)
            str = Replace(str, "Filename", NowPlaying(s).Infor.Filename)
            str = Replace(str, "Pathname", NowPlaying(s).Infor.FullName)
            .List.AddSortString J, str
        End With
    Next J
End Sub
Public Function OpenGetFile() As Boolean
    Dim strFilePath As String
    Dim vFileSelected As Variant
    Dim lngNumber As Long
    
    On Error GoTo beep
    
    With frmMenu.cdloColor
        .InitDir = strLastDir
        .Filename = ""
        .DialogTitle = "MP3_ProPlayer - Open Media"
        .Filter = "M3P - able files|*.mp3;*.wma;*.wav;*.mid;*.midi;*.aif;*.avi;*.wmv;*.asf;*.mov;*.mpeg;*.mpg;*.mpe;*.mp2;*.ogg;*.mp1;*mp4" & _
                "|All Supported Audio Files |*.mp3;*.wma;*.wav;*.mid;*.ogg" & _
                "|All Supported Video Files |*.avi;*.wmv;*.asf;*.mov;*.mpeg;*.mpg;*.mpe;*.mp2;*.mp1;*.mp4;*.ogg" & _
                "|MP3 Files (*.mp3)|*.mp3" & _
                "|MIDI Files (*.mid,*.midi)|*.mid;*.midi" & _
                "|Mpeg File (*.mp3,*.mpeg,*.mpg,*.mpe,*.mp2,mp1)|*.mp3;*.mpeg;*.mpg;*.mpe;*.mp2;*.mp1" & _
                "|Mp4 (*.mp4)|*.mp4" & _
                "|Ogg Vorbis Files (*.ogg,*.ogm)|*.ogg;*.ogm" & _
                "|Wave Files (*.wav)|*.wav|Midi Files (*.mid)|*.mid" & _
                "|Windows Media File (*.wma,*.asf,*.avi,*.wmv)|*.wma;*.asf;*.avi;*.wmv" & _
                "|CD Track (*.cda)|*.cda" & _
                "|VCD Track (*.dat)|*.dat" & _
                "|All File (*.*)|*.*"
        .MaxFileSize = 16384
        .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
        .ShowOpen
        .CancelError = True
        vFileSelected = Split(.Filename, Chr(0))
        If UBound(vFileSelected) = 0 Then
            strFilePath = frmMenu.cdloColor.Filename
            strLastDir = ReturnFolderPath(strFilePath)
            If FileLen(strFilePath) > 0 Then
                Select Case LCase(Right(strFilePath, 4))
                    Case ".mp3", ".wav", ".wma", ".aif", ".asf", ".wmv", ".mpg", ".mpe", ".mp2", "mpeg", ".ogg", ".avi", ".mov", ".cda", ".dat", ".mid", "midi", ".mp1", ".mp4", ".ogm"
                        Call AddFile(strFilePath)
                    Case ".m3u", "m3u8"
                        Call LoadPlaylistM3U(strFilePath)
                    Case ".pls"
                        Call LoadPlaylistPLS(strFilePath)
                    Case Else
                        MsgBox "MP3_proPlayer not supported for this type", vbCritical, "Error"
                End Select
            End If
        Else
            For lngNumber = 1 To UBound(vFileSelected)
                strFilePath = vFileSelected(0) + "\" & vFileSelected(lngNumber)
                Select Case LCase(Right(strFilePath, 4))
                    Case ".mp3", ".mp1", ".mp2", ".mp4", ".mpe", ".mpg", "mpeg", ".mov", ".wav", ".wma", ".wmv", ".asf", ".avi", ".ogg", ".ogm", ".mid", "midi"
                        Call AddFile(strFilePath)
                    Case ".m3u", "m3u8"
                        Call LoadPlaylistM3U(strFilePath)
                        OpenGetFile = True
                        Exit Function
                    Case ".pls"
                        Call LoadPlaylistPLS(strFilePath)
                        OpenGetFile = True
                        Exit Function
                    Case Else
                        MsgBox "MP3_proPlayer not supported for this type", vbCritical, "Error"
                End Select
            Next
            strLastDir = ReturnFolderPath(strFilePath)
        End If
        With frmPlayList
            .List.DisplayList
        End With
        OpenGetFile = True
        Exit Function
    End With
beep:
    If Err.Number <> 0 Then
        OpenGetFile = False
    End If
End Function
Public Sub KillFile(Index As Long)
    If Index < 1 Or Index > UBound(NowPlaying) Then
        Exit Sub
    End If
    If MsgBox("Are you sure want to delete" & " ' " & GetShortName(NowPlaying(frmPlayList.List.Key(Index)).Infor.Filename) & " ' out disk ?", vbYesNo, "MP3_ProPlayer") = vbYes Then
        Kill NowPlaying(frmPlayList.List.Key(Index)).Infor.FullName
        Call SubFile(Index)
    End If
End Sub
Public Function GetShortName(strFileName As String)
    GetShortName = Mid(strFileName, InStrRev(strFileName, "\", -1, vbBinaryCompare) + 1)
End Function
Public Sub SubFile(Index As Long)
    On Error Resume Next
    Dim i, E As Long
    With frmPlayList
        If Index < 1 Or Index > .List.ListItemCount Then
            Exit Sub
        Else
            i = .List.Key(Index)
            
            For E = i To UBound(NowPlaying)
                NowPlaying(E) = NowPlaying(E + 1)
            Next E
            ReDim Preserve NowPlaying(UBound(NowPlaying))
            
            For E = 1 To .List.ListItemCount
                If .List.Key(E) > .List.Key(Index) Then
                    .List.Key(E) = .List.Key(E) - 1
                End If
            Next E
            
            .List.RemoveItem Index
            .List.DisplayList
            
            If Index = currentPlayIndex And (frmMedia.Player.PlayState = 1 Or frmVD.Video.State = playing) Then
                Call Play(Index - 1)
            End If
            If Index < currentPlayIndex Then currentPlayIndex = currentPlayIndex - 1
        End If
        .lblTotalTime = Time2String(CalTime)
        .lblCurrentperTotal.Caption = currentIndex & "/" & frmPlayList.List.ListItemCount
        .sldPl.max = .List.ListItemCount
        If .List.ListItemCount < .List.ItemPerPage Then
            .sldPl.Enabled = False
        Else
            .sldPl.Enabled = True
        End If
    End With
End Sub
Public Sub SubFolder()
    On Error Resume Next
    With frmMedia
        .lblDuration = "00:00"
        .lblPosition.Caption = "00:00"
    End With
    If frmPlayList.List.ListItemCount > 0 Then
        StopPlayer
        With frmPlayList
            .List.ClearItem
            .lblCurrentperTotal.Caption = ""
            .lblTotalTime.Caption = ""
            .sldPl.max = 1
            .sldPl.value = 1
        End With
        ReDim NowPlaying(0)
    End If
    currentIndex = 0
    currentRIndex = 0
End Sub
Public Sub AddPl()
    
    On Error GoTo beep
    Dim strFilePath As String
    Dim i As Long
    
    With frmMenu.cdloColor
        .InitDir = strLastDir
        .DefaultExt = "m3u"
        .Filter = "MediaPlaylist (*.m3u,*.pls) |*.m3u;*.pls|m3u Playlist (*.m3u)|*.m3u|Pls Playlist (*.pls )|*.pls"
        .flags = cdlOFNFileMustExist Or cdlOFNExplorer
        .ShowOpen
        .CancelError = True
        strFilePath = .Filename
        strLastDir = .Name
    End With
    If strFilePath <> "" Then
        Call SubFolder
        For i = 0 To frmMenu.mnuSkinC.Count - 1
            frmMenu.mnuSortC(i).Checked = False
        Next i
        If LCase(Right(strFilePath, 3)) = "m3u" Then Call LoadPlaylistM3U(strFilePath)
        If LCase(Right(strFilePath, 3)) = "pls" Then Call LoadPlaylistPLS(strFilePath)
    End If
beep:
    If Err.Number <> 0 Then
        Exit Sub
    End If
End Sub
Public Sub AddFile(strFilePath As String)
    On Error Resume Next
    
    Dim strText As String
    Dim i As Integer
    Dim x As Long
    Dim tTrack As track
    strText = ""
    strText = tPlaylistConfig.strDisplay
    
    If FileExists(strFilePath) Then
        With frmPlayList
            If LibOption.bolUse Then
                Select Case LCase(Right(strFilePath, 4))
                    Case ".cda", ".dat", ".mid", "midi"
                        Call OpenTrack(strFilePath, tTrack)
                    Case Else
                        i = TrackIndex(strFilePath)
                        If i = -1 Then 'Not exists in library so we add it to lirary
                            Call AddNewFile(strFilePath)
                        End If
                        
                        x = TrackIndex(strFilePath)
                        tTrack = Library(x).Infor
                        If x = -1 Then 'Can't add to Library because files on removeable disk
                            Call OpenTrack(strFilePath, tTrack)
                        End If
                End Select
            Else
                If tPlaylistConfig.intLoadID = 0 Then
                    Call OpenTrack(strFilePath, tTrack)
                End If
            End If
            
            'NowPlaying
            ReDim Preserve NowPlaying(UBound(NowPlaying) + 1)
            
            
            If tPlaylistConfig.intLoadID = 0 Then
                NowPlaying(UBound(NowPlaying)).Infor = tTrack
                strText = Replace(strText, "%1", tTrack.Artist)
                strText = Replace(strText, "%2", tTrack.Title)
                strText = Replace(strText, "%3", tTrack.Album)
                strText = Replace(strText, "%4", tTrack.Genre)
                strText = Replace(strText, "%5", tTrack.Year)
                strText = Replace(strText, "%6", tTrack.Filename)
                strText = Replace(strText, "%7", strFilePath)
                NowPlaying(UBound(NowPlaying)).strText = strText
            Else
                NowPlaying(UBound(NowPlaying)).Infor.FullName = strFilePath
                NowPlaying(UBound(NowPlaying)).Infor.ExtType = LCase(Right(strFilePath, 4))
                If Mid(NowPlaying(UBound(NowPlaying)).Infor.ExtType, 1, 1) <> "." Then NowPlaying(UBound(NowPlaying)).Infor.ExtType = "." & NowPlaying(UBound(NowPlaying)).Infor.ExtType
                NowPlaying(UBound(NowPlaying)).strText = GetShortName(strFilePath)
            End If
            
            .List.AddItem UBound(NowPlaying), NowPlaying(UBound(NowPlaying)).strText, Time2String(NowPlaying(UBound(NowPlaying)).Infor.Duration)
            .sldPl.max = .List.ListItemCount
            .sldPl.value = currentIndex
            .lblTotalTime = Time2String(CalTime)
            .lblCurrentperTotal.Caption = currentIndex & "/" & frmPlayList.List.ListItemCount
            
            If .sldPl.max > .List.ItemPerPage Then
                .sldPl.Enabled = True
            End If
            
            .lblTotalTime.Left = .ScaleWidth - PlaylistRad(1).Right - .lblTotalTime.width
        
        End With
    End If
End Sub
Public Sub LoadPlaylistM3U(Filename As String)
    On Error Resume Next
    Dim playlistitem As String
    Dim strFile As String
    Dim fn As Long
    fn = FreeFile
    Open Filename For Input As #fn
        Do Until EOF(fn)
            Line Input #fn, playlistitem
            If LCase(Mid(playlistitem, 1, 8)) <> "#extinf:" And UCase(Mid(playlistitem, 1, 8)) <> "#extm3u" Then
                If InStr(1, playlistitem, ":") <> 0 Then
                    strFile = playlistitem
                Else
                    If FileExists(Left(Filename, InStrRev(Filename, "\")) & playlistitem) Then
                        strFile = Left(Filename, InStrRev(Filename, "\")) & playlistitem
                    Else
                        strFile = playlistitem
                    End If
                End If
                Call AddFile(playlistitem)
            End If
        Loop
    Close #fn
    frmPlayList.List.DisplayList
End Sub
Public Sub LoadPlaylistPLS(Filename As String)
    On Error Resume Next
    Dim playlistitem As String
    Dim strFile As String
    Dim fn As Long
    fn = FreeFile
    Open Filename For Input As #fn
        Do Until EOF(fn)
            Input #fn, playlistitem
            If LCase(Mid(playlistitem, 1, 4)) = "file" Then
                strFile = Mid(playlistitem, InStr(1, playlistitem, "=", vbBinaryCompare) + 1)
                If FileExists(strFile) Then
                    Call AddFile(strFile)
                End If
            End If
        Loop
    Close #fn
    frmPlayList.List.DisplayList
End Sub
Public Sub SaveM3U(strPath As String)
    On Error GoTo beep
    With frmPlayList
        Dim intRecord As Long
        Dim fn As Integer
        fn = FreeFile
        Open strPath For Output As #fn
            Print #fn, "#EXTM3U"
            For intRecord = 1 To .List.ListItemCount
                Print #fn, "#EXTINF:" & NowPlaying(.List.Key(intRecord)).Infor.Duration & "," & NowPlaying(.List.Key(intRecord)).strText
                Print #fn, NowPlaying(.List.Key(intRecord)).Infor.FullName
            Next intRecord
        Close #fn
    End With
    Exit Sub
beep:
    Debug.Print Err.Description
    Exit Sub
End Sub
Public Sub SavePLS(strPath As String)
    On Error Resume Next
    With frmPlayList
        Dim intRecord As Long
        Dim fn As Integer
        fn = FreeFile
        Open strPath For Output As #fn
            Print #fn, "[playlist]"
            For intRecord = 1 To .List.ListItemCount
                Print #fn, "File" & intRecord & "=" & NowPlaying(.List.Key(intRecord)).Infor.FullName
                Print #fn, "Title" & intRecord & "=" & NowPlaying(.List.Key(intRecord)).strText
                Print #fn, "Lenght" & intRecord & "=" & NowPlaying(.List.Key(intRecord)).Infor.Duration
            Next intRecord
            Print #fn, "NumberOfEntries=" & .List.ListItemCount
            Print #fn, "Version=2"
        Close #fn
    End With
End Sub
Public Function CalTime() As Long
    Dim Timelenght As Long
    Dim i As Long
    For i = 1 To UBound(NowPlaying)
        Timelenght = Timelenght + NowPlaying(i).Infor.Duration
    Next i
    CalTime = Timelenght
End Function
Public Function RepairPath(ByVal path As String, file As String) As String
    If Right$(path, 1) <> "\" Then path = path & "\"
    RepairPath = path & file
End Function
Public Function ReturnFolderPath(Filepath As String) As String
    ReturnFolderPath = Mid(Filepath, 1, Len(Filepath) - Len(GetShortName(Filepath)))
End Function
Public Sub OpenTrack(strFile As String, media As track)
    On Error Resume Next
    ClearTrack media
    Dim fso As New FileSystemObject
    Dim fFile As file
    Set fFile = fso.GetFile(strFile)
    With media
        .FullName = strFile
        .Filename = GetShortName(strFile)
        .Size = fFile.Size
        .ExtType = LCase(Right(.Filename, 4))
        If Mid(.ExtType, 1, 1) <> "." Then .ExtType = "." & .ExtType
    End With
    Select Case media.ExtType
        Case ".mp3", ".mp1", ".mp2"
            With media
                If .ExtType = ".mp3" Then
                    Dim mID3 As New clsID3v2
                    If mID3.ReadID3v2Tag(.FullName) = True Then
                        .Album = mID3.Album
                        .Artist = mID3.Artist
                        .Title = mID3.Title
                        .Genre = mID3.Genre
                        .Year = mID3.Year
                    Else
                        Dim mID31 As New clsID3v1
                        Call mID31.ReadTag(.FullName)
                        .Album = mID31.Album
                        .Artist = mID31.Artist
                        .Title = mID31.Title
                        .Genre = ReturnGenre(mID31.Genre)
                        .Year = mID31.Year
                    End If
                End If
                Dim MPG As New clsMPEG
                Call MPG.ReadMPEGHeader(.FullName)
                .bitrate = MPG.bitrate
                .Duration = MPG.length
                .Frequency = MPG.Frequency
            End With
        Case ".wma"
            Dim WMA As New clsWMA
            
            Call WMA.Get_WMA_Header(media.FullName)
            With media
                .Album = WMA.Album
                .Title = WMA.Title
                .Artist = WMA.Artist
                .bitrate = WMA.bitrate
                .Duration = WMA.length
                .Frequency = WMA.Frequency
                .Genre = WMA.Genre
                .Year = WMA.Year
            End With
        Case ".ogg"
            Dim Ogg As New clsOGG
            Call Ogg.GetOGGTag(media.FullName)
            With media
                .Album = Ogg.oAlbum
                .Artist = Ogg.oArtist
                .Title = Ogg.oTitle
                .Genre = Ogg.oGenre
            End With
        Case ".wav"
            Dim wav As WAVE
            With media
                Call ReadWave(.FullName, wav)
                .Duration = wav.Lenght
                .Frequency = wav.Freq
                If wav.Mode = "Stereo" Then
                    .bitrate = 1411
                Else
                    .bitrate = 705
                End If
            End With
        Case ".avi", ".mov", ".asf", ".wmv", ".dat", ".mpg", ".mpe", ".mpeg"
            Dim VID As New clsAVI
            With media
                Call VID.ReadVideo(.FullName)
                .Artist = VID.Artist
                .Title = VID.Title
                .bitrate = VID.bitrate
                .Duration = VID.Lenght
            End With
    End Select
    If Trim(media.Album) = "" Then media.Album = "&No Album"
    If Trim(media.Artist) = "" Then media.Artist = "&No Artist"
    If Trim(media.Title) = "" Then media.Title = Mid(media.Filename, 1, Len(media.Filename) - Len(media.ExtType))
    If Trim(media.Genre) = "" Then media.Genre = "&No Genre"
    If Trim(media.Year) = "" Then media.Year = Year(Now)
    Set fFile = Nothing
End Sub
Private Sub ClearTrack(media As track)
    On Error Resume Next
    With media
        .Album = ""
        .Artist = ""
        .bitrate = 1411
        .Duration = 0
        .ExtType = ""
        .Filename = ""
        .FullName = ""
        .Frequency = 44100
        .Genre = ""
        .Title = ""
        .Year = ""
    End With
End Sub
Public Sub AddSingleFolder(strFolder As String)
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Call FindFileFolder(strFolder, "*.mp3")
    Call FindFileFolder(strFolder, "*.wma")
    Call FindFileFolder(strFolder, "*.aif")
    Call FindFileFolder(strFolder, "*.wav")
    Call FindFileFolder(strFolder, "*.wmv")
    Call FindFileFolder(strFolder, "*.mpg")
    Call FindFileFolder(strFolder, "*.mpe")
    Call FindFileFolder(strFolder, "*.mp2")
    Call FindFileFolder(strFolder, "*.mp1")
    Call FindFileFolder(strFolder, "*.ogg")
    Call FindFileFolder(strFolder, "*.avi")
    Call FindFileFolder(strFolder, "*.asf")
    Call FindFileFolder(strFolder, "*.mov")
    Call FindFileFolder(strFolder, "*.cda")
    Call FindFileFolder(strFolder, "*.dat")
    Call FindFileFolder(strFolder, "*.mid")
    
    frmPlayList.List.DisplayList
    Screen.MousePointer = vbNormal
End Sub
Public Sub AddSubFolder(strTopFolder As String)
    Dim DirList As New Collection
    Dim temp As String
    DirList.Add RepairPath(strTopFolder, "")
    Do While DirList.Count
        temp = Dir$(DirList(1), vbDirectory)
        Do Until temp = ""
            If temp = "." Or temp = ".." Then
            ElseIf (GetAttr(RepairPath(DirList(1), temp)) And vbDirectory) = vbDirectory Then
                DirList.Add RepairPath(DirList(1), temp) & "\"
            ElseIf InStr(temp, ".") Then
            End If
            temp = Dir$
        Loop
        AddSingleFolder (DirList(1))
        DirList.Remove 1
    Loop
End Sub

'This function return a string from lLenghtTime of Stream
Public Function Time2String(TimeLen As Long) As String
    Dim str As String
    Dim strHour As String
    Dim strMin As String
    Dim strSec As String
    If TimeLen >= 3600 Then
        strHour = CStr(TimeLen \ 3600)
        strSec = CStr(TimeLen - (TimeLen \ 3600) * 3600)
        strMin = CStr(Val(strSec \ 60))
        strSec = CStr(TimeLen - Val(strHour) * 3600 - Val(strMin) * 60)
        Time2String = Format(strHour, "00") & ":" & Format(strMin, "00") & ":" & Format(strSec, "00")
    Else
        strMin = CStr(TimeLen \ 60)
        strSec = CStr(TimeLen / 60 - TimeLen \ 60) * 60
        Time2String = Format(strMin, "00") & ":" & Format(strSec, "00")
    End If
End Function
Public Function String2Time(strTime As String) As Long
    Dim strMin As String
    Dim strSec As String
    If strTime = "" Then strTime = "00:00"
        strMin = Mid(strTime, 1, InStr(1, strTime, ":", vbBinaryCompare) - 1)
        strSec = Mid(strTime, InStr(1, strTime, ":", vbBinaryCompare) + 1)
        String2Time = CInt(strMin) * 60 + CInt(strSec)
End Function
Public Function Bol2String(bBol As Boolean) As String
    If bBol = True Then
        Bol2String = "Yes"
    Else
        Bol2String = "No"
    End If
End Function
Public Sub FindFileFolder(strFolder As String, ExtFile As String)
    Dim file$, hFile&, FD As WIN32_FIND_DATA, Patt$, Root$
    Dim x As Long
    
    Root = strFolder & "\"
    Patt = ExtFile
    hFile = FindFirstFile(Root & Patt, FD)
    If hFile = 0 Or hFile = -1 Then Exit Sub
    Do
       file = Left(FD.cFileName, InStr(FD.cFileName, Chr(0)) - 1)
       If Not (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
         = FILE_ATTRIBUTE_DIRECTORY Then
         If (file <> ".") And (file <> "..") Then
            AddFile (strFolder & "\" & file)
         End If
       End If
    Loop While FindNextFile(hFile, FD)
    Call FindClose(hFile)
End Sub
Public Function AddCDTrack(ByVal nDrive As Long) As Boolean
    If BASS_CD_IsReady(nDrive) Then
        Dim a As Long, L As Long, cdtext As Long, text As String
        cdtext = BASS_CD_GetID(nDrive, BASS_CDID_TEXT) 'get CD-TEXT
        
        Call SubFolder
        
        For a = 1 To frmMedia.Player.GetTotalTrack(nDrive)
            ReDim Preserve NowPlaying(a)
            L = BASS_CD_GetTrackLength(nDrive, a - 1)
            text = "Track " & Format(a, "00")
            If (cdtext) Then
                Dim T As Long, tag As String
                T = cdtext
                tag = "TITLE" & a + 1 & "="  'the CD-TEXT tag to look for
                Do
                    If (Mid(VBStrFromAnsiPtr(T), 1, Len(tag)) = tag) Then 'found the track title...
                        text = VBStrFromAnsiPtr(T + Len(tag)) 'replace "track x" with title
                        Exit Do
                    End If
                    T = T + Len(VBStrFromAnsiPtr(T)) + 1
                Loop While (VBStrFromAnsiPtr(T) <> "")
            End If
            NowPlaying(a).strText = text
            If (L = -1) Then
                NowPlaying(a).Infor.Duration = 0
            Else
                L = L / 176400
                NowPlaying(a).Infor.Duration = L
            End If
            frmPlayList.List.AddItem UBound(NowPlaying), NowPlaying(a).strText, Time2String(NowPlaying(a).Infor.Duration)
            frmPlayList.lblCurrentperTotal.Caption = currentIndex & "/" & frmPlayList.List.ListItemCount
            If frmPlayList.sldPl.max > frmPlayList.List.ItemPerPage Then
                frmPlayList.sldPl.Enabled = True
            End If
            frmPlayList.lblTotalTime.Left = frmPlayList.ScaleWidth - PlaylistRad(1).Right - frmPlayList.lblTotalTime.width
        Next a
        AddCDTrack = True
        tCDplay.CurrentDrive = nDrive
        tCDplay.CurrentTrack = 0
        tCDplay.TotalTrack = frmMedia.Player.GetTotalTrack(nDrive)
        frmPlayList.List.DisplayList
        frmPlayList.sldPl.max = tCDplay.TotalTrack + 1
        frmPlayList.sldPl.Enabled = True
    Else
        AddCDTrack = False
    End If
End Function
Public Sub UpdateCDInfo()
    Dim i As Long
    frmMenu.mnuCDRom(0).Enabled = False
    frmMenu.mnuCDRom(0).Caption = "System CDDrive"
    For i = 1 To frmMenu.mnuCDRom.UBound
        Unload frmMenu.mnuCDRom(i)
    Next i
    For i = 0 To frmMedia.Player.GetTotalCDDrive - 1
        Load frmMenu.mnuCDRom(i + 1)
        frmMenu.mnuCDRom(i + 1).Caption = Mid(frmMedia.Player.GetCDDecs(i), 1, InStr(1, frmMedia.Player.GetCDDecs(i), " ", vbBinaryCompare))
        If BASS_CD_IsReady(i) Then
            frmMenu.mnuCDRom(i + 1).Caption = frmMenu.mnuCDRom(i + 1).Caption & " Unknow CD"
        Else
            frmMenu.mnuCDRom(i + 1).Caption = frmMenu.mnuCDRom(i + 1).Caption & " No Disc In Drive"
        End If
        frmMenu.mnuCDRom(i + 1).Visible = True
        frmMenu.mnuCDRom(i + 1).Enabled = True
    Next i
End Sub

