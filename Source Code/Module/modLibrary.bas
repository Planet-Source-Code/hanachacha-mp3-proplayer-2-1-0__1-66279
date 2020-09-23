Attribute VB_Name = "modLibrary"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim fso As New FileSystemObject
Public Sub AddNewFile(strFile As String)
    On Error GoTo handle
    Dim tTrack As track
    Dim x As Long
    Dim dDrive As drive
    
    If Not FileExists(strFile) Then Exit Sub
    If TrackIndex(strFile) <> -1 Then Exit Sub 'File already exists
    
    'Check If File on Removale Disk
    Set dDrive = fso.GetDrive(Mid(strFile, 1, 2))
    If dDrive.DriveType = CDRom Or dDrive.DriveType = Removable Or dDrive.DriveType = RamDisk Then Set dDrive = Nothing: Exit Sub
    Set dDrive = Nothing
    
    Call OpenTrack(strFile, tTrack)
    x = UBound(Library) + 1
    ReDim Preserve Library(x)
    Library(x - 1).Infor = tTrack
    Library(x - 1).intPlaycount = 0
    Library(x - 1).intRate = 0
    Library(x - 1).strDay = date
    Library(x - 1).strDayUpdate = date
    Select Case tTrack.ExtType
        Case ".mp3", ".wma", ".wav", ".ogg"
            Library(x - 1).strType = "Audio"
        Case ".mpg", ".wmv", ".asf", ".mpe", ".mpeg", ".mov", ".avi", ".aif", ".mp2", ".dat"
            Library(x - 1).strType = "Video"
    End Select
    Call WriteDataFile(x - 1)
    WriteINI "M3P_Library", "FileLen", UBound(Library), strLibrary
    Exit Sub
handle:
    If Err.Number <> 0 Then
        MsgBox "AddNewFile " & Err.Number & " " & Err.Description, vbOKOnly, "Error"
    End If
    
End Sub
Public Sub AddNewPlaylist(Index As Long)
    On Error Resume Next
    Dim temp As String
    If Index = 0 Then
        frmMenu.mnuSendCC(0).Caption = Playlist(0).Name
        frmMenu.mnuSendCC(0).Enabled = True
    Else
        Load frmMenu.mnuSendCC(Index)
        frmMenu.mnuSendCC(Index).Caption = Playlist(Index).Name
        frmMenu.mnuSendCC(Index).Enabled = True
    End If
End Sub


Public Sub AddToNowPlaying(strFilePath As String)
    On Error Resume Next
    
    Dim strText As String
    Dim x As Long
    Dim tTrack As track
    strText = ""
    strText = tPlaylistConfig.strDisplay
    If FileExists(strFilePath) Then
        With frmPlayList
            
            x = TrackIndex(strFilePath)
            tTrack = Library(x).Infor
            'NowPlaying
            ReDim Preserve NowPlaying(UBound(NowPlaying) + 1)
            NowPlaying(UBound(NowPlaying)).Infor = tTrack
            
            strText = Replace(strText, "%1", tTrack.Artist)
            strText = Replace(strText, "%2", tTrack.Title)
            strText = Replace(strText, "%3", tTrack.Album)
            strText = Replace(strText, "%4", tTrack.Genre)
            strText = Replace(strText, "%5", tTrack.Year)
            strText = Replace(strText, "%6", tTrack.Filename)
            strText = Replace(strText, "%7", strFilePath)
            NowPlaying(UBound(NowPlaying)).strText = strText
            
            .List.AddItem UBound(NowPlaying), NowPlaying(UBound(NowPlaying)).strText, Time2String(NowPlaying(UBound(NowPlaying)).Infor.Duration)
            .sldPl.max = .List.ListItemCount
            .sldPl.value = currentIndex
            .lblTotalTime = Time2String(CalTime)
            .lblCurrentperTotal.Caption = currentIndex & "/" & frmPlayList.List.ListItemCount
            If .sldPl.max > .List.ItemPerPage Then
                .sldPl.Enabled = True
            End If
            .lblTotalTime.Left = .ScaleWidth - PlaylistRad(1).Right - .lblTotalTime.width
            .List.DisplayList
        End With
    End If
End Sub

Public Function TrackIndex(strFile As String) As Long
    'This function to check file ? Exists on Database
    'If Exists return index of strFile
    'If No return -1
    Dim x As Long
    Dim s As Long
    s = -1
    For x = 0 To UBound(Library) - 1
        If Len(strFile) = Len(Library(x).Infor.FullName) Then
            If LCase(strFile) = LCase(Library(x).Infor.FullName) Then
                s = x
                Exit For
            End If
        End If
    Next x
    TrackIndex = s
End Function
Public Sub WriteDataFile(Index As Long)
    On Error GoTo handle
    Dim fn As Long
    
    fn = FreeFile
    Open strLibrary For Append As #fn
        Print #fn, Len(Library(Index).Infor.Artist) & _
                    "|" & Len(Library(Index).Infor.Album) & _
                    "|" & Len(Library(Index).Infor.Title) & _
                    "|" & Len(Library(Index).Infor.Genre) & _
                    "|" & Len(Library(Index).Infor.Year) & _
                    "|" & Len(Library(Index).Infor.Filename) & _
                    "|" & Len(Library(Index).Infor.FullName) & _
                    "|" & Len(CStr(Library(Index).Infor.bitrate)) & _
                    "|" & Len(CStr(Library(Index).Infor.Frequency)) & _
                    "|" & Len(CStr(Library(Index).Infor.Duration)) & _
                    "|" & Len(Library(Index).Infor.ExtType) & _
                    "|" & Len(CStr(Library(Index).Infor.Size)) & _
                    "|" & Len(CStr(Library(Index).intPlaycount)) & _
                    "|" & Len(CStr(Library(Index).intRate)) & _
                    "|" & Len(Library(Index).strDay) & _
                    "|" & Len(Library(Index).strDayUpdate)
        Print #fn, Library(Index).Infor.Artist & _
                    Library(Index).Infor.Album & _
                    Library(Index).Infor.Title & _
                    Library(Index).Infor.Genre & _
                    Library(Index).Infor.Year & _
                    Library(Index).Infor.Filename & _
                    Library(Index).Infor.FullName & _
                    Library(Index).Infor.bitrate & _
                    Library(Index).Infor.Frequency & _
                    Library(Index).Infor.Duration & _
                    Library(Index).Infor.ExtType & _
                    Library(Index).Infor.Size & _
                    Library(Index).intPlaycount & _
                    Library(Index).intRate & _
                    Library(Index).strDay & _
                    Library(Index).strDayUpdate & _
                    Library(Index).strType
    Close #fn
    Exit Sub
handle:
    If Err.Number <> 0 Then
        MsgBox "Failed when upload database" & vbCrLf & Err.Number, vbOKOnly, "Database Error"
        Exit Sub
    End If
End Sub

Public Sub WriteDataList(Index As Long)
    On Error GoTo handle
    Dim fn As Long
    
    fn = FreeFile
    
    
    If Mid(Playlist(Index).file, 1, 1) = "," Then Playlist(Index).file = Mid(Playlist(Index).file, 2)
    Open strLibrary For Append As #fn
        Print #fn, "M3PList_" & Playlist(Index).Name & _
                    "#" & Playlist(Index).file
    Close #fn
    Exit Sub
handle:
    If Err.Number <> 0 Then
        MsgBox "Failed when upload database playlist" & Err.Number, vbOKOnly, "Database Error"
        Exit Sub
    End If
    
End Sub

Public Sub LoadDatabase()
    On Error GoTo handle
    
    Dim lngFile As Long
    Dim x As Long
    Dim tTrack As track
    Dim StrTemp As String
    Dim strLen() As String
    Dim intLen(15) As Integer
            
    ReDim Library(0)
    ReDim Playlist(0)
    
    Open strLibrary For Input As #1
        Line Input #1, StrTemp
        If StrTemp <> "[M3P_Library]" Then
            MsgBox "Uncorrect database file" & vbCrLf & "You must restart program and delete 'M3P.db'", vbOKOnly, "M3P_Error"
        End If
        Line Input #1, StrTemp
        lngFile = CInt(Mid(StrTemp, 9))
        If lngFile > 0 Then
            ReDim Library(lngFile)
            Do Until EOF(1)
                Line Input #1, StrTemp
                If InStr(1, StrTemp, "|", vbBinaryCompare) Then
                    strLen = Split(StrTemp, "|")
                    intLen(0) = CInt(strLen(0))
                    intLen(1) = CInt(strLen(1))
                    intLen(2) = CInt(strLen(2))
                    intLen(3) = CInt(strLen(3))
                    intLen(4) = CInt(strLen(4))
                    intLen(5) = CInt(strLen(5))
                    intLen(6) = CInt(strLen(6))
                    intLen(7) = CInt(strLen(7))
                    intLen(8) = CInt(strLen(8))
                    intLen(9) = CInt(strLen(9))
                    intLen(10) = CInt(strLen(10))
                    intLen(11) = CInt(strLen(11))
                    intLen(12) = CInt(strLen(12))
                    intLen(13) = CInt(strLen(13))
                    intLen(14) = CInt(strLen(14))
                    intLen(15) = CInt(strLen(15))
                Else
                    If Mid(StrTemp, 1, 8) <> "M3PList_" Then
                        Library(x).Infor.Artist = Mid(StrTemp, 1, intLen(0))
                        Library(x).Infor.Album = Mid(StrTemp, intLen(0) + 1, intLen(1))
                        Library(x).Infor.Title = Mid(StrTemp, intLen(0) + intLen(1) + 1, intLen(2))
                        Library(x).Infor.Genre = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + 1, intLen(3))
                        Library(x).Infor.Year = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + 1, intLen(4))
                        Library(x).Infor.Filename = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + 1, intLen(5))
                        Library(x).Infor.FullName = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + 1, intLen(6))
                        Library(x).Infor.bitrate = CInt(Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + 1, intLen(7)))
                        Library(x).Infor.Frequency = CLng(Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + 1, intLen(8)))
                        Library(x).Infor.Duration = CLng(Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + 1, intLen(9)))
                        Library(x).Infor.ExtType = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + 1, intLen(10))
                        Library(x).Infor.Size = CLng(Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + intLen(10) + 1, intLen(11)))
                        Library(x).intPlaycount = CInt(Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + intLen(10) + intLen(11) + 1, intLen(12)))
                        Library(x).intRate = CInt(Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + intLen(10) + intLen(11) + intLen(12) + 1, intLen(13)))
                        Library(x).strDay = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + intLen(10) + intLen(11) + intLen(12) + intLen(13) + 1, intLen(14))
                        Library(x).strDayUpdate = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + intLen(10) + intLen(11) + intLen(12) + intLen(13) + intLen(14) + 1, intLen(15))
                        Library(x).strType = Mid(StrTemp, intLen(0) + intLen(1) + intLen(2) + intLen(3) + intLen(4) + intLen(5) + intLen(6) + intLen(7) + intLen(8) + intLen(9) + intLen(10) + intLen(11) + intLen(12) + intLen(13) + intLen(14) + intLen(15) + 1)
                        
                        x = x + 1
                    Else
                        ReDim Preserve Playlist(UBound(Playlist) + 1)
                        Playlist(UBound(Playlist) - 1).Name = Mid(StrTemp, 9, InStr(1, StrTemp, "#", vbBinaryCompare) - 9)
                        Playlist(UBound(Playlist) - 1).file = Mid(StrTemp, Len(Playlist(UBound(Playlist) - 1).Name) + 10)
                    End If
                End If
            Loop
        End If
    Close #1
    Call RefreshLibrary
    Call frmLibrary.ReadLibrary

    Exit Sub
handle:
    If Err.Number <> 0 Then
        Close #1
        MsgBox "LoadDatabase Error : " & Err.Number & " " & Err.Description, vbOKOnly, "Error"
    End If
End Sub
Public Sub RemoveTrack(Index As Long)
    On Error Resume Next
    Dim z As Long
    
    For z = Index To UBound(Library) - 1
        Library(z) = Library(z + 1)
    Next
    ReDim Preserve Library(UBound(Library) - 1)
End Sub
Public Sub RemoveAllTrack()
    On Error Resume Next
    ReDim Library(0)  'cut all record (data  will be lost)
    Call LoadDatabase
End Sub

Public Sub SearchMedia(path As String, Optional bShow As Boolean)
    On Error Resume Next
    
    If bShow = vbNull Then bShow = True
    Load frmSearch
    frmSearch.Visible = bShow
    Wait 100
    
    Dim DirList As New Collection
    Dim FileTemp As New Collection
    Dim temp As String
    Dim strFolder As String
    Dim fFile As file
    
    Screen.MousePointer = vbHourglass
    DirList.Add RepairPath(path, "")
    With frmSearch
        .cmdOk.Enabled = False
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
            Dim i As Long
            Dim ExtName As String
            .File1.path = DirList(1)
            For i = 0 To .File1.ListCount - 1 Step 1
                .File1.ListIndex = i
                If TrackIndex(.File1.path & "\" & .File1.Filename) = -1 Then
                    Set fFile = fso.GetFile(DirList(1) & .File1.Filename)
                    ExtName = LCase(Right(fFile.ShortName, 4))
                    Select Case ExtName
                        Case ".mp3", ".wav", ".ogg", ".wma", ".mp1"
                            If fFile.Size > LibOption.lngAudioSkip * 1024 Then FileTemp.Add fFile.path
                        Case ".aif", ".asf", ".wmv", ".mpg", ".mpe", ".mp2", "mpeg", ".avi", ".mov"
                            If fFile.Size > LibOption.lngVideoSkip * 1024 Then FileTemp.Add fFile.path
                    End Select
                End If
            Next i
            DirList.Remove 1
        Loop
        .prgStatus.max = FileTemp.Count
        .lblSearch.Caption = "Now adding to Library ... "
        .lblStatus.Caption = "Processing ..."
        For i = 0 To FileTemp.Count
            AddNewFile FileTemp(i)
            .prgStatus.value = i
            .lblProgess = CInt((i / .prgStatus.max) * 100) & " %"
            .lblCurrent.Caption = GetShortName(FileTemp(i))
            Wait 1
        Next i
        
        
        Call frmLibrary.ReadLibrary
        .lblCurrent.Caption = ""
        .lblTotal.Caption = .prgStatus.max
        .lblSearch.Caption = "Search completed"
        .lblStatus.Caption = "Completed"
        Screen.MousePointer = vbNormal
        .cmdOk.Enabled = True
    End With
    Set fFile = Nothing
End Sub

Public Sub MonitorFolder()
    On Error Resume Next
    Dim x As Long
    For x = 0 To frmOption.lstLibrary.ListCount - 1
        Call SearchMedia(frmOption.lstLibrary.List(x), False)
    Next x
    Unload frmSearch
End Sub


Public Sub RefreshLibrary()
    On Error Resume Next
    
    Dim x As Long
    Dim tTrack As track
    Dim fso As New FileSystemObject
    Dim fFile As file
    
    For x = LBound(Library) To UBound(Library) - 1
        Set fFile = fso.GetFile(Library(x).Infor.FullName)
        If fFile.DateLastModified > CDate(Library(x).strDayUpdate) Then
            Call OpenTrack(Library(x).Infor.FullName, tTrack)
            Library(x).Infor.Artist = tTrack.Artist
            Library(x).Infor.Album = tTrack.Album
            Library(x).Infor.Duration = tTrack.Duration
            Library(x).Infor.bitrate = tTrack.bitrate
            Library(x).Infor.Frequency = tTrack.Frequency
            Library(x).Infor.Genre = tTrack.Genre
            Library(x).strDayUpdate = date
            Library(x).strDay = date
            Library(x).Infor.Size = tTrack.Size
        End If
    Next x
    Set fFile = Nothing

End Sub




