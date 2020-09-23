Attribute VB_Name = "modSkin"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Declare Function ExtCreateRegionByte Lib "gdi32" Alias "ExtCreateRegion" (lpXform As Long, ByVal nCount As Long, lpRgnData As Byte) As Long

Private Type RegionDataType
    RegionData() As Byte
    DataLength As Long
End Type

Private SkinRegions(1) As RegionDataType


Public PlaylistRad(11) As RECT 'Resize Playlist
Public VideoRad(11) As RECT 'Resize Video
Private lngRegion As Long
Public Sub LoadSkin(NewSkin As String, mini As Boolean)
    On Error GoTo beep
    
    
    Dim DirList As New Collection
    Dim temp As String
    
    If tCurrentSkin.Infor.Name <> "Default" Then
        DirList.Add tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\"
        Do While DirList.Count
            temp = Dir$(DirList(1), vbDirectory)
            Do Until temp = ""
                If temp = "." Or temp = ".." Then
                ElseIf (GetAttr(RepairPath(DirList(1), temp)) And vbDirectory) = vbDirectory Then
                    ElseIf InStr(temp, ".") Then
                    Kill tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\" & temp
                End If
                temp = Dir$
            Loop
            DirList.Remove 1
        Loop
        If Dir(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name, vbDirectory) <> "" Then
            RemoveDirectory tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name
        End If
    End If
    
    Dim strSkin As String   'Skin name
    Dim strSkinconfig As String    'Config file of skin
    Dim picTemp As StdPicture
    Dim i As Long
    
        
    If NewSkin <> "Default.skn" Then
        Call UnZipSkin(tSkinOption.SkinDir & "\" & NewSkin)
    End If
    
    tCurrentSkin.strConfig = tSkinOption.SkinDir & "\" & Mid(NewSkin, 1, Len(NewSkin) - 4) & "\Skin.ini"
    tCurrentSkin.Infor.Name = ReadINI("Infor", "SkinName", tCurrentSkin.strConfig)
    tCurrentSkin.Infor.Comment = ReadINI("Infor", "Comment", tCurrentSkin.strConfig)
    tCurrentSkin.Infor.Location = ReadINI("Infor", "Location", tCurrentSkin.strConfig)
    tCurrentSkin.Infor.Author = ReadINI("Infor", "Author", tCurrentSkin.strConfig)
    
    strSkinconfig = tCurrentSkin.strConfig
    strSkin = tCurrentSkin.Infor.Name
    '[Main]
        With frmMedia
            .Visible = False
            .picMain.Cls
            
            Dim scrX As Long, scrY  As Long
            Dim scrEQX As Long, scrEQY As Long
            
            Set picTemp = LoadPicture(tSkinOption.SkinDir & "\" & strSkin & "\Main.bmp") 'picTempMain.Picture
            'Draw main form
            .picEqua.width = CLng(ReadINI("Main", "EquaWidth", strSkinconfig))
            .picEqua.height = CLng(ReadINI("Main", "EquaHeight", strSkinconfig))
            .picEqua.Left = CLng(ReadINI("Main", "EquaLeft", strSkinconfig))
            .picEqua.Top = CLng(ReadINI("Main", "EquaTop", strSkinconfig))
            
            .picMainMask.Move 0, 0
            .picMainMask.width = CLng(ReadINI("Main", "MaskWidth", strSkinconfig))
            .picMainMask.height = CLng(ReadINI("Main", "MaskHeight", strSkinconfig))
            
            scrX = CLng(ReadINI("Main", "X", tCurrentSkin.strConfig))
            scrY = CLng(ReadINI("Main", "Y", tCurrentSkin.strConfig))
            scrEQX = CLng(ReadINI("Main", "EQX", tCurrentSkin.strConfig))
            scrEQY = CLng(ReadINI("Main", "EQY", tCurrentSkin.strConfig))
            .picMainMask.PaintPicture picTemp, 0, 0, .picMainMask.width, .picMainMask.height, scrX, scrY, .picMainMask.width, .picMainMask.height, vbSrcCopy
            .picEqua.PaintPicture picTemp, 0, 0, .picEqua.width, .picEqua.height, scrEQX, scrEQY, .picEqua.width, .picEqua.height, vbSrcCopy
            '[Infor Title]
                
                .srlInfor.width = ReadINI("Display", "Width", strSkinconfig)
                .srlInfor.height = ReadINI("Display", "Height", strSkinconfig)
                .srlInfor.Left = ReadINI("Display", "Left", strSkinconfig)
                .srlInfor.Top = ReadINI("Display", "Top", strSkinconfig)
                .srlInfor.Font = ReadINI("LabelTitle", "Font", strSkinconfig)
                .srlInfor.FontBold = ReadINI("LabelTitle", "Bold", strSkinconfig)
                .srlInfor.Size = ReadINI("LabelTitle", "Size", strSkinconfig)
                .srlInfor.ForeColor = Hex2VB(ReadINI("LabelTitle", "Color", strSkinconfig))
                .srlInfor.Draw picTemp, CLng(ReadINI("Display", "X", strSkinconfig)), CLng(ReadINI("Display", "Y", strSkinconfig)), .srlInfor.width, .srlInfor.height
                
            '[picSpectrum]
                .Vis.width = ReadINI("Spectrum", "Width", strSkinconfig)
                .Vis.height = ReadINI("Spectrum", "Height", strSkinconfig)
                .Vis.Left = ReadINI("Spectrum", "Left", strSkinconfig)
                .Vis.Top = ReadINI("Spectrum", "Top", strSkinconfig)
                .Vis.Draw picTemp, CLng(ReadINI("Spectrum", "X", strSkinconfig)), CLng(ReadINI("Spectrum", "Y", strSkinconfig)), .Vis.width, .Vis.height
                'Setup Spectrum Color
                .Vis.SpecLowColor = Hex2VB(ReadINI("VisColor", "ColorS1", strSkinconfig))
                .Vis.SpecMidColor = Hex2VB(ReadINI("VisColor", "ColorS2", strSkinconfig))
                .Vis.SpecHiColor = Hex2VB(ReadINI("VisColor", "ColorS3", strSkinconfig))
                .Vis.SpecPeakColor = Hex2VB(ReadINI("VisColor", "ColorPeak", strSkinconfig))
                .Vis.SpecBarPic = tSkinOption.SkinDir & "\" & NewSkin & "\" & ReadINI("VisColor", "Picture", strSkinconfig)
                'Setup Oscillscope Color
                .Vis.OscLowColor = Hex2VB(ReadINI("VisColor", "ColorO1", strSkinconfig))
                .Vis.OscMidColor = Hex2VB(ReadINI("VisColor", "ColorO2", strSkinconfig))
                .Vis.OscHiColor = Hex2VB(ReadINI("VisColor", "ColorO3", strSkinconfig))
                .Vis.setupVisual
            '[Label]
                Call LoadLabel(.lblDuration, "LabelDuration")
                .lblPosition.Caption = "-00:00"
                Call LoadLabel(.lblPosition, "LabelPosition")
                Call LoadLabel(.lblMpgInfo(0), "LabelKbps")
                Call LoadLabel(.lblMpgInfo(1), "LabelKhz")
                
            '[Button]
                Call LoadButton(.btnMedia(0), "Previous", False)
                Call LoadButton(.btnMedia(1), "Play", False)
                Call LoadButton(.btnMedia(2), "Pause", False)
                Call LoadButton(.btnMedia(3), "Next", False)
                Call LoadButton(.btnMedia(4), "Stop", False)
                Call LoadButton(.btnMedia(5), "Shuffe", False)
                Call LoadButton(.btnMedia(6), "Repeat", False)
                Call LoadButton(.btnMedia(7), "ShowEQ", False)
                Call LoadButton(.btnMedia(8), "ShowPL", False)
                Call LoadButton(.btnMedia(9), "ChangeMask", False)
                Call LoadButton(.btnMedia(10), "Minimize", False)
                Call LoadButton(.btnMedia(11), "Close", False)
                Call LoadButton(.btnMedia(12), "EquaEnabled", False)
                Call LoadButton(.btnMedia(13), "EquaPreset", False)
                
            '[Slider] : Position,Volume,Equalizer,Ampli,Balance
                Call LoadSlider(.sldPosition, "Position", False)
                Call LoadSlider(.sldVolume, "Volume", False)
                Call LoadSlider(.sldAmp, "Ampli", False)
                Call LoadSlider(.sldBalance, "Balance", False)
                
                Dim intScale As Integer
                intScale = ReadINI("Equalizer", "Scale", strSkinconfig)
                Call LoadSlider(.sldEqua(0), "Equalizer", False)
                For i = 1 To 9
                    Call LoadSlider(.sldEqua(i), "Equalizer", False)
                    .sldEqua(i).Left = .sldEqua(i - 1).Left + intScale
                Next i
                    
            '[Mini mode]
            Set picTemp = LoadPicture(tSkinOption.SkinDir & "\" & strSkin & "\Mini.bmp") 'picTempMain.Picture
                
            .picMiniEqua.width = CLng(ReadINI("MiniMain", "EquaWidth", strSkinconfig))
            .picMiniEqua.height = CLng(ReadINI("MiniMain", "EquaHeight", strSkinconfig))
            .picMiniEqua.Left = CLng(ReadINI("MiniMain", "EquaLeft", strSkinconfig))
            .picMiniEqua.Top = CLng(ReadINI("MiniMain", "EquaTop", strSkinconfig))
            
            .picMiniMask.Move 0, 0
            .picMiniMask.width = CLng(ReadINI("MiniMain", "MaskWidth", strSkinconfig))
            .picMiniMask.height = CLng(ReadINI("MiniMain", "MaskHeight", strSkinconfig))
            
            scrX = CLng(ReadINI("MiniMain", "X", tCurrentSkin.strConfig))
            scrY = CLng(ReadINI("MiniMain", "Y", tCurrentSkin.strConfig))
            .picMiniEqua.Cls
            .picMiniEqua.PaintPicture picTemp, 0, 0, .picMiniEqua.width, .picMiniEqua.height, .picMiniEqua.Left, .picMiniEqua.Top, .picMiniEqua.width, .picMiniEqua.height
            .picMiniMask.Cls
            .picMiniMask.PaintPicture picTemp, 0, 0, .picMiniMask.width, .picMiniMask.height, scrX, scrY, .picMiniMask.width, .picMiniMask.height
                
            '[Mini Infor Tilte]
                .srlMiniInfor.width = ReadINI("MiniDisplay", "Width", strSkinconfig)
                .srlMiniInfor.height = ReadINI("MiniDisplay", "Height", strSkinconfig)
                .srlMiniInfor.Left = ReadINI("MiniDisplay", "Left", strSkinconfig)
                .srlMiniInfor.Top = ReadINI("MiniDisplay", "Top", strSkinconfig)
                .srlMiniInfor.Font = ReadINI("MiniLabelTitle", "Font", strSkinconfig)
                .srlMiniInfor.FontBold = ReadINI("MiniLabelTitle", "Bold", strSkinconfig)
                .srlMiniInfor.Size = ReadINI("MiniLabelTitle", "Size", strSkinconfig)
                .srlMiniInfor.ForeColor = Hex2VB(ReadINI("MiniLabelTitle", "Color", strSkinconfig))
                .srlMiniInfor.Draw picTemp, CLng(ReadINI("MiniDisplay", "X", strSkinconfig)), CLng(ReadINI("MiniDisplay", "Y", strSkinconfig)), .srlMiniInfor.width, .srlMiniInfor.height
                
            '[VU Meter]
                LoadProgress .prgVU(0), "VULeft"
                LoadProgress .prgVU(1), "VURight"
                
            '[label]
                Call LoadLabel(.lblMiniDur, "MiniLabelDuration")
                .lblMiniPos.Caption = "-00:00"
                Call LoadLabel(.lblMiniPos, "MiniLabelPosition")
                Call LoadLabel(.lblMiniMpgInfo(0), "MiniLabelKbps")
                Call LoadLabel(.lblMiniMpgInfo(1), "MiniLabelKhz")
                
            '[Button]
                Call LoadButton(.btnMiniMedia(0), "MiniPrevious", True)
                Call LoadButton(.btnMiniMedia(1), "MiniPlay", True)
                Call LoadButton(.btnMiniMedia(2), "MiniPause", True)
                Call LoadButton(.btnMiniMedia(3), "MiniNext", True)
                Call LoadButton(.btnMiniMedia(4), "MiniStop", True)
                Call LoadButton(.btnMiniMedia(5), "MiniShuffe", True)
                Call LoadButton(.btnMiniMedia(6), "MiniRepeat", True)
                Call LoadButton(.btnMiniMedia(7), "MiniShowEQ", True)
                Call LoadButton(.btnMiniMedia(8), "MiniShowPL", True)
                Call LoadButton(.btnMiniMedia(9), "MiniChangeMask", True)
                Call LoadButton(.btnMiniMedia(10), "MiniMinimize", True)
                Call LoadButton(.btnMiniMedia(11), "MiniClose", True)
                Call LoadButton(.btnMiniMedia(12), "MiniEquaEnabled", True)
                Call LoadButton(.btnMiniMedia(13), "MiniEquaPreset", True)
                
            '[Slider] : Position,Volume,Equalizer,Ampli,Balance
                Call LoadSlider(.sldMiniPosition, "MiniPosition", True)
                Call LoadSlider(.sldMiniVol, "MiniVolume", True)
                Call LoadSlider(.sldMiniAmp, "MiniAmpli", True)
                Call LoadSlider(.sldMiniBal, "MiniBalance", True)
                
                intScale = ReadINI("MiniEqualizer", "Scale", strSkinconfig)
                Call LoadSlider(.sldMiniEqua(0), "MiniEqualizer", True)
                For i = 1 To 9
                    Call LoadSlider(.sldMiniEqua(i), "MiniEqualizer", True)
                    .sldMiniEqua(i).Left = .sldMiniEqua(i - 1).Left + intScale
                Next i
                    
            Call ChangeMask(mini)
            
            Dim tmp As Boolean
            tmp = tSkinOption.bolEQSlide
            tSkinOption.bolEQSlide = False
            Call ShowEQ
            tSkinOption.bolEQSlide = tmp
        End With
    
    Dim strImgPlaylist As String
    Dim lngColor As Long

    strImgPlaylist = tSkinOption.SkinDir & "\" & strSkin & "\Playlist.bmp"
    With frmPlayList
        
        .Visible = False
        Call LoadButton(.btnPlaylist(0), "Add", False)
        Call LoadButton(.btnPlaylist(1), "Sub", False)
        Call LoadButton(.btnPlaylist(2), "Select", False)
        Call LoadButton(.btnPlaylist(3), "Sort", False)
        Call LoadButton(.btnPlaylist(4), "EditList", False)
        Call LoadButton(.btnPlaylist(5), "ClosePL", False)

        Call LoadSlider(.sldPl, "ScrollList", False)
            
        .picSource.Picture = LoadPicture(strImgPlaylist)
        Call SetImage(PlaylistRad(0), "Playlist", "TopLeft")
        Call SetImage(PlaylistRad(1), "Playlist", "TopRight")
        Call SetImage(PlaylistRad(2), "Playlist", "TopMid")
        Call SetImage(PlaylistRad(3), "Playlist", "TopLeftRez")
        Call SetImage(PlaylistRad(4), "Playlist", "TopRightRez")
        Call SetImage(PlaylistRad(5), "Playlist", "BottomLeft")
        Call SetImage(PlaylistRad(6), "Playlist", "BottomRight")
        Call SetImage(PlaylistRad(7), "Playlist", "BottomMid")
        Call SetImage(PlaylistRad(8), "Playlist", "BottomLeftRez")
        Call SetImage(PlaylistRad(9), "Playlist", "BottomRightRez")
        Call SetImage(PlaylistRad(10), "Playlist", "MidLeft")
        Call SetImage(PlaylistRad(11), "Playlist", "MidRight")
        
        tPlaylistConfig.lngBackColor = Hex2VB(ReadINI("Playlist", "BGColor", strSkinconfig))
        tPlaylistConfig.lngForeColor = Hex2VB(ReadINI("Playlist", "FGColor", strSkinconfig))
        tPlaylistConfig.lngPlayColor = Hex2VB(ReadINI("Playlist", "PlayingColor", strSkinconfig))
        tPlaylistConfig.lngPlayBackColor = Hex2VB(ReadINI("Playlist", "PlayBackColor", strSkinconfig))
        tPlaylistConfig.bolPlayBold = ReadINI("Playlist", "PlayBold", strSkinconfig, True)
        tPlaylistConfig.lngSelectedColor = Hex2VB(ReadINI("Playlist", "SelectedColor", strSkinconfig))
        tPlaylistConfig.lngSelectedBorderColor = Hex2VB(ReadINI("Playlist", "SelectedBorderColor", strSkinconfig))
        tPlaylistConfig.lngFontSize = ReadINI("Playlist", "FontSize", strSkinconfig)
        tPlaylistConfig.strFontName = ReadINI("Playlist", "FontName", strSkinconfig)
            
        .List.Font.Name = tPlaylistConfig.strFontName
        .List.Font.Size = tPlaylistConfig.lngFontSize
        .List.BackColor = tPlaylistConfig.lngBackColor
        .List.ForeColor = tPlaylistConfig.lngForeColor
        .List.SelectColor = tPlaylistConfig.lngSelectedColor
        .List.SelectBorderColor = tPlaylistConfig.lngSelectedBorderColor
        .List.PlayBold = tPlaylistConfig.bolPlayBold
        .List.PlayColor = tPlaylistConfig.lngPlayBackColor
        .List.PlayForeColor = tPlaylistConfig.lngPlayColor
        
        Dim tmpPic As String
        tmpPic = tSkinOption.SkinDir & "\" & strSkin & "\" & ReadINI("Playlist", "Picture", strSkinconfig, 0)
        If FileExists(tmpPic) Then
            .List.PicAlign = ReadINI("Playlist", "PictureAlignment", strSkinconfig, 0)
            .List.Picture = tmpPic
        Else
            .List.Picture = ""
        End If
        
        Call LoadLabel(.lblCurrentperTotal, "LabelCurrentTotal")
        Call LoadLabel(.lblTotalTime, "LabelTotalTime")
        
        .width = CLng(ReadINI("Playlist", "Width", strSkinconfig)) * Screen.TwipsPerPixelX
        .height = CLng(ReadINI("Playlist", "Height", strSkinconfig)) * Screen.TwipsPerPixelY
        .width = ReadINI("Demension", "PlaylistWidth", strFileconfig)
        .height = ReadINI("Demension", "PlaylistHeight", strFileconfig)
        .Left = ReadINI("Demension", "PlaylistLeft", strFileconfig)
        .Top = ReadINI("Demension", "PlaylistTop", strFileconfig)
        .Refresh
        .Visible = Not tPlaylistConfig.bolHidePL
        Call SnapForm(frmMedia, frmPlayList)
    End With
            
    Dim strImgVideo As String
    Dim tmpDefault As Boolean
    
    strImgVideo = tSkinOption.SkinDir & "\" & strSkin & "\Video.bmp"
    
    With frmVD
        .Visible = False
        If FileExists(strImgVideo) = False Then
            strImgVideo = tSkinOption.SkinDir & "\Default\Video.bmp"
            tmpDefault = True
        Else
            tmpDefault = False
        End If
        
        Call LoadButton(.btnVideo(0), "1X", False, tmpDefault)
        Call LoadButton(.btnVideo(1), "2X", False, tmpDefault)
        Call LoadButton(.btnVideo(2), "FS", False, tmpDefault)
        Call LoadButton(.btnVideo(3), "CloseVideo", False, tmpDefault)
            
        .picSource.Picture = LoadPicture(strImgVideo)
        Call SetImage(VideoRad(0), "Video", "TopLeft", tmpDefault)
        Call SetImage(VideoRad(1), "Video", "TopRight", tmpDefault)
        Call SetImage(VideoRad(2), "Video", "TopMid", tmpDefault)
        Call SetImage(VideoRad(3), "Video", "TopLeftRez", tmpDefault)
        Call SetImage(VideoRad(4), "Video", "TopRightRez", tmpDefault)
        Call SetImage(VideoRad(5), "Video", "BottomLeft", tmpDefault)
        Call SetImage(VideoRad(6), "Video", "BottomRight", tmpDefault)
        Call SetImage(VideoRad(7), "Video", "BottomMid", tmpDefault)
        Call SetImage(VideoRad(8), "Video", "BottomLeftRez", tmpDefault)
        Call SetImage(VideoRad(9), "Video", "BottomRightRez", tmpDefault)
        Call SetImage(VideoRad(10), "Video", "MidLeft", tmpDefault)
        Call SetImage(VideoRad(11), "Video", "MidRight", tmpDefault)
        
        If bolVideoOn Then .Visible = True
    
    End With
        
    Dim strImgMenu As String
    Dim strImgMenuSel As String
    
    strImgMenu = tSkinOption.SkinDir & "\" & strSkin & "\Menu.bmp"
    strImgMenuSel = tSkinOption.SkinDir & "\" & strSkin & "\MenuSelect.bmp"
    
    If FileExists(strImgMenu) = False Then
        strImgMenu = tSkinOption.SkinDir & "\Default\Menu.bmp"
    End If
    
    With CustomMenu
        Set .Picture = LoadPicture(strImgMenu)
        If FileExists(strImgMenuSel) = True Then
            Set .HightLightPicture = LoadPicture(strImgMenuSel)
        Else
            Set .HightLightPicture = Nothing
        End If
        .PosX = CLng(ReadINI("Menu", "X", strSkinconfig))
    End With
    
    With CustomColor
        .BackColor = Hex2VB(ReadINI("Menu", "BackColor", strSkinconfig))
        .BarColor = Hex2VB(ReadINI("Menu", "BarColor", strSkinconfig))
        .BorderColor = Hex2VB(ReadINI("Menu", "BorderColor", strSkinconfig))
        .DefTextColor = Hex2VB(ReadINI("Menu", "DefTextColor", strSkinconfig))
        .ForeColor = Hex2VB(ReadINI("Menu", "ForeColor", strSkinconfig))
        .HilightColor = Hex2VB(ReadINI("Menu", "HilightColor", strSkinconfig))
        .HilightBorderColor = Hex2VB(ReadINI("Menu", "HilightBorderColor", strSkinconfig))
        .MenuTextColor = Hex2VB(ReadINI("Menu", "MenuTextColor", strSkinconfig))
        .NormalColor = Hex2VB(ReadINI("Menu", "NormalColor", strSkinconfig))
        .SelectedTextColor = Hex2VB(ReadINI("Menu", "SelectedTextColor", strSkinconfig))
    End With
    
    For i = 0 To frmMenu.mnuSkinC.Count - 1
        frmMenu.mnuSkinC(i).Checked = False
        If frmMenu.mnuSkinC(i).Caption = strSkin Then frmMenu.mnuSkinC(i).Checked = True
        If bolLoading = False Then
            If frmOption.lstSkin.List(i) = strSkin Then frmOption.lstSkin.Selected(i) = True
        End If
    Next i

    Exit Sub
beep:
    If Err.Number <> 0 Then
        Debug.Print "LoadSkin Error " & Err.Description
        If tCurrentSkin.Infor.Name <> "Default" Then
            Call LoadSkin("Default.skn", False)
            
            For i = 4 To frmMenu.mnuSkinC.Count - 1
                Unload frmMenu.mnuSkinC(i)
            Next i
            Call frmMenu.Add_SkinInstall(tSkinOption.SkinDir)
            
            For i = 0 To frmMenu.mnuSkinC.Count - 1
                frmMenu.mnuSkinC(i).Checked = False
                If frmMenu.mnuSkinC(i).Caption = strSkin Then frmMenu.mnuSkinC(i).Checked = True
            Next i
            
            frmOption.lstSkin.Clear
            For i = 3 To frmMenu.mnuSkinC.Count - 1
                frmOption.lstSkin.AddItem frmMenu.mnuSkinC(i).Caption
            Next i
            
        Else
            MsgBox "You default skin has an error" & vbCrLf & "Please select orther skin", vbInformation, "Skin Error"
            Exit Sub
        End If
    End If
End Sub
Public Sub LoadLabel(lblLabel As Label, Section As String, Optional Default As Boolean)
    On Error GoTo beep
    Dim strSkinconfig As String
    
    If Default = False Then
        strSkinconfig = tCurrentSkin.strConfig
    Else
        strSkinconfig = tSkinOption.SkinDir & "\Default\Default.ini"
    End If
    
    lblLabel.Font = ReadINI(Section, "Font", strSkinconfig)
    lblLabel.FontSize = ReadINI(Section, "Size", strSkinconfig)
    lblLabel.FontBold = ReadINI(Section, "Bold", strSkinconfig)
    lblLabel.ForeColor = Hex2VB(ReadINI(Section, "Color", strSkinconfig))
    lblLabel.Left = ReadINI(Section, "Left", strSkinconfig)
    lblLabel.Top = ReadINI(Section, "Top", strSkinconfig)
beep:
    If Err.Number <> 0 Then
        Debug.Print Section
        lblLabel.Font = "Tahoma"
        lblLabel.FontSize = 8
        lblLabel.FontBold = True
        lblLabel.ForeColor = RGB(0, 0, 0)
        lblLabel.Left = 0
        lblLabel.Top = 0
    End If
End Sub
Public Sub LoadButton(Button As Object, Section As String, mini As Boolean, Optional Default As Boolean)
    On Error GoTo beep
    Dim strSkinconfig As String
    Dim pic As StdPicture
    Dim picStateButton(5) As StdPicture
    Dim scrX(5) As Long
    Dim scrY(5) As Long
    If Default = False Then
        If mini = False Then
            Set pic = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\ButtonGroup.bmp")
        Else
            Set pic = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\MiniButtonGroup.bmp")
        End If
        strSkinconfig = tCurrentSkin.strConfig
    Else
        If mini = False Then
            Set pic = LoadPicture(tSkinOption.SkinDir & "\Default\ButtonGroup.bmp")
        Else
            Set pic = LoadPicture(tSkinOption.SkinDir & "\Default\MiniButtonGroup.bmp")
        End If
        strSkinconfig = tSkinOption.SkinDir & "\Default\Default.ini"
    End If
    
    Button.Left = CLng(ReadINI(Section, "Left", strSkinconfig))
    Button.Top = CLng(ReadINI(Section, "Top", strSkinconfig))
    frmMedia.picButton.height = CLng(ReadINI(Section, "Height", strSkinconfig))
    frmMedia.picButton.width = CLng(ReadINI(Section, "Width", strSkinconfig))
    
    scrX(0) = CLng(ReadINI(Section, "X", strSkinconfig))
    scrX(1) = CLng(ReadINI(Section, "XDown", strSkinconfig))
    scrX(2) = CLng(ReadINI(Section, "XOver", strSkinconfig))
    scrX(3) = CLng(ReadINI(Section, "XOn", strSkinconfig))
    scrX(4) = CLng(ReadINI(Section, "XOnDown", strSkinconfig))
    scrX(5) = CLng(ReadINI(Section, "XOnOver", strSkinconfig))
    
    scrY(0) = CLng(ReadINI(Section, "Y", strSkinconfig))
    scrY(1) = CLng(ReadINI(Section, "YDown", strSkinconfig))
    scrY(2) = CLng(ReadINI(Section, "YOver", strSkinconfig))
    scrY(3) = CLng(ReadINI(Section, "YOn", strSkinconfig))
    scrY(4) = CLng(ReadINI(Section, "YOnDown", strSkinconfig))
    scrY(5) = CLng(ReadINI(Section, "YOnOver", strSkinconfig))
    
    
    Dim i As Integer
    For i = 0 To 5
        With frmMedia
            .picButton.Cls
            .picButton.PaintPicture pic, 0, 0, .picButton.width, .picButton.height, scrX(i), scrY(i), .picButton.width, .picButton.height
            .picButton.Picture = .picButton.Image
            Set picStateButton(i) = .picButton.Picture
        End With
    Next i
    With Button
        Set .Picture = picStateButton(0)
        Set .PictureDown = picStateButton(1)
        Set .PictureOver = picStateButton(2)
        Set .PictureOn = picStateButton(3)
        Set .PictureOnDown = picStateButton(4)
        Set .PictureOnOver = picStateButton(5)
    End With
    
    frmMedia.picButton.Cls
    Set pic = Nothing
    For i = 0 To 5
        Set picStateButton(i) = Nothing
    Next i
beep:
    If Err.Number <> 0 Then
        Debug.Print Section
        Exit Sub
    End If
End Sub

Public Sub LoadSlider(Slider As Object, Section As String, mini As Boolean)
    On Error GoTo beep
    Dim strSkinconfig As String    'Config file of skin
    Dim pic As StdPicture
    Dim picStateSlider(4) As StdPicture
    Dim scrX(4) As Long
    Dim scrY(4) As Long
    Dim Style As String
    Dim Orien As Byte
    Dim W As Long, H As Long
    Dim WCue As Long, HCue As Long
    
    If mini = False Then
        Set pic = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\ButtonGroup.bmp")
    Else
        Set pic = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\MiniButtonGroup.bmp")
    End If
    strSkinconfig = tCurrentSkin.strConfig
    
    Slider.Left = CLng(ReadINI(Section, "Left", strSkinconfig))
    Slider.Top = CLng(ReadINI(Section, "Top", strSkinconfig))
    H = CLng(ReadINI(Section, "Height", strSkinconfig))
    W = CLng(ReadINI(Section, "Width", strSkinconfig))
    HCue = CLng(ReadINI(Section, "CueHeight", strSkinconfig))
    WCue = CLng(ReadINI(Section, "CueWidth", strSkinconfig))
    
    Slider.width = W
    Slider.height = H
    
    Orien = CByte(ReadINI(Section, "Orien", strSkinconfig))
    Slider.SliderOrien = Orien
    
    Style = ReadINI(Section, "Style", strSkinconfig)
        If Style = "Graphic" Then
             Slider.Style = 1
            
             scrX(0) = CLng(ReadINI(Section, "X", strSkinconfig))
             scrX(1) = CLng(ReadINI(Section, "XOver", strSkinconfig))
             scrX(2) = CLng(ReadINI(Section, "CueX", strSkinconfig))
             scrX(3) = CLng(ReadINI(Section, "CueXDown", strSkinconfig))
             scrX(4) = CLng(ReadINI(Section, "CueXOver", strSkinconfig))
             
             scrY(0) = CLng(ReadINI(Section, "Y", strSkinconfig))
             scrY(1) = CLng(ReadINI(Section, "YOver", strSkinconfig))
             scrY(2) = CLng(ReadINI(Section, "CueY", strSkinconfig))
             scrY(3) = CLng(ReadINI(Section, "CueYDown", strSkinconfig))
             scrY(4) = CLng(ReadINI(Section, "CueYOver", strSkinconfig))
             
             
             Dim i As Integer
             For i = 0 To 1
                 With frmMedia
                     .picButton.Cls
                     .picButton.height = H
                     .picButton.width = W
                     .picButton.PaintPicture pic, 0, 0, W, H, scrX(i), scrY(i), W, H
                     .picButton.Picture = .picButton.Image
                     Set picStateSlider(i) = .picButton.Picture
                 End With
             Next i
             For i = 2 To 4
                 With frmMedia
                     .picButton.Cls
                     .picButton.height = HCue
                     .picButton.width = WCue
                     .picButton.PaintPicture pic, 0, 0, WCue, HCue, scrX(i), scrY(i), WCue, HCue
                     .picButton.Picture = .picButton.Image
                     Set picStateSlider(i) = .picButton.Picture
                 End With
             Next i
             With Slider
                 Set .Picture = picStateSlider(0)
                 Set .PictureOver = picStateSlider(1)
                 Set .PictureCue = picStateSlider(2)
                 Set .PictureCueDown = picStateSlider(3)
                 Set .PictureCueOver = picStateSlider(4)
             End With
             
             frmMedia.picButton.Cls
             Set pic = Nothing
             For i = 0 To 4
                 Set picStateSlider(i) = Nothing
             Next i
        Else
            Slider.Style = 0
            Slider.BackColor = Hex2VB(ReadINI(Section, "BackColor", strSkinconfig))
            Slider.ForeColor = Hex2VB(ReadINI(Section, "ForeColor", strSkinconfig))
        End If
beep:
    If Err.Number <> 0 Then
        Debug.Print Section
    End If
End Sub
Public Sub LoadProgress(ProgressBar As Object, Section As String)
    On Error Resume Next
    Dim strSkinconfig As String    'Config file of skin
    Dim pic As StdPicture
    Dim picState(4) As StdPicture
    Dim Style As String
    Dim scrX(4) As Long
    Dim scrY(4) As Long
    Dim W As Long, H As Long
    
    Set pic = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\MiniButtonGroup.bmp")
    strSkinconfig = tCurrentSkin.strConfig

    ProgressBar.Left = CLng(ReadINI(Section, "Left", strSkinconfig))
    ProgressBar.Top = CLng(ReadINI(Section, "Top", strSkinconfig))
    H = CLng(ReadINI(Section, "Height", strSkinconfig))
    W = CLng(ReadINI(Section, "Width", strSkinconfig))
    
    Style = ReadINI(Section, "Style", strSkinconfig)
    If Style = "Graphic" Then
            ProgressBar.ProgressStyle = 2
            scrX(0) = CLng(ReadINI(Section, "X", strSkinconfig))
            scrX(1) = CLng(ReadINI(Section, "XOver", strSkinconfig))
            scrX(2) = CLng(ReadINI(Section, "StatusX", strSkinconfig))
            scrX(3) = CLng(ReadINI(Section, "StatusXDown", strSkinconfig))
            scrX(4) = CLng(ReadINI(Section, "StatusXOver", strSkinconfig))
            
            scrY(0) = CLng(ReadINI(Section, "Y", strSkinconfig))
            scrY(1) = CLng(ReadINI(Section, "YOver", strSkinconfig))
            scrY(2) = CLng(ReadINI(Section, "StatusY", strSkinconfig))
            scrY(3) = CLng(ReadINI(Section, "StatusYDown", strSkinconfig))
            scrY(4) = CLng(ReadINI(Section, "StatusYOver", strSkinconfig))
            
            
            Dim i As Integer
            For i = 0 To 4
                With frmMedia
                    .picButton.Cls
                    .picButton.height = H
                    .picButton.width = W
                    .picButton.PaintPicture pic, 0, 0, W, H, scrX(i), scrY(i), W, H
                    .picButton.Picture = .picButton.Image
                    Set picState(i) = .picButton.Picture
                End With
            Next i
            With ProgressBar
                Set .Picture = picState(0)
                Set .PictureOver = picState(1)
                Set .PictureStatus = picState(2)
                Set .PictureStatusDown = picState(3)
                Set .PictureStatusOver = picState(4)
                .width = W
                .height = H
            End With
            
            frmMedia.picButton.Cls
            Set pic = Nothing
            For i = 0 To 4
                Set picState(i) = Nothing
            Next i
    Else
    
        ProgressBar.ProgressStyle = 0
        ProgressBar.width = W
        ProgressBar.height = H
        ProgressBar.BackColor = Hex2VB(ReadINI(Section, "BackColor", strSkinconfig))
        ProgressBar.ForeColor = Hex2VB(ReadINI(Section, "ForeColor", strSkinconfig))
    End If
End Sub
Public Sub SetImage(rad As RECT, Section As String, Key As String, Optional Default As Boolean)
    Dim arrNumber() As String
    Dim Des(3) As Long
    Dim strSkinconfig As String
    Dim i As Integer
    
    If Default Then
        strSkinconfig = tSkinOption.SkinDir & "\Default\Default.ini"
    Else
        strSkinconfig = tCurrentSkin.strConfig
    End If
    arrNumber = Split(ReadINI(Section, Key, strSkinconfig), ",")
    For i = 0 To 3
        Des(i) = arrNumber(i)
    Next i
    rad.Bottom = Des(0)
    rad.Right = Des(1)
    rad.Left = Des(2)
    rad.Top = Des(3)
End Sub
Public Function MakeRegion(picSoucre As Object, lngTransparentColor As Long) As Long
    Dim x As Long
    Dim y As Long
    Dim StartLineX As Long
    Dim LineRegion As Long
    Dim FinishRegion As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hDC As Long
    Dim lWidth As Long
    Dim lHeight As Long
    
    hDC = picSoucre.hDC
    lWidth = picSoucre.ScaleWidth
    lHeight = picSoucre.ScaleHeight
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    For y = 0 To lHeight
        For x = 0 To lWidth
            If GetPixel(hDC, x, y) = lngTransparentColor Or x = lWidth Then
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    If InFirstRegion Then
                        FinishRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FinishRegion, FinishRegion, LineRegion, RGN_OR
                        DeleteObject LineRegion
                    End If
                End If
            Else
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    MakeRegion = FinishRegion
End Function
Public Sub ShowEQ()
    Dim mH As Long, mW As Long, W As Long, H As Long
    Dim ELeft As Long, ETop As Long
    
    Dim EQLeft As Long
    Dim EQTop As Long
    Dim MainWidth As Long
    Dim MainHeight As Long
    With frmMedia
        .Enabled = False
        If tCurrentSkin.mini = False Then
            H = CLng(ReadINI("Main", "MaxHeight", tCurrentSkin.strConfig))
            W = CLng(ReadINI("Main", "MaxWidth", tCurrentSkin.strConfig))
            mH = CLng(ReadINI("Main", "MinHeight", tCurrentSkin.strConfig))
            mW = CLng(ReadINI("Main", "MinWidth", tCurrentSkin.strConfig))
            ELeft = CLng(ReadINI("Main", "EquaLeft", tCurrentSkin.strConfig))
            ETop = CLng(ReadINI("Main", "EquaTop", tCurrentSkin.strConfig))
            If mW <> W Then
                EQLeft = .picEqua.Left
                MainWidth = .width
                If Not tPlayerConfig.bolShowEQ Then
                    If tSkinOption.bolEQSlide Then
                        Do While EQLeft > mW - .picEqua.width
                            EQLeft = EQLeft - 1
                            MainWidth = MainWidth - 1 * Screen.TwipsPerPixelX
                            Wait 10
                            .width = MainWidth
                            .picEqua.Left = EQLeft
                        Loop
                        If .ScaleWidth > mW Then .width = mW * Screen.TwipsPerPixelX
                    Else
                        .picEqua.Left = mW - .picEqua.width
                        .width = mW * Screen.TwipsPerPixelX
                    End If
                Else
                    If tSkinOption.bolEQSlide Then
                        Do While .width < W * Screen.TwipsPerPixelX
                            MainWidth = MainWidth + 1 * Screen.TwipsPerPixelX
                            EQLeft = EQLeft + 1
                            Wait 1
                            .width = MainWidth
                            .picEqua.Left = EQLeft
                        Loop
                        If .ScaleWidth > W Then .width = W * Screen.TwipsPerPixelX
                    Else
                        .width = W * Screen.TwipsPerPixelX
                        .picEqua.Left = ELeft
                    End If
                End If
            End If
            If mH <> H Then
                EQTop = .picEqua.Top
                MainHeight = .height
                If Not tPlayerConfig.bolShowEQ Then
                    If tSkinOption.bolEQSlide Then
                        Do While EQTop > mH - .picEqua.height
                            MainHeight = MainHeight - 1 * Screen.TwipsPerPixelY
                            EQTop = EQTop - 1
                            Wait 10
                            .height = MainHeight
                            .picEqua.Top = EQTop
                        Loop
                        If .ScaleHeight > mH Then .height = mH * Screen.TwipsPerPixelY
                    Else
                        .height = mH * Screen.TwipsPerPixelY
                        .picEqua.Top = mH - .picEqua.height
                    End If
                Else
                    If tSkinOption.bolEQSlide Then
                        Do While MainHeight < H * Screen.TwipsPerPixelY
                            MainHeight = MainHeight + 1 * Screen.TwipsPerPixelY
                            EQTop = EQTop + 1
                            Wait 10
                            .height = MainHeight
                            .picEqua.Top = EQTop
                        Loop
                        If .ScaleHeight > H Then .height = H * Screen.TwipsPerPixelY
                    Else
                        .picEqua.Top = ETop
                        .height = H * Screen.TwipsPerPixelY
                    End If
                End If
            End If
        Else
            mH = CLng(ReadINI("MiniMain", "MinHeight", tCurrentSkin.strConfig))
            H = CLng(ReadINI("MiniMain", "MaxHeight", tCurrentSkin.strConfig))
            ETop = CLng(ReadINI("MiniMain", "EquaTop", tCurrentSkin.strConfig))
            EQTop = .picMiniEqua.Top
            MainHeight = .height
            If Not tPlayerConfig.bolShowEQ Then
                If tSkinOption.bolEQSlide Then
                    Do While EQTop > mH - .picMiniEqua.height
                        MainHeight = MainHeight - 1 * Screen.TwipsPerPixelY
                        EQTop = EQTop - 1
                        Wait 10
                        .height = MainHeight
                        .picMiniEqua.Top = EQTop
                    Loop
                    If .ScaleHeight > mH Then .height = mH * Screen.TwipsPerPixelY
                Else
                    .picMiniEqua.Top = mH - .picMiniEqua.height
                    .height = mH * Screen.TwipsPerPixelY
                End If
            Else
                If tSkinOption.bolEQSlide Then
                    Do While MainHeight < H * Screen.TwipsPerPixelY
                        MainHeight = MainHeight + 1 * Screen.TwipsPerPixelY
                        EQTop = EQTop + 1
                        Wait 10
                        .picMiniEqua.Top = EQTop
                        .height = MainHeight
                    Loop
                    If .ScaleHeight > H Then .height = H * Screen.TwipsPerPixelY
                Else
                    .height = H * Screen.TwipsPerPixelY
                    .picMiniEqua.Top = ETop
                End If
            End If
        End If
        .Enabled = True
    End With
    InvalidateRect 0&, 0&, False
End Sub
Public Sub CenterForm(frm As Form)
    frm.Left = (Screen.width - frm.width) / 2
    frm.Top = (Screen.height - frm.height) / 2
End Sub
Public Sub ChangeMask(mini As Boolean)
    Dim scrX As Long, scrY  As Long
    Dim picTemp As StdPicture
    Dim i As Integer
    Dim bolHasTransDynamic As Boolean
    
    tCurrentSkin.mini = mini
    With frmMedia
        .Visible = False
        
        If mini = False Then
            Set picTemp = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Main.bmp")
            .picMain.Cls
            .picMini.Visible = False
            .picMain.Move 0, 0
            .picMain.width = CLng(ReadINI("Main", "MaxWidth", tCurrentSkin.strConfig))
            .picMain.height = CLng(ReadINI("Main", "MaxHeight", tCurrentSkin.strConfig))
            .height = .picMain.height * Screen.TwipsPerPixelX
            .width = .picMain.width * Screen.TwipsPerPixelY
            scrX = CLng(ReadINI("Main", "MaskX", tCurrentSkin.strConfig))
            scrY = CLng(ReadINI("Main", "MaskY", tCurrentSkin.strConfig))
            .picMain.PaintPicture picTemp, 0, 0, .picMain.width, .picMain.height, scrX, scrY, .picMain.width, .picMain.height
            If LoadRegions(SkinRegions()) Then
                lngRegion = ExtCreateRegionByte(ByVal 0&, SkinRegions(0).DataLength, SkinRegions(0).RegionData(0))
            Else
                lngRegion = MakeRegion(.picMain, Hex2VB(ReadINI("Main", "TransColor", tCurrentSkin.strConfig)))
            End If
            SetWindowRgn .hwnd, lngRegion, True
            DeleteObject lngRegion
            bolHasTransDynamic = ReadINI("Main", "HasTransDynamic", tCurrentSkin.strConfig, True)
            If bolHasTransDynamic Then
                WinAPI.Make_TransColor .hwnd, Hex2VB(ReadINI("Main", "TransColorDynamic", tCurrentSkin.strConfig))
            End If
            scrX = CLng(ReadINI("Main", "X", tCurrentSkin.strConfig))
            scrY = CLng(ReadINI("Main", "Y", tCurrentSkin.strConfig))
            .picMain.PaintPicture picTemp, 0, 0, .picMain.width, .picMain.height, scrX, scrY, .picMain.width, .picMain.height, vbSrcCopy
            .picMain.Visible = True
        Else
            Set picTemp = LoadPicture(tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\Mini.bmp")
            .picMini.Cls
            .picMain.Visible = False
            .picMini.Move 0, 0
            .picMini.width = CLng(ReadINI("MiniMain", "MaxWidth", tCurrentSkin.strConfig))
            .picMini.height = CLng(ReadINI("MiniMain", "MaxHeight", tCurrentSkin.strConfig))
            .height = .picMini.height * Screen.TwipsPerPixelX
            .width = .picMini.width * Screen.TwipsPerPixelY
            scrX = CLng(ReadINI("MiniMain", "MaskX", tCurrentSkin.strConfig))
            scrY = CLng(ReadINI("MiniMain", "MaskY", tCurrentSkin.strConfig))
            .picMini.PaintPicture picTemp, 0, 0, .picMini.width, .picMini.height, scrX, scrY, .picMini.width, .picMini.height
            If LoadRegions(SkinRegions()) Then
                lngRegion = ExtCreateRegionByte(ByVal 0&, SkinRegions(1).DataLength, SkinRegions(1).RegionData(0))
            Else
                lngRegion = MakeRegion(.picMini, Hex2VB(ReadINI("MiniMain", "TransColor", tCurrentSkin.strConfig)))
            End If
            SetWindowRgn .hwnd, lngRegion, True
            DeleteObject lngRegion
            bolHasTransDynamic = ReadINI("MiniMain", "HasTransDynamic", tCurrentSkin.strConfig, True)
            If bolHasTransDynamic Then
                WinAPI.Make_TransColor frmMedia.hwnd, Hex2VB(ReadINI("MiniMain", "TransColorDynamic", tCurrentSkin.strConfig))
            End If
            scrX = CLng(ReadINI("MiniMain", "X", tCurrentSkin.strConfig))
            scrY = CLng(ReadINI("MiniMain", "Y", tCurrentSkin.strConfig))
            .picMini.PaintPicture picTemp, 0, 0, .picMini.width, .picMini.height, scrX, scrY, .picMini.width, .picMini.height, vbSrcCopy
            .picMini.Visible = True
        End If
        'Refresh button
        .btnMedia(5).bolOn = tPlayerConfig.bolShuffe
        .btnMedia(6).bolOn = tPlayerConfig.bolRepeat
        .btnMedia(7).bolOn = tPlayerConfig.bolShowEQ
        .btnMedia(8).bolOn = Not tPlaylistConfig.bolHidePL
        .btnMedia(12).bolOn = bolEQEnabled
        .btnMiniMedia(5).bolOn = tPlayerConfig.bolShuffe
        .btnMiniMedia(6).bolOn = tPlayerConfig.bolRepeat
        .btnMiniMedia(7).bolOn = tPlayerConfig.bolShowEQ
        .btnMiniMedia(8).bolOn = Not tPlaylistConfig.bolHidePL
        .btnMiniMedia(12).bolOn = bolEQEnabled
        For i = 0 To .btnMedia.Count - 1
            .btnMedia(i).Refresh
            .btnMiniMedia(i).Refresh
        Next i
        .Enabled = True
        .Visible = True
    End With
End Sub
Public Sub DragForm(frm As Form)
    On Error Resume Next
    ReleaseCapture
    SendMessage frm.hwnd, &HA1, 2, 0&
End Sub

Public Sub SnapForm(frmParent As Form, frmChild As Form)
    On Error Resume Next
    If frmChild.Left - (frmParent.Left + frmParent.width) < 500 And frmChild.Left - (frmParent.Left + frmParent.width) >= 0 Then
        frmChild.Left = frmParent.Left + frmParent.width
    End If
    If frmParent.Left - (frmChild.Left + frmChild.width) < 500 And frmParent.Left - (frmChild.Left + frmChild.width) >= 0 Then
        frmChild.Left = frmParent.Left - frmChild.width
    End If
    If frmChild.Left - frmParent.Left < 500 And frmChild.Left - frmParent.Left > 0 Then
        frmChild.Left = frmParent.Left
    End If
    If frmChild.Left - frmParent.Left > -500 And frmChild.Left - frmParent.Left < 0 Then
        frmChild.Left = frmParent.Left
    End If
    If frmChild.Left - (frmParent.Left + frmParent.width) > -500 And frmChild.Left - (frmParent.Left + frmParent.width) < 0 Then
        frmChild.Left = frmParent.Left + frmParent.width
    End If
    If frmChild.Top - (frmParent.height + frmParent.Top) < 500 And frmChild.Top - (frmParent.height + frmParent.Top) >= 0 Then
        frmChild.Top = frmParent.Top + frmParent.height
    End If
    If frmParent.Top - (frmChild.Top + frmChild.height) < 500 And frmParent.Top - (frmChild.Top + frmChild.height) >= 0 Then
        frmChild.Top = frmParent.Top - frmChild.height
    End If
    If frmChild.Top - frmParent.Top < 500 And frmChild.Top - frmParent.Top > 0 Then
        frmChild.Top = frmParent.Top
    End If
    If frmChild.Top - frmParent.Top > -500 And frmChild.Top - frmParent.Top < 0 Then
        frmChild.Top = frmParent.Top
    End If
    If frmChild.Top - (frmParent.height + frmParent.Top) > -500 And frmChild.Top - (frmParent.height + frmParent.Top) < 0 Then
        frmChild.Top = frmParent.Top + frmParent.height
    End If
    InvalidateRect 0&, ByVal 0, 1&
End Sub

Public Sub MoveSnapForm(oLeft As Long, oTop As Long)
    On Error Resume Next
    With frmMedia
            If tPlaylistConfig.bolHidePL = False Then
                frmPlayList.Left = frmPlayList.Left + (.Left - oLeft)
                frmPlayList.Top = frmPlayList.Top + (.Top - oTop)
            End If
            If frmVisual.bolShow Then
                frmVisual.Left = frmVisual.Left + (.Left - oLeft)
                frmVisual.Top = frmVisual.Top + (.Top - oTop)
            End If
            If (frmVD.bolShow And frmVD.Video.VideoMode = Window) = True Then
                frmVD.Left = frmVD.Left + (.Left - oLeft)
                frmVD.Top = frmVD.Top + (.Top - oTop)
            End If
    End With
End Sub
Public Function ManualShowEQ(mini As Boolean, oldX As Single, oldY As Single, x As Single, y As Single) As Boolean
    Dim mH As Long, mW As Long, W As Long, H As Long
    Dim ELeft As Long, ETop As Long
    
    Dim EQLeft As Long
    Dim EQTop As Long
    Dim MainWidth As Long
    Dim MainHeight As Long
    
    With frmMedia
        If mini = False Then
            H = CLng(ReadINI("Main", "MaxHeight", tCurrentSkin.strConfig))
            W = CLng(ReadINI("Main", "MaxWidth", tCurrentSkin.strConfig))
            mH = CLng(ReadINI("Main", "MinHeight", tCurrentSkin.strConfig))
            mW = CLng(ReadINI("Main", "MinWidth", tCurrentSkin.strConfig))
            ELeft = CLng(ReadINI("Main", "EquaLeft", tCurrentSkin.strConfig))
            ETop = CLng(ReadINI("Main", "EquaTop", tCurrentSkin.strConfig))
            If mW <> W Then
                If (.picEqua.Left = ELeft) And (x > oldX) Then ManualShowEQ = True: Exit Function
                .picEqua.Left = .picEqua.Left + (x - oldX)
                .width = .width + (x - oldX) * Screen.TwipsPerPixelX
                If .width > (mW + (W - mW) / 2) * Screen.TwipsPerPixelX Then
                    ManualShowEQ = True
                Else
                    ManualShowEQ = False
                End If
                If .picEqua.Left >= ELeft Then
                    .picEqua.Left = ELeft
                    .width = W * Screen.TwipsPerPixelX
                    ManualShowEQ = True
                    Exit Function
                End If
                If .width <= mW * Screen.TwipsPerPixelX Then
                    .width = mW * Screen.TwipsPerPixelX
                    .picEqua.Left = mW - .picEqua.width
                    ManualShowEQ = False
                    Exit Function
                End If
            End If
            If mH <> H Then
                If (.picEqua.Top = ETop) And (y > oldY) Then ManualShowEQ = True: Exit Function
                .picEqua.Top = .picEqua.Top + (y - oldY)
                .height = .height + (y - oldY) * Screen.TwipsPerPixelY
               If .height > (mH + (H - mH) / 2) * Screen.TwipsPerPixelY Then
                    ManualShowEQ = True
                Else
                    ManualShowEQ = False
                End If
                If .picEqua.Top >= ETop Then
                    .picEqua.Top = ETop
                    .height = H * Screen.TwipsPerPixelY
                    ManualShowEQ = True
                    Exit Function
                End If
                If .height < mH * Screen.TwipsPerPixelY Then
                    .height = mH * Screen.TwipsPerPixelY
                    .picEqua.Top = mH - .picEqua.height
                    ManualShowEQ = False
                    Exit Function
                End If
            End If
        Else
            H = CLng(ReadINI("MiniMain", "MaxHeight", tCurrentSkin.strConfig))
            mH = CLng(ReadINI("MiniMain", "MinHeight", tCurrentSkin.strConfig))
            ETop = CLng(ReadINI("MiniMain", "EquaTop", tCurrentSkin.strConfig))
            If (.picMiniEqua.Top = ETop) And (y > oldY) Then ManualShowEQ = True: Exit Function
            .picMiniEqua.Top = .picMiniEqua.Top + (y - oldY)
            .height = .height + (y - oldY) * Screen.TwipsPerPixelY
            If .height > (mH + (H - mH) / 2) * Screen.TwipsPerPixelY Then
                ManualShowEQ = True
            Else
                ManualShowEQ = False
            End If
            If .picMiniEqua.Top > ETop Then
                .picMiniEqua.Top = ETop
                .height = H * Screen.TwipsPerPixelY
                ManualShowEQ = True
                Exit Function
            End If
            If .height < mH * Screen.TwipsPerPixelY Then
                .height = mH * Screen.TwipsPerPixelY
                .picMiniEqua.Top = mH - .picMiniEqua.height
                ManualShowEQ = False
                Exit Function
            End If
        End If
    End With
    
    InvalidateRect 0&, 0&, False 'refresh desktop
End Function
Public Function LoadRegions(EdgeRegions() As RegionDataType) As Boolean
    On Error GoTo handle
    Dim i As Long
    Dim Filename As String
       
    Filename = tSkinOption.SkinDir & "\" & tCurrentSkin.Infor.Name & "\region.dat"
       
    If Dir(Filename) = "" Then Exit Function
    Open Filename For Binary As #1
        
        For i = 0 To 1
            Get 1, , SkinRegions(i).DataLength
            ReDim SkinRegions(i).RegionData(SkinRegions(i).DataLength + 32)
            Get 1, , SkinRegions(i).RegionData
        Next
        
    Close #1
    LoadRegions = True
    Exit Function
handle:
    Close #1
End Function

Public Sub UnZipSkin(strSkin As String)
    On Error Resume Next
    '-- Init Global Message Variables
    uZipInfo = ""
    uZipNumber = 0   ' Holds The Number Of Zip Files
    
    '-- Select UNZIP32.DLL Options - Change As Required!
    uPromptOverWrite = 1  ' 1 = Prompt To Overwrite
    uOverWriteFiles = 0   ' 1 = Always Overwrite Files
    uDisplayComment = 0   ' 1 = Display comment ONLY!!!
    
    '-- Change The Next Line To Do The Actual Unzip!
    uExtractList = 1       ' 1 = List Contents Of Zip 0 = Extract
    uHonorDirectories = 1  ' 1 = Honour Zip Directories
    
    '-- Select Filenames If Required
    '-- Or Just Select All Files
    uZipNames.uzFiles(0) = vbNullString
    uNumberFiles = 0
    
    '-- Select Filenames To Exclude From Processing
    ' Note UNIX convention!
    '   vbxnames.s(0) = "VBSYX/VBSYX.MID"
    '   vbxnames.s(1) = "VBSYX/VBSYX.SYX"
    '   numx = 2
    
    '-- Or Just Select All Files
    uExcludeNames.uzFiles(0) = vbNullString
    uNumberXFiles = 0
    
    '-- Change The Next 2 Lines As Required!
    '-- These Should Point To Your Directory
    uZipFileName = strSkin
    uExtractDir = App.path & "\Skins\"
    If uExtractDir <> "" Then uExtractList = 0 ' unzip if dir specified
    
    
    '-- Let's Go And Unzip Them!
    Call VBUnZip32
    
End Sub

Public Function MakeRecRegion(picSoucre As Object, rec As RECT, lngTransparentColor As Long) As Long
    Dim x As Long
    Dim y As Long
    Dim StartLineX As Long
    Dim LineRegion As Long
    Dim FinishRegion As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hDC As Long
    Dim lWidth As Long
    Dim lHeight As Long
    
    hDC = picSoucre.hDC
    lWidth = rec.Right - rec.Left
    lHeight = rec.Bottom - rec.Top
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    For y = rec.Top To lHeight
        For x = rec.Left To lWidth
            If GetPixel(hDC, x, y) = lngTransparentColor Or x = lWidth Then
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    If InFirstRegion Then
                        FinishRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FinishRegion, FinishRegion, LineRegion, RGN_OR
                        DeleteObject LineRegion
                    End If
                End If
            Else
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    MakeRecRegion = FinishRegion
End Function

