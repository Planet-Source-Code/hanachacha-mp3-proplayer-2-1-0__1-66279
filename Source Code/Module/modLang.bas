Attribute VB_Name = "modLang"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Public LangCap(143) As String
Public CurrentLang As String

Public Sub LoadLang(strLang As String)
    'On Error Resume Next
    Dim str As String
    Dim StrTemp As String
    Dim strFont As String
    Dim i As Long
    
    Dim fn As Long
    
    str = App.path & "\Lang\" & strLang & ".lng"
    
    If FileExists(str) = False Then Exit Sub
    
    fn = FreeFile
    Open str For Input As #fn
        'Check format
        Line Input #fn, StrTemp
        If StrTemp <> "[M3p-Lang]" Then
            MsgBox "Uncorrect language file", vbOKOnly, "M3P_Lang Error"
            Exit Sub
        End If

        'Check font for menu
        Line Input #fn, StrTemp
        strFont = Mid(StrTemp, 6)
        
        With CustomMenu
            .UseCustomFonts = True
            .CustomFonts.Add strFont
            .FontName = strFont
        End With
        
        'Check Lang Name
        Line Input #fn, StrTemp
        CurrentLang = Mid(StrTemp, 6)
        Debug.Print CurrentLang
               
        'load all string for menu
        For i = 0 To 143
            Line Input #fn, StrTemp
            LangCap(i) = Mid(StrTemp, 5)
        Next i
    Close #fn
    
    'Apply new language
    'Please don't change anything
    With frmMenu
        .mnuAbout.Caption = LangCap(0)
        .mnuOnTop.Caption = LangCap(1)
        .mnuPlay.Caption = LangCap(2)
        .mnuCDDrive.Caption = LangCap(3)
        .mnuShowPL.Caption = LangCap(4)
        .mnuShowEQ.Caption = LangCap(5)
        .mnuDSP.Caption = LangCap(6)
        .mnuLibrary.Caption = LangCap(7)
        .mnuScreen.Caption = LangCap(8)
        .mnuPb.Caption = LangCap(9)
        .mnuOption.Caption = LangCap(10)
        .mnuSkin.Caption = LangCap(11)
        .mnuVisual.Caption = LangCap(12)
        .mnuExit.Caption = LangCap(13)
        .mnuPlayC(0).Caption = LangCap(14)
        .mnuPlayC(1).Caption = LangCap(15)
        .mnuPlayC(2).Caption = LangCap(16)
        .mnuScreenC(0).Caption = LangCap(17)
        .mnuScreenC(1).Caption = LangCap(18)
        .mnuScreenC(2).Caption = LangCap(19)
        .mnuScreenC(3).Caption = LangCap(20)
        .mnuScreenC(4).Caption = LangCap(21)
        .mnuShowDesktop.Caption = LangCap(22)
        .mnuVolume.Caption = LangCap(23)
        .mnuMute.Caption = LangCap(24)
        .mnuMediaControl(0).Caption = LangCap(25)
        .mnuMediaControl(1).Caption = LangCap(26)
        .mnuMediaControl(2).Caption = LangCap(27)
        .mnuMediaControl(3).Caption = LangCap(28)
        .mnuMediaControl(4).Caption = LangCap(29)
        .mnuMediaControl(5).Caption = LangCap(30)
        .mnuMediaControl(6).Caption = LangCap(31)
        .mnuMediaControl(7).Caption = LangCap(32)
        .mnuMediaControl(8).Caption = LangCap(33)
        .mnuOptionC(0).Caption = LangCap(34)
        .mnuOptionC(2).Caption = LangCap(35)
        .mnuOptionC(3).Caption = LangCap(36)
        .mnuOptionC(5).Caption = LangCap(37)
        .mnuOptionC(6).Caption = LangCap(38)
        .mnuOptionC(7).Caption = LangCap(39)
        .mnuOptionC(9).Caption = LangCap(40)
        .mnuSkinC(0).Caption = LangCap(41)
        .mnuSkinC(1).Caption = LangCap(42)
        .mnuVisualC(0).Caption = LangCap(43)
        .mnuVisualC(1).Caption = LangCap(44)
        .mnuPl(0).Caption = LangCap(45)
        .mnuPl(1).Caption = LangCap(46)
        .mnuPl(2).Caption = LangCap(47)
        .mnuAddC(0).Caption = LangCap(48)
        .mnuAddC(1).Caption = LangCap(49)
        .mnuSubC(0).Caption = LangCap(50)
        .mnuSubC(1).Caption = LangCap(51)
        .mnuSubC(2).Caption = LangCap(52)
        .mnuSubC(3).Caption = LangCap(53)
        .mnuSubC(4).Caption = LangCap(54)
        .mnuSel(0).Caption = LangCap(55)
        .mnuSel(1).Caption = LangCap(56)
        .mnuSel(2).Caption = LangCap(57)
        .mnuSortC(0).Caption = LangCap(58)
        .mnuSortC(1).Caption = LangCap(59)
        .mnuSortC(2).Caption = LangCap(60)
        .mnuSortC(3).Caption = LangCap(61)
        .mnuSortC(4).Caption = LangCap(62)
        .mnuSortC(5).Caption = LangCap(63)
        .mnuSortStyle.Caption = LangCap(64)
        .mnuSortStyleC(0).Caption = LangCap(65)
        .mnuSortStyleC(1).Caption = LangCap(66)
        .mnuManList(0).Caption = LangCap(67)
        .mnuManList(1).Caption = LangCap(68)
        .mnuEquaLoad.Caption = LangCap(69)
        .mnuEquaSave.Caption = LangCap(70)
        .mnuEquaLoadC(0).Caption = LangCap(71)
        .mnuEquaLoadC(2).Caption = LangCap(72)
        .mnuEquaSaveC(0).Caption = LangCap(73)
        .mnuEquaSaveC(2).Caption = LangCap(74)
        .mnuMainSpecC(0).Caption = LangCap(75)
        .mnuMainSpecC(1).Caption = LangCap(76)
        .mnuMainSpecC(2).Caption = LangCap(77)
        .mnuDbl(0).Caption = LangCap(78)
        .mnuDbl(1).Caption = LangCap(79)
        .mnuDbl(2).Caption = LangCap(80)
        .mnuLibAddC(0).Caption = LangCap(81)
        .mnuLibAddC(1).Caption = LangCap(82)
        .mnuLibSubC(0).Caption = LangCap(83)
        .mnuLibSubC(1).Caption = LangCap(84)
        .mnuLibSubC(2).Caption = LangCap(85)
        .mnuLibSubC(3).Caption = LangCap(86)
        .mnuViewC(0).Caption = LangCap(87)
        .mnuViewC(1).Caption = LangCap(88)
        .mnuViewC(2).Caption = LangCap(89)
        .mnuViewC(3).Caption = LangCap(90)
        .mnuViewC(4).Caption = LangCap(91)
        .mnuViewC(5).Caption = LangCap(92)
        .mnuViewC(6).Caption = LangCap(93)
        .mnuViewC(7).Caption = LangCap(94)
        .mnuViewC(8).Caption = LangCap(95)
        .mnuViewC(9).Caption = LangCap(96)
        .mnuViewC(10).Caption = LangCap(97)
        .mnuViewC(11).Caption = LangCap(98)
        .mnuViewC(12).Caption = LangCap(99)
        .mnuViewC(13).Caption = LangCap(100)
        .mnuViewC(14).Caption = LangCap(101)
        .mnuEditLibC(0).Caption = LangCap(102)
        .mnuEditLibC(2).Caption = LangCap(103)
        .mnuEditLibC(3).Caption = LangCap(104)
        .mnuRate.Caption = LangCap(105)
        .mnuSend.Caption = LangCap(106)
        .mnuTreeViewC(0).Caption = LangCap(107)
        .mnuTreeViewC(1).Caption = LangCap(108)
        .mnuTreeViewC(3).Caption = LangCap(109)
        .mnuTreeViewC(4).Caption = LangCap(110)
        .mnuSendC(0).Caption = LangCap(111)
        .mnuSendC(2).Caption = LangCap(112)
        .mnuSendC(3).Caption = LangCap(113)
        .mnuVideoModeC(0).Caption = LangCap(114)
        .mnuVideoModeC(1).Caption = LangCap(115)
        .mnuVol(0).Caption = LangCap(116)
        .mnuVol(1).Caption = LangCap(117)
    End With
    
    Dim s As Integer
    With frmLibrary
        .cmdColumnHeader(0).Caption = LangCap(86)
        .cmdColumnHeader(1).Caption = LangCap(87)
        .cmdColumnHeader(2).Caption = LangCap(88)
        .cmdColumnHeader(3).Caption = LangCap(89)
        .cmdColumnHeader(4).Caption = LangCap(90)
        .cmdColumnHeader(5).Caption = LangCap(91)
        .cmdColumnHeader(6).Caption = LangCap(92)
        .cmdColumnHeader(7).Caption = LangCap(93)
        .cmdColumnHeader(8).Caption = LangCap(94)
        .cmdColumnHeader(9).Caption = LangCap(95)
        .cmdColumnHeader(10).Caption = LangCap(96)
        .cmdColumnHeader(11).Caption = LangCap(97)
        .cmdColumnHeader(12).Caption = LangCap(98)
        .cmdColumnHeader(13).Caption = LangCap(99)
        .cmdColumnHeader(14).Caption = LangCap(100)
        For s = 0 To 14
            .cmdColumnHeader(s).Font.Name = strFont
        Next s
    End With
    
    
    
    With CustomMenu
        i = .Icon.Count
        Do While i > 0
            .Icon.Remove i
            i = i - 1
        Loop
        .Icon.Add Array(102, 102), LangCap(0)
        .Icon.Add Array(103, 103), LangCap(1)
        .Icon.Add Array(104, 104), LangCap(2)
        .Icon.Add Array(105, 105), LangCap(3)
        .Icon.Add Array(106, 106), LangCap(4)
        .Icon.Add Array(107, 107), LangCap(5)
        .Icon.Add Array(108, 108), LangCap(6)
        .Icon.Add Array(109, 109), LangCap(7)
        .Icon.Add Array(110, 110), LangCap(8)
        .Icon.Add Array(111, 111), LangCap(9)
        .Icon.Add Array(112, 112), LangCap(10)
        .Icon.Add Array(113, 113), LangCap(11)
        .Icon.Add Array(114, 114), LangCap(12)
        .Icon.Add Array(115, 115), LangCap(13)
        .Icon.Add Array(116, 116), LangCap(14)
        .Icon.Add Array(117, 117), LangCap(15)
        .Icon.Add Array(118, 118), LangCap(16)
        .Icon.Add Array(119, 119), LangCap(23)
        .Icon.Add Array(120, 120), LangCap(24)
        .Icon.Add Array(121, 121), LangCap(25)
        .Icon.Add Array(122, 122), LangCap(26)
        .Icon.Add Array(123, 123), LangCap(27)
        .Icon.Add Array(124, 124), LangCap(28)
        .Icon.Add Array(125, 125), LangCap(29)
        .Icon.Add Array(126, 126), LangCap(30)
        .Icon.Add Array(127, 127), LangCap(31)
        .Icon.Add Array(128, 128), LangCap(32)
        .Icon.Add Array(129, 129), LangCap(33)
        .Icon.Add Array(130, 130), LangCap(34)
        .Icon.Add Array(131, 131), LangCap(35)
        .Icon.Add Array(132, 132), LangCap(36)
        .Icon.Add Array(133, 133), LangCap(37)
        .Icon.Add Array(134, 134), LangCap(38)
        .Icon.Add Array(135, 135), LangCap(39)
        .Icon.Add Array(136, 136), LangCap(40)
    End With
End Sub









