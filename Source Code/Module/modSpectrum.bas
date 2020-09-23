Attribute VB_Name = "modSpectrum"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim module_index As Long
Dim WinampInfo As String
Dim lptstrString As Long
Dim ModuleInfo As String
Dim lpModuleInfo As Long
Dim cntModule As Long
Dim NumOfModules As Long
Dim lpStr As String
Dim i As Integer

Public Sub Start_VisPlg()
    On Error GoTo beep
    Dim Index As Long
    If tWinamp.bolEnabled Then
        If (tWinamp.bolInit = False) Then
            Index = tWinamp.intCurrentPlugin
            module_index = tWinamp.intSubPlugin
            If Index = -1 Then Exit Sub
            If module_index = -1 Then Exit Sub
            Call BASS_WA_SetModule(module_index)
            Call BASS_WA_Start_Vis(Index, frmMedia.Player.handle)
            Call BASS_WA_IsPlaying(1)
            tWinamp.bolInit = True
        Else
            Call BASS_WA_IsPlaying(0)
            Index = tWinamp.intCurrentPlugin
            Call BASS_WA_Stop_Vis(Index)
            tWinamp.bolInit = False
        End If
    End If
beep:
    If Err.Number <> 0 Then
        tWinamp.bolEnabled = False
        tWinamp.intCurrentPlugin = -1
        tWinamp.intSubPlugin = -1
        Exit Sub
    End If
End Sub

Public Sub Stop_VisPlg()
    On Error GoTo beep
    Dim hindex As Long
    If tWinamp.bolEnabled Then
        If (tWinamp.bolInit = True) Then
            Call BASS_WA_IsPlaying(0)
            hindex = tWinamp.intCurrentPlugin
            Call BASS_WA_Stop_Vis(hindex)
            tWinamp.bolInit = False
        End If
    End If
beep:
    If Err.Number <> 0 Then
        tWinamp.bolEnabled = False
        tWinamp.intCurrentPlugin = -1
        tWinamp.intSubPlugin = -1
        Exit Sub
    End If
End Sub
Public Sub InitWinampVis()
    On Error GoTo beep
        Call BASS_WA_LoadVisPlugin(App.path & "\Plugins\")
        For i = 0 To BASS_WA_GetWinampPluginCount - 1
            lptstrString = BASS_WA_GetWinampPluginInfo(i)
            WinampInfo = GetStringFromPointer(lptstrString)
            frmOption.lstPlugins.AddItem WinampInfo, i
        Next i
        If BASS_WA_GetWinampPluginCount > 0 Then
            frmOption.lstPlugins.ListIndex = tWinamp.intCurrentPlugin
        End If
beep:
    If Err.Number <> 0 Then
        tWinamp.bolEnabled = False
        tWinamp.intCurrentPlugin = -1
        tWinamp.intSubPlugin = -1
        Exit Sub
    End If
End Sub
Public Sub InitVis()
    On Error Resume Next
    Dim str As String
    
    str = App.path & "\Plugins"
    With frmOption
        .File1.Pattern = "*.dll"
        .File1.path = str
        For i = 0 To frmOption.File1.ListCount - 1
            .File1.ListIndex = i
            If Mid(.File1.List(i), 1, 7) = "M3P_vis" Then
                .lstAVS.AddItem .File1.Filename
            End If
        Next i
    End With
End Sub
