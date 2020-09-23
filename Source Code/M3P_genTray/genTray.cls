VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "genTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub About()
    MsgBox "M3P_genTray v1.0 ,remote control for M3P", vbOKOnly, "About"
End Sub
Public Sub Config()
    frmTray.Show
    frmTray.Visible = True
End Sub
Public Sub Run()
        If tTray.bolStop Then
            Call AddIcon(frmTray.picTray(4).HWND, frmTray.imgTray.ListImages(5).ExtractIcon.Handle)
            Call SysTip(frmTray.picTray(4).HWND, "Stop")
        End If
        If tTray.bolPause Then
            Call AddIcon(frmTray.picTray(3).HWND, frmTray.imgTray.ListImages(4).ExtractIcon.Handle)
            Call SysTip(frmTray.picTray(3).HWND, "Pause")
        End If
        If tTray.bolNext Then
            Call AddIcon(frmTray.picTray(2).HWND, frmTray.imgTray.ListImages(3).ExtractIcon.Handle)
            Call SysTip(frmTray.picTray(2).HWND, "Next")
        End If
        If tTray.bolPlay Then
            Call AddIcon(frmTray.picTray(1).HWND, frmTray.imgTray.ListImages(2).ExtractIcon.Handle)
            Call SysTip(frmTray.picTray(1).HWND, "Play")
        End If
        If tTray.bolPrevious Then
            Call AddIcon(frmTray.picTray(0).HWND, frmTray.imgTray.ListImages(1).ExtractIcon.Handle)
            Call SysTip(frmTray.picTray(0).HWND, "Previous")
        End If
End Sub

Private Sub Class_Initialize()
    strFileconfig = Mid(App.Path, 1, Len(App.Path) - Len("Plugins")) & "M3P.ini"
    
    tTray.bolNext = ReadINI("M3PTray", "Next", strFileconfig, False)
    tTray.bolPlay = ReadINI("M3PTray", "Play", strFileconfig, False)
    tTray.bolPrevious = ReadINI("M3PTray", "Previous", strFileconfig, False)
    tTray.bolPause = ReadINI("M3PTray", "Pause", strFileconfig, False)
    tTray.bolStop = ReadINI("M3PTray", "Stop", strFileconfig, False)
    Load frmTray
    frmTray.Visible = False
End Sub

Private Sub Class_Terminate()
    WriteINI "M3PTray", "Play", tTray.bolPlay, strFileconfig
    WriteINI "M3PTray", "Previous", tTray.bolPrevious, strFileconfig
    WriteINI "M3PTray", "Pause", tTray.bolPause, strFileconfig
    WriteINI "M3PTray", "Next", tTray.bolNext, strFileconfig
    WriteINI "M3PTray", "Stop", tTray.bolStop, strFileconfig
    
End Sub
