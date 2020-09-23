Attribute VB_Name = "modDDE"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public lpPrevWndProc As Long


Public Sub Hook(ByVal hwnd As Long)
    lpPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook(ByVal hwnd As Long)
    Dim tmp As Long
    tmp = SetWindowLong(hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
    If uMsg = WM_COPYDATA Then
        Call ReceiveData(lngParam)
    End If
    
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lngParam)
End Function

Public Sub ReceiveData(ByVal lngParam As Long)
    Dim cdCopyData As COPYDATASTRUCT
    Dim bytBuffer(1 To 255) As Byte
    Dim StrTemp As String
          
    Call CopyMemory(cdCopyData, ByVal lngParam, Len(cdCopyData))
    
    If cdCopyData.dwData = 3 Then
        Call CopyMemory(bytBuffer(1), ByVal cdCopyData.lpData, cdCopyData.cbData)
        
        StrTemp = StrConv(bytBuffer, vbUnicode)
        StrTemp = Left$(StrTemp, InStr(1, StrTemp, Chr$(0)) - 1) 'remove null chars
        
        DoCommand StrTemp
    End If
  
End Sub

Public Sub SendData(ByVal sData As String)
    Dim cdCopyData As COPYDATASTRUCT
    Dim ThWnd As Long
    Dim bytBuffer(1 To 255) As Byte
      
    ThWnd = FindWindow(vbNullString, "MP3_ProPlayer 2.1.0")
    
    CopyMemory bytBuffer(1), ByVal sData, Len(sData)
    cdCopyData.dwData = 3
    cdCopyData.cbData = Len(sData) + 1
    cdCopyData.lpData = VarPtr(bytBuffer(1))
    Call SendMessage(ThWnd, WM_COPYDATA, frmMedia.hwnd, cdCopyData)

End Sub
