Attribute VB_Name = "modSys"
Option Explicit

Dim tData As NOTIFYICONDATA
Public Sub SysTrayIcon(HWND As Long, Icon As Long)
    With tData
        .cbSize = Len(tData)
        .HWND = HWND
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Icon
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, tData
End Sub
Public Sub AddIcon(HWND As Long, Icon As Long)
    With tData
        .cbSize = Len(tData)
        .HWND = HWND
        .uCallbackMessage = WM_MOUSEMOVE
        .uID = 1&
        .hIcon = Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE
    End With
    Shell_NotifyIcon NIM_ADD, tData
End Sub
Public Sub RemoveIcon(HWND As Long)
    With tData
        .HWND = HWND
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, tData
End Sub
Public Sub SysTip(HWND As Long, ToolTip As String)
    With tData
        .HWND = HWND
        .szTip = ToolTip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, tData
End Sub
