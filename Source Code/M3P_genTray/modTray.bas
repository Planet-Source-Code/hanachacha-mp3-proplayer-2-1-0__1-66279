Attribute VB_Name = "modTray"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                        (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                                        ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                        ByVal lpString As Any, ByVal lplFileName As String) As Long


Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
                                        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Type NOTIFYICONDATA
        cbSize              As Long
        HWND                As Long
        uID                 As Long
        uFlags              As Long
        uCallbackMessage    As Long
        hIcon               As Long
        szTip               As String * 64
End Type

Public Type TrayControl
    bolPlay As Boolean
    bolPrevious As Boolean
    bolPause As Boolean
    bolNext As Boolean
    bolStop As Boolean
End Type

Dim tData As NOTIFYICONDATA
Public tTray As TrayControl
Public strFileconfig As String
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
Public Function ReadINI(strSection As String, strKey As String, strFileINI As String, Optional strDefault As Variant) As String
    On Error GoTo beep
    Dim StrTemp As String * 255
    GetPrivateProfileString strSection, strKey, vbNull, StrTemp, Len(StrTemp), strFileINI
    ReadINI = Mid(StrTemp, 1, InStr(1, StrTemp, vbNullChar) - 1)
    Exit Function
beep:
    ReadINI = strDefault
End Function
Public Function WriteINI(strSection As String, strKey As String, KeyValue As Variant, strFileINI As String) As Boolean
    Dim ret As Long
    ret = WritePrivateProfileString(strSection, strKey, CStr(KeyValue), strFileINI)
    If ret = 0 Then
        WriteINI = True
    Else
        WriteINI = False
    End If
End Function

