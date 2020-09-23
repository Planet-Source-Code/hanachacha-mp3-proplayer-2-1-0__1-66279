VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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

