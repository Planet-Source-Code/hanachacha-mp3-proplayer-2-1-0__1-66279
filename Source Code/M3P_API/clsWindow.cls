VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetSysColors Lib "user32.dll" (ByVal nChanges As Long, _
                                                        lpSysColor As Long, _
                                                        lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

'Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                ByVal hWndNewParent As Long) As Long
                                                
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByVal lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                                        (ByVal Hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long


Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SW_NORMAL = 1
Const SW_HIDE = 0


Public Enum DIR_ID
    DIR_USER = &H28
    DIR_USER_DESKTOP = &H10
    DIR_USER_MY_DOCUMENTS = &H5
    DIR_USER_START_MENU = &HB
    DIR_USER_START_MENU_PROGRAMS = &H2
    DIR_USER_START_MENU_PROGRAMS_STARTUP = &H7
    DIR_COMMON_DESKTOP = &H19
    DIR_COMMON_DOCUMENTS = &H2E
    DIR_COMMON_START_MENU = &H16
    DIR_COMMON_START_MENU_PROGRAMS = &H17
    DIR_COMMON_START_MENU_PROGRAMS_STARTUP = &H18
    DIR_WINDOWS = &H24
    DIR_SYSTEM = &H25
    DIR_FONTS = &H14
    DIR_PROGRAM_FILES = &H26
    DIR_PROGRAM_FILES_COMMON_FILES = &H2B
End Enum
Public Enum WinDir
    eDIR_USER = 1
    eDIR_USER_DESKTOP = 2
    eDIR_USER_MY_DOCUMENTS = 3
    eDIR_USER_START_MENU = 4
    eDIR_USER_START_MENU_PROGRAMS = 5
    eDIR_USER_START_MENU_PROGRAMS_STARTUP = 6
    eDIR_COMMON_DESKTOP = 7
    eDIR_COMMON_DOCUMENTS = 8
    eDIR_COMMON_START_MENU = 9
    eDIR_COMMON_START_MENU_PROGRAMS = 10
    eDIR_COMMON_START_MENU_PROGRAMS_STARTUP = 11
    eDIR_WINDOWS = 12
    eDIR_SYSTEM = 13
    eDIR_FONTS = 14
    eDIR_PROGRAM_FILES = 15
    eDIR_PROGRAM_FILES_COMMON_FILES = 16
End Enum

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Const MAX_PATH = 260
Const WS_EX_LAYERED As Long = &H80000
Const LWA_ALPHA As Long = &H2
Const GWL_EXSTYLE = (-20)
Const LWA_COLORKEY = &H1
Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
    
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
    
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuiDUDT As LUID
    Attributes As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
    
Public Enum DesktopTheme
    ClearWall
    StoreWall
    RestoreWall
End Enum

    
Private Const COLOR_BACKGROUND            As Integer = 1
Private Const ERROR_SUCCESS = 0
Private Const ERROR_NONE                  As Integer = 0
Private Const SPI_SETDESKWALLPAPER        As Integer = 20
Private Const SPIF_SENDWININICHANGE       As Long = &H2
Private Const SPIF_UPDATEINIFILE          As Long = &H1


Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
                                        (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
                                        NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
                                        ByRef PreviousState As Any, _
                                        ByRef ReturnLength As Any) As Long

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, _
                                        ByVal dwReserved As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function LookupPrivilegeValueA Lib "advapi32.dll" _
                                        (ByVal lpSystemName As String, ByVal lpName As String, _
                                        lpLuid As LUID) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                                        (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, _
                                        ByVal dwFlags As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                        (lpVersionInformation As OSVersionInfo) As Long


Private Declare Function GetComputerNameA Lib "kernel32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const SPI_SETSCREENSAVEACTIVE     As Integer = 17

Private Wallpaper As String
Private WallpaperStyle As String
Private OriginalWallpaper As String
Private OriginalBackground As Long

Private Reg As New clsRegistry


Public Function GetUserName() As String
    Dim var1 As String, ns As Long
    ns = 255
    var1 = String(ns, 0)
    GetUserNameA var1, ns
    var1 = Left$(var1, ns - 1)
    GetUserName = var1
End Function
Public Function GetComputerName() As String
    Dim var1 As String, ns As Long
    ns = 32
    var1 = String(ns, 0)
    GetComputerNameA var1, ns
    var1 = Left$(var1, ns)
    GetComputerName = var1
End Function

Private Function GetWindowsDir() As String
    Dim StrLen As Long, zPath As String
    zPath = String$(MAX_PATH, 0)
    StrLen = GetWindowsDirectory(zPath, MAX_PATH)
    GetWindowsDir = Left$(zPath, StrLen)
End Function

Private Function GetSystemDir() As String
    Dim StrLen As Long, zPath As String
    zPath = String$(MAX_PATH, 0)
    StrLen = GetSystemDirectory(zPath, MAX_PATH)
    GetSystemDir = Left$(zPath, StrLen)
End Function

Private Function GetSpecialFolder(CSIDL As DIR_ID) As String
    Dim zPath As String, r As Long, IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(ByVal 0&, CSIDL, IDL)
    If r = 0 Then
        zPath = String$(MAX_PATH, 0)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal zPath)
        GetSpecialFolder = Left$(zPath, InStr(zPath, Chr$(0)) - 1)
    Else
        GetSpecialFolder = ""
    End If
End Function

Public Sub Make_TransPercent(Hwnd As Long, Percent As Integer)
    On Error Resume Next
    Dim OSV As OSVersionInfo
    
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) <> 1 Then Exit Sub
    If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub '/* Win 98/ME
    If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
        Call SetWindowLong(Hwnd, GWL_EXSTYLE, GetWindowLong(Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(Hwnd, 0, (Percent * 255) / 100, LWA_ALPHA)
    End If
End Sub
Public Sub Make_TransColor(Hwnd As Long, Color As Long)
    On Error Resume Next
    Dim OSV As OSVersionInfo
    
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) <> 1 Then Exit Sub
    If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub '/* Win 98/ME
    If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
        Call SetWindowLong(Hwnd, GWL_EXSTYLE, GetWindowLong(Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(Hwnd, Color, 0, LWA_COLORKEY)
    End If
End Sub
Public Sub SHUTDOWN(Optional ByVal REBOOT As Boolean = False, Optional ByVal Force As Boolean = False)
    Const EWX_FORCE = 4
    Const EWX_REBOOT = 2
    Const EWX_SHUTDOWN = 1
    Const EWX_LOGOFF = 0
    Const EWX_POWEROFF = 8
    Const SE_PRIVILEGE_ENABLED = &H2
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8

    On Error Resume Next
    
    Dim flags As Long
    Dim Token As Long
    Dim TP As TOKEN_PRIVILEGES
    Dim OSV As OSVersionInfo
    
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) <> 1 Then Exit Sub
    If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
        OpenProcessToken GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, Token
        LookupPrivilegeValueA "", "SeShutdownPrivilege", TP.LuiDUDT
        TP.PrivilegeCount = 1
        TP.Attributes = SE_PRIVILEGE_ENABLED
        AdjustTokenPrivileges Token, False, TP, 0, ByVal 0&, ByVal 0&
    End If
    'Shutdown or Reboot
    flags = EWX_SHUTDOWN
    If REBOOT Then
        flags = flags Or EWX_REBOOT
    Else
        flags = flags Or EWX_POWEROFF
    End If
    If Force Then flags = flags Or EWX_FORCE
    ExitWindowsEx flags, &HFFFF
End Sub

Public Function GetWinInfor(eDir As WinDir) As String
    On Error Resume Next
    Select Case eDir
        Case 1
            GetWinInfor = GetSpecialFolder(DIR_USER)
        Case 2
            GetWinInfor = GetSpecialFolder(DIR_USER_DESKTOP)
        Case 3
            GetWinInfor = GetSpecialFolder(DIR_USER_MY_DOCUMENTS)
        Case 4
            GetWinInfor = GetSpecialFolder(DIR_USER_START_MENU)
        Case 5
            GetWinInfor = GetSpecialFolder(DIR_USER_START_MENU_PROGRAMS)
        Case 6
            GetWinInfor = GetSpecialFolder(DIR_USER_START_MENU_PROGRAMS_STARTUP)
        Case 7
            GetWinInfor = GetSpecialFolder(DIR_COMMON_DESKTOP)
        Case 8
            GetWinInfor = GetSpecialFolder(DIR_COMMON_DOCUMENTS)
        Case 9
            GetWinInfor = GetSpecialFolder(DIR_COMMON_START_MENU)
        Case 10
            GetWinInfor = GetSpecialFolder(DIR_COMMON_START_MENU_PROGRAMS)
        Case 11
            GetWinInfor = GetSpecialFolder(DIR_COMMON_START_MENU_PROGRAMS_STARTUP)
        Case 12
            GetWinInfor = GetWindowsDir
        Case 13
            GetWinInfor = GetSystemDir
        Case 14
            GetWinInfor = GetSpecialFolder(DIR_FONTS)
        Case 15
            GetWinInfor = GetSpecialFolder(DIR_PROGRAM_FILES)
        Case 16
            GetWinInfor = GetSpecialFolder(DIR_PROGRAM_FILES_COMMON_FILES)
    End Select
End Function
Private Sub ClearDesktop()
    Dim i As Long
    OriginalBackground = GetSysColor(i)
    SetSysColors 1, COLOR_BACKGROUND, &H0
    Reg.ClassKey = HKEY_CURRENT_USER
    Reg.SectionKey = "Control Panel\Desktop"
    Reg.ValueKey = "Wallpaper"
    Reg.ValueType = REG_SZ
    Reg.Value = ""
    Reg.CreateKey
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, "", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Sub
Private Sub SetWallPaper(ByVal Display As Integer, ByVal strWallpaper As String)
    On Error Resume Next
    
    Reg.ClassKey = HKEY_CURRENT_USER
    Reg.SectionKey = "Control Panel\Desktop"
    Reg.ValueType = REG_SZ
    
    ' TileWallpaper = 0 --> Stretch
    ' TileWallpaper > 0 --> Tile
    ' WallpaperStyle = 0 --> Center
    ' WallpaperStyle = 1 --> Tile
    ' WallpaperStyle = 2 --> Stretch
    Select Case Display
        Case 0
            Reg.ValueKey = "TileWallpaper"
            Reg.Value = "0"
            Reg.ValueKey = "WallpaperStyle"
            Reg.Value = "0"
            Reg.CreateKey
        Case 1
            Reg.ValueKey = "TileWallpaper"
            Reg.Value = "0"
            Reg.ValueKey = "WallpaperStyle"
            Reg.Value = "1"
            Reg.CreateKey
        Case 2
            Reg.ValueKey = "TileWallpaper"
            Reg.Value = "0"
            Reg.ValueKey = "WallpaperStyle"
            Reg.Value = "2"
            Reg.CreateKey
    
    End Select
    
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, strWallpaper, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE

End Sub
Public Sub StoreWallpaper()
    On Error Resume Next
    Reg.ClassKey = HKEY_CURRENT_USER
    Reg.SectionKey = "Control Panel\Desktop"
    
    Reg.ValueKey = "Wallpaper"
    Reg.ValueType = REG_SZ
    Wallpaper = Reg.Value
    
    Reg.ValueKey = "WallpaperStyle"
    WallpaperStyle = Reg.Value
    
    Reg.ValueKey = "OriginalWallpaper"
    OriginalWallpaper = Reg.Value
    
    If LenB(Wallpaper) = 0 Then
        Wallpaper = OriginalWallpaper
    ElseIf WallpaperStyle = Not IsNumeric(WallpaperStyle) Then
        WallpaperStyle = 2 'as default [ stretch ]
    End If
    
    SetSysColors 1, COLOR_BACKGROUND, OriginalBackground

End Sub
Public Sub SettingWallpaper(action As DesktopTheme)
    Select Case action
        Case ClearWall
            StoreWallpaper
            ClearDesktop
        Case StoreWall
            StoreWallpaper
        Case RestoreWall
            If Not IsNumeric(WallpaperStyle) Then
                WallpaperStyle = 2 'as default [ stretch ]
            End If
            SetWallPaper WallpaperStyle, Wallpaper
    End Select
End Sub

Public Sub HideDesktop(ByVal bolShow As Boolean)
    Dim Hwnd As Long
    Hwnd = GetDesktopWindow
    If bolShow Then
        ShowWindow Hwnd, SW_HIDE
    End If
    If Not bolShow Then
        ShowWindow Hwnd, SW_NORMAL
    End If
End Sub
Public Sub ScreenSaverActive(ByVal Active As Boolean)
    Dim Enabled As Long
    
    Enabled = IIf(Active, 1, 0)
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, Enabled, 0&, 0
End Sub


