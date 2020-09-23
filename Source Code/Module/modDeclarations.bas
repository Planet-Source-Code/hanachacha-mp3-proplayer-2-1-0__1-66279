Attribute VB_Name = "modDeclarations"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long

Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
                                        (ByVal hDC As Long, ByVal lpStr As String, _
                                        ByVal nCount As Long, lpRect As RECT, _
                                        ByVal wFormat As Long) As Long
Public Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" _
                                        (ByVal lpsz As String, ByVal pt As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function WindowFromPoint Lib "user32" _
                                        (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, ByVal wMsg As Long, _
                                        ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                        (ByVal hwnd As Long, ByVal lpOperation As String, _
                                        ByVal lpFile As String, ByVal lpParameters As String, _
                                        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
                                        (ByVal lpFileName As String, ByVal nBufferLength As Long, _
                                        ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                        (Destination As Any, Source As Any, ByVal length As Long)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Regoin
Public Declare Function SetWindowRgn Lib "user32" _
                                        (ByVal hwnd As Long, ByVal hRgn As Long, _
                                        ByVal bRedraw As Boolean) As Long

Public Declare Function CreateRectRgn Lib "gdi32" _
                                        (ByVal X1 As Long, ByVal Y1 As Long, _
                                        ByVal X2 As Long, _
                                        ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, _
                                        ByVal Y1 As Long, ByVal X2 As Long, _
                                        ByVal Y2 As Long, _
                                        ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function CombineRgn Lib "gdi32" _
                                        (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
                                        ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Declare Function GetPixel Lib "gdi32" _
                                        (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function SetPixel Lib "gdi32" _
                                        (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
                                        ByVal crColor As Long) As Long
                                        
Public Declare Function SetWindowPos Lib "user32" _
                                        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                        ByVal x As Long, ByVal y As Long, ByVal cX As Long, _
                                        ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" _
                                        (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, _
                                        ByVal dwFlags As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                                        (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'INI Edit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                        (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                                        ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                        ByVal lpString As Any, ByVal lplFileName As String) As Long
'XPTheme
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'Graphic

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
                                        ByVal hObject As Long) As Long

Public Declare Function TransparentBlt Lib "msimg32" _
                                        (ByVal hDCDst As Long, ByVal nXOriginDst As Long, _
                                        ByVal nYOriginDst As Long, ByVal nWidthDst As Long, _
                                        ByVal nHeightDst As Long, ByVal hDCSrc As Long, _
                                        ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                                        ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                                        ByVal crTransparent As Long) As Long


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
                                        ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal dwRop As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
                                        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
                                        ByVal nHeight As Long, ByVal hSrcDC As Long, _
                                        ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                        ByVal dwRop As Long) As Long

    
    Public Const RGN_COPY = 5
    Public Const RGN_AND = 1
    Public Const RGN_DIFF = 4
    Public Const RGN_MAX = RGN_COPY
    Public Const RGN_MIN = RGN_AND
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3

    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_SHOWWINDOW = &H40
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const WM_SETTEXT = &HC
    
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDOWN         As Long = &H201
    Public Const WM_LBUTTONUP           As Long = &H202
    Public Const WM_LBUTTONDBLCLK       As Long = &H203
    Public Const WM_RBUTTONDOWN         As Long = &H204
    Public Const WM_RBUTTONUP           As Long = &H205
    Public Const WM_RBUTTONDBLCLK       As Long = &H206
    Public Const WM_MBUTTONDOWN         As Long = &H207
    Public Const WM_MBUTTONUP           As Long = &H208
    Public Const WM_MBUTTONDBLCLK       As Long = &H209
    Public Const WM_MOUSEWHEEL          As Long = &H20A

    Public Const GWL_STYLE = (-16)
    Public Const ES_NUMBER = &H2000&
    Public Const SW_SHOWNORMAL = 1
    
    Public Const DT_RIGHT = &H2
    Public Const DT_LEFT = &H0
    Public Const DT_CENTER = &H1
    
    Public Const GWL_WNDPROC = (-4)
    Public Const WM_COPYDATA = &H4A
    
    Public Type COPYDATASTRUCT
            dwData As Long
            cbData As Long
            lpData As Long
    End Type
    
    Public Type POINTAPI
            x As Long
            y As Long
    End Type
    
    Public Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
    End Type




