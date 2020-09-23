Attribute VB_Name = "modDeclarations"
Option Explicit

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function WindowFromPoint Lib "user32" _
                                        (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Graphic

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
                                        ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal dwRop As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
                                        ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, _
                                        ByVal nHeight As Long, ByVal hSrcDC As Long, _
                                        ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                        ByVal dwRop As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, _
                                        ByVal x As Long, ByVal Y As Long, ByVal dx As Long, _
                                        ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, _
                                        ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
                                        lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, _
                                        ByVal dwRop As Long) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
                                        (ByVal hDC As Long, ByVal lpStr As String, _
                                        ByVal nCount As Long, lpRect As RECT, _
                                        ByVal wFormat As Long) As Long
                                        
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function CreateRectRgn Lib "gdi32" _
                                        (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
                                        ByVal Y2 As Long) As Long


Public Declare Function CombineRgn Lib "gdi32" _
                                        (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
                                        ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Declare Function GetPixel Lib "gdi32" _
                                        (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long

Public Declare Function SetPixel Lib "gdi32" _
                                        (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, _
                                        ByVal crColor As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
                                        ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" _
                                        (ByVal hwnd As Long, ByVal hRgn As Long, _
                                        ByVal bRedraw As Boolean) As Long


Public Type POINTAPI
    x As Long
    Y As Long
End Type
    
Public Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Public Const RGN_COPY = 5
Public Const RGN_AND = 1
Public Const RGN_DIFF = 4
Public Const RGN_MAX = RGN_COPY
Public Const RGN_MIN = RGN_AND
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Function GetProperPath(ByVal fullPath As String) As String
    GetProperPath = IIf(Mid(fullPath, Len(fullPath), 1) = "\", fullPath, fullPath & "\")
End Function
Public Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Dir(fullPath) <> "")
End Function
