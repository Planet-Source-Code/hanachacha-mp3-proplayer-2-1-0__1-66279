VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "visOscillscope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
                                        ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" _
                                        (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
                                        ByVal crColor As Long) As Long

Dim x, r, H, y, i As Long
Dim nDataSize As Integer
Dim nColor As Long
Public Sub drawVis(hdc As Long, Indata() As Integer, Height As Single, Width As Single)
    
    If frmConfig.pic.ScaleHeight <> Height Then frmConfig.pic.Height = Height
    If frmConfig.pic.ScaleWidth <> Width Then frmConfig.pic.Width = Width
    
    Select Case tOscill.DataSize
        Case 0
            nDataSize = 500
        Case 1
            nDataSize = 1000
        Case 2
            nDataSize = 2000
        Case 3
            nDataSize = 4000
    End Select
    
    x = 0
    
    For r = 0 To nDataSize / 2 Step 1
        H = ((Indata(r) + 32768) / 65535 * Height)
        x = ((Width * r * 256) / (nDataSize / 256 * 2048)) / 8
        y = ((Indata(r * 2) + 32768) / 65535 * Height + 1)
        Select Case tOscill.DrawStyle
            Case 0 'dot
                nColor = GetPixel(frmConfig.pic.hdc, 1, H)
                If tOscill.Double = False Then
                    SetPixel hdc, x, H, nColor
                Else
                    SetPixel hdc, x, (Height - H), nColor
                    SetPixel hdc, x, y, nColor
                End If
            Case 1 'Line
                If tOscill.Double = False Then
                    BitBlt hdc, x, H, 3, 3, frmConfig.pic.hdc, 0, H, vbSrcCopy
                Else
                    BitBlt hdc, x, Height - Round(H), 3, 3, frmConfig.pic.hdc, 0, Height - (Height - H), vbSrcCopy
                    BitBlt hdc, x, Round(y), 3, 3, frmConfig.pic.hdc, 0, Height - (Height - y), vbSrcCopy
                End If
            Case 2 'Solid
                If tOscill.Double = False Then
                    BitBlt hdc, x, H, 2, Height / 2 - H, frmConfig.pic.hdc, 0, H, vbSrcCopy
                Else
                    BitBlt hdc, x, Height - Round(H), 2, Height / 2 - H, frmConfig.pic.hdc, 0, Height - (Height - H), vbSrcCopy
                    BitBlt hdc, x, Round(y), 2, Height / 2 - H, frmConfig.pic.hdc, 0, Height - (Height - y), vbSrcCopy
                End If
        End Select
    Next
End Sub


Private Sub Class_Initialize()
    Dim strINI As String
    strINI = App.Path & "\M3P_vis.ini"
    tOscill.LowColor = CLng(ReadINI("visOscillscope", "Color1", strINI, RGB(0, 255, 0)))
    tOscill.MidColor = CLng(ReadINI("visOscillscope", "Color2", strINI, RGB(255, 255, 0)))
    tOscill.HightColor = CLng(ReadINI("visOscillscope", "Color3", strINI, RGB(255, 0, 0)))
    tOscill.Double = ((ReadINI("visOscillscope", "Double", strINI, False)))
    tOscill.DrawStyle = (ReadINI("visOscillscope", "Style", strINI, 3))
    tOscill.DataSize = (ReadINI("visOscillscope", "Data", strINI, 4))

    Load frmConfig
    frmConfig.Visible = False
End Sub
Public Sub doConfig()
    frmConfig.Visible = True
End Sub

Private Sub Class_Terminate()
    Unload frmConfig
End Sub

Function OTrim(num As Long)
OTrim = num
If OTrim < 0 Then OTrim = 0
End Function


