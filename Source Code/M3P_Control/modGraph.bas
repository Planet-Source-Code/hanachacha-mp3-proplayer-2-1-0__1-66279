Attribute VB_Name = "modGraph"
Option Explicit
Const DIB_RGB_ColS      As Long = 0

Private Type BITMAPINFOHEADER    '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Public Enum GradientDirectionEnum
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_HorizontalMiddleOut] = 2
    [Fill_Vertical] = 3
    [Fill_VerticalMiddleOut] = 4
    [Fill_DownwardDiagonal] = 5
    [Fill_UpwardDiagonal] = 6
End Enum


Private Sub DIBGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim J       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    '-- Decompose Cols
    Col1 = Col1 And &HFFFFFF
    R1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    G1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    B1 = Col1 Mod &H100&
    Col2 = Col2 And &HFFFFFF
    R2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    G2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    B2 = Col2 Mod &H100&
    
    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To Width - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For J = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next J
        
        Case [Fill_Vertical]
        
            For J = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(J)
                Next i
                iOffset = iOffset + Scan
            Next J
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For J = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = J
            Next J
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For J = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = J
            Next J
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)

End Sub
Public Sub FillGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum, _
                         Optional Right2Left As Boolean = True)
                         
    Dim tmpCol  As Long
  
    ' Exit if needed
    If GradientDirection = Fill_None Then Exit Sub
    
    ' Right-To-Left
    If Right2Left Then
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    Select Case GradientDirection
        Case Fill_HorizontalMiddleOut
            DIBGradient hDC, x, y, Width / 2, Height, Col1, Col2, Fill_Horizontal
            DIBGradient hDC, x + Width / 2 - 1, y, Width / 2, Height, Col2, Col1, Fill_Horizontal

        Case Fill_VerticalMiddleOut
            DIBGradient hDC, x, y, Width, Height / 2, Col1, Col2, Fill_Vertical
            DIBGradient hDC, x, y + Height / 2 - 1, Width, Height / 2, Col2, Col1, Fill_Vertical

        Case Else
            DIBGradient hDC, x, y, Width, Height, Col1, Col2, GradientDirection
    End Select
    
End Sub
Public Function MakeRegion(picSoucre As Object, lngTransparentColor As Long) As Long
    Dim x As Long
    Dim y As Long
    Dim StartLineX As Long
    Dim LineRegion As Long
    Dim FinishRegion As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hDC As Long
    Dim lWidth As Long
    Dim lHeight As Long
    
    hDC = picSoucre.hDC
    lWidth = picSoucre.ScaleWidth
    lHeight = picSoucre.ScaleHeight
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    For y = 0 To lHeight
        For x = 0 To lWidth
            If GetPixel(hDC, x, y) = lngTransparentColor Or x = lWidth Then
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    If InFirstRegion Then
                        FinishRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FinishRegion, FinishRegion, LineRegion, RGN_OR
                        DeleteObject LineRegion
                    End If
                End If
            Else
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    MakeRegion = FinishRegion
End Function

