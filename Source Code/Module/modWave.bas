Attribute VB_Name = "modWave"
Option Explicit

Public Type WAVE
    Format As Integer
    Mode As String
    Freq As Long
    Bit As Byte
    Lenght As Long
End Type
Dim HOLDER$

Dim fn As Long



Public Sub ReadWave(strName As String, wavfile As WAVE)
    On Error Resume Next
    Dim yLec As Long, ydate As Date, ysg As Single
    Dim yint As Integer, YBT As Byte
    Dim LenData As Long
    Dim yDiv As Long
    Dim Extemp As Double
    Dim LenTemp As Long
    Dim n As Long
    Dim x As String, y As String, z As String
    Dim temp As Variant
    Dim fn As Long
    fn = FreeFile
    Open strName For Binary Access Read As #fn
    
            For n = 1 To 100
                x$ = Input(4, #fn)
                If n = 2 Then HOLDER$ = x$
                If x$ = "fmt " Then Exit For
            Next n
            'Get the Wave File Header Info
            Get #fn, , yLec ' 16
            Get #fn, , yint 'Compression Type (1=PCM)
            wavfile.Format = yint
            
            Get #fn, , yint 'is Channels, 1 if mono and 2 if stereo
        
            If yint = 2 Then
                wavfile.Mode = "Stereo"
              ElseIf yint = 1 Then
                wavfile.Mode = "Mono"
              Else
                wavfile.Mode = "Error!"
                GoTo beep
            End If
            Get #fn, , yLec
        
            wavfile.Freq = yLec
            Get #fn, , yLec
        
            Get #fn, , yint
            yDiv = yint
            Get #fn, , yint
        
            If yint = 8 Or yint = 16 Then
                wavfile.Bit = yint
              Else
                wavfile.Bit = 0
                GoTo beep
            End If
GotTheData:
            For n = 1 To 100
                y$ = Input(1, #fn)
                If y$ = "d" Then Exit For
            Next n
    
            z$ = Input(3, #fn)
            If z$ <> "ata" Then
                  If n > 90 Then GoTo beep
                  temp = Seek(fn)
                  Seek #fn, temp - 3
                  GoTo GotTheData
            End If
            
            Get #fn, , yLec
            LenData = yLec / yDiv
            LenTemp = LenData / wavfile.Freq
            Extemp = Int((LenTemp * 1000)) / 1000
            If LenTemp - Extemp >= 0.0005 Then
                Extemp = Extemp + 0.001
            End If
            wavfile.Lenght = Extemp
    Close #fn
    Exit Sub
beep:
    MsgBox "Error!!", vbOKOnly
    Close #fn
    Exit Sub
End Sub

