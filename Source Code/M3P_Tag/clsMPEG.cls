VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Function ReadMPEGHeader is a part of Magic MP3 Tagger. Visit www.magic-tagger.com if you're interested in
' this top MP3 tagger!
'
' Copyright by Mathias Kunter.
' You're free to use this module within your programs. These programs may also be commercial.
' The only condition for usage of this module is that you show the following line within the
' credits of your program:
'
' ID3 tagging module by Mathias Kunter (www.magic-tagger.com)

Private Type MPEG
    Version As String
    Layer As String
    Bitrate As Long
    Frequency As Long
    HasCRC As Boolean
    HasVBR As Boolean
    ChannelMode As String
    Copyrighted As Boolean
    Original As Boolean
    HasEmphasis As Boolean
    Length As Long
    FileSize As Long
End Type

Private Type v2TagHeader
    Identifier(2) As Byte
    Version(1) As Byte
    Flags As Byte
    Size(3) As Byte
End Type

Private Enum v2_StrEncoding
    ENC_ISO = 0
    ENC_UNICODE_UTF16_BOM = 1
    ENC_UNICODE_UTF16 = 2
    ENC_UNICODE_UTF8 = 3
End Enum

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                        (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = &H3

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const INVALID_HANDLE_VALUE = -1
Private MPG As MPEG

Public Function ReadMPEGHeader(strFileName As String) As Boolean
    Dim fh As Long, fp As Long, fsize As Long
    Dim fraHeader(3) As Byte, lngVers As Long, lngLayer As Long
    Dim opByte As Byte, opLong As Long, VBRHeaderPos(1) As Long
    Dim VBRHeader(17) As Byte, VBRIdent As String
    Dim SamplesPerFrame As Long, VBRNrFrames As Long, VBRBytes As Long

    fh = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function
    
    fp = GetTagSize(fh, 0)
    fsize = GetFileSize(fh, 0)
    If fsize < 4 Then
        CloseHandle fh
        Exit Function
    End If
    
    'Search for the MP3 frame header.
    Do
        SetFilePointer fh, fp, 0, FILE_BEGIN
        ReadFile fh, fraHeader(0), 4, 0, ByVal 0
        If fraHeader(0) = &HFF& And (fraHeader(1) And &HE0&) = &HE0& Then
            'Got the frame synchronisation (11 bits set).
            'Read the information from the set bits.
            
            'MPEG version
            opByte = (fraHeader(1) And &H18&) \ &H8&
            If opByte = 0 Then
                MPG.Version = "MPEG 2.5"
                lngVers = 3
            ElseIf opByte = 2 Then
                MPG.Version = "MPEG 2"
                lngVers = 2
            ElseIf opByte = 3 Then
                MPG.Version = "MPEG 1"
                lngVers = 1
            Else
                Exit Do
            End If
            
            'Layer version
            opByte = (fraHeader(1) And &H6&) \ &H2&
            If opByte = 1 Then
                lngLayer = 3
                SamplesPerFrame = IIf(lngVers = 1, 1152, 576)
            ElseIf opByte = 2 Then
                lngLayer = 2
                SamplesPerFrame = 1152
            ElseIf opByte = 3 Then
                lngLayer = 1
                SamplesPerFrame = 384
            Else
                Exit Do
            End If
            MPG.Layer = CStr(lngLayer)
            
            'CRC
            MPG.HasCRC = IIf(fraHeader(1) And &H1&, False, True)
            
            'Channel mode.
            opByte = (fraHeader(3) And &HC0&) \ &H40&
            
            'The position of an possibly existing VBR header depends on the channel mode.
            If Not opByte = 3 Then
                VBRHeaderPos(0) = fp + 4 + IIf(lngVers = 1, 32, 17)
            Else
                VBRHeaderPos(0) = fp + 4 + IIf(lngVers = 1, 17, 9)
            End If
            VBRHeaderPos(1) = fp + 4 + 32
            
            'Channel mode.
            If opByte = 0 Then
                MPG.ChannelMode = "Stereo"
            ElseIf opByte = 1 Then
                MPG.ChannelMode = "Joint stereo"
            ElseIf opByte = 2 Then
                MPG.ChannelMode = "Dual channel"
            ElseIf opByte = 3 Then
                MPG.ChannelMode = "Mono"
            End If
            
            'Frequency
            opByte = (fraHeader(2) And &HC&) \ &H4&
            If opByte = 0 Then
                opLong = 44100
            ElseIf opByte = 1 Then
                opLong = 48000
            ElseIf opByte = 2 Then
                opLong = 32000
            End If
            MPG.Frequency = opLong / IIf(lngVers = 3, 4, lngVers)
            
            'Check if there is a VBR header. If present, use it to read out the number of frames.
            SetFilePointer fh, VBRHeaderPos(0), 0, FILE_BEGIN
            ReadFile fh, VBRHeader(0), 16, 0, ByVal 0
            VBRIdent = Data2String(VarPtr(VBRHeader(0)), 4, ENC_ISO)
            If VBRIdent = "Xing" Or VBRIdent = "Info" Then
                'A VBR header is present.
                If VBRHeader(7) And &H1& Then
                    'The number of frames is stored.
                    VBRNrFrames = Data2Long(VarPtr(VBRHeader(8)), False)
                    If VBRHeader(7) And &H2& Then
                        VBRBytes = Data2Long(VarPtr(VBRHeader(12)), False)
                    End If
                End If
            End If
            'Check if there is a VBRI header.
            SetFilePointer fh, VBRHeaderPos(1), 0, FILE_BEGIN
            ReadFile fh, VBRHeader(0), 18, 0, ByVal 0
            If Data2String(VarPtr(VBRHeader(0)), 4, ENC_ISO) = "VBRI" Then
                VBRBytes = Data2Long(VarPtr(VBRHeader(10)), False)
                VBRNrFrames = Data2Long(VarPtr(VBRHeader(14)), False)
            End If
            
            If Not VBRBytes = 0 And Not VBRNrFrames = 0 Then
                'VBR bitrate.
                MPG.HasVBR = True
                MPG.Bitrate = VBRBytes / VBRNrFrames / SamplesPerFrame / 125 * MPG.Frequency
                MPG.Length = VBRNrFrames / MPG.Frequency * SamplesPerFrame
            Else
                'CBR bitrate
                MPG.HasVBR = False
                opByte = (fraHeader(2) And &HF0&) \ &H10&
                If Not opByte = 0 And Not opByte = 15 Then
                    If lngVers = 1 And lngLayer = 1 Then
                        MPG.Bitrate = opByte * 32
                    ElseIf lngVers = 1 And lngLayer = 2 Then
                        If opByte = 1 Then
                            MPG.Bitrate = 32
                        ElseIf opByte = 2 Then
                            MPG.Bitrate = 48
                        ElseIf opByte <= 4 Then
                            MPG.Bitrate = 48 + (opByte - 2) * 8
                        ElseIf opByte <= 8 Then
                            MPG.Bitrate = 64 + (opByte - 4) * 16
                        ElseIf opByte <= 12 Then
                            MPG.Bitrate = 128 + (opByte - 8) * 32
                        Else
                            MPG.Bitrate = 256 + (opByte - 12) * 64
                        End If
                    ElseIf lngVers = 1 And lngLayer = 3 Then
                        If opByte = 1 Then
                            MPG.Bitrate = 32
                        ElseIf opByte <= 5 Then
                            MPG.Bitrate = 32 + (opByte - 1) * 8
                        ElseIf opByte <= 9 Then
                            MPG.Bitrate = 64 + (opByte - 5) * 16
                        ElseIf opByte <= 13 Then
                            MPG.Bitrate = 128 + (opByte - 9) * 32
                        Else
                            MPG.Bitrate = 320
                        End If
                    ElseIf lngVers >= 2 And lngLayer = 1 Then
                        If opByte = 1 Then
                            MPG.Bitrate = 32
                        ElseIf opByte = 2 Then
                            MPG.Bitrate = 48
                        ElseIf opByte <= 4 Then
                            MPG.Bitrate = 48 + (opByte - 2) * 8
                        ElseIf opByte <= 12 Then
                            MPG.Bitrate = 64 + (opByte - 4) * 16
                        Else
                            MPG.Bitrate = 192 + (opByte - 12) * 32
                        End If
                    Else
                        'mVers >= 2, lVers >= 2
                        If opByte <= 8 Then
                            MPG.Bitrate = opByte * 8
                        Else
                            MPG.Bitrate = 64 + (opByte - 8) * 16
                        End If
                    End If
                End If
                MPG.Length = (fsize - fp) / (MPG.Bitrate * 125)
            End If
                        
            'Copyright
            MPG.Copyrighted = IIf(fraHeader(3) And &H8&, True, False)
            
            'Original
            MPG.Original = IIf(fraHeader(3) And &H4&, True, False)
            
            'Emphasis
            MPG.HasEmphasis = IIf(fraHeader(3) And &H3&, True, False)
            
            ReadMPEGHeader = True
            Exit Do
        Else
            fp = fp + 1
            If fp > fsize - 4 Then Exit Do
        End If
    Loop
    
    MPG.FileSize = fsize
    
    CloseHandle fh
End Function
Private Function Data2Long(ByVal pData As Long, ByVal bSynchSafe As Boolean) As Long
    Dim i As Integer, Data(3) As Byte

    CopyMemory Data(0), ByVal pData, 4

    'Avoid converting wrong synchsafe integers. If bit 7 of any byte is set, it is not synchsafe.
    'However, we can't detect wrong coded values which have bit 7 zeroed.
    For i = 0 To 3
        If Data(i) And &H80& Then bSynchSafe = False
    Next i

    'Perform left-shifts, done by multiplication with the hex values of 2^n. Finally, bit-or the values.
    If bSynchSafe Then
        Data2Long = (Data(0) * &H200000) Or (Data(1) * &H4000&) Or (Data(2) * &H80&) Or Data(3)
    Else
        Data2Long = (Data(0) * &H1000000) Or (Data(1) * &H10000) Or (Data(2) * &H100&) Or Data(3)
    End If
End Function
Private Function Data2String(ByVal pData As Long, ByVal Length As Long, ByVal EncFormat As v2_StrEncoding, Optional ByVal BreakOnNull As Boolean = True) As String
    Dim i As Long, curData As Byte, curSign As String

    For i = 0 To Length - 1
        CopyMemory curData, ByVal pData + i, 1
        'New lines are represented by &0A& (which is chr(10)) only in ID3 v2 tags.
        'However, many programs seem to code the newline with chr(13) & chr(10),
        'which is the windows default. Therefore, skip chr(13) and change chr(10) to vbNewLines.
        If curData = 13 Then
            curSign = ""
        ElseIf curData = 10 Then
            curSign = vbNewLine
        Else
            curSign = Chr(curData)
        End If
        If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
            'Clear text, null terminated.
            If curData = 0 And BreakOnNull Then
                Exit Function
            Else
                Data2String = Data2String & curSign
            End If
        ElseIf EncFormat = ENC_UNICODE_UTF16_BOM Then
            'UNICODE text with BOM, double-null terminated.
            If i >= 2 And i Mod 2 = 0 Then
                If curData = 0 And BreakOnNull Then
                    Exit Function
                Else
                    Data2String = Data2String & curSign
                End If
            End If
        ElseIf EncFormat = ENC_UNICODE_UTF16 Then
            'UNICODE text without BOM, double-null terminated.
            If i Mod 2 = 0 Then
                If curData = 0 And BreakOnNull Then
                    Exit Function
                Else
                    Data2String = Data2String & curSign
                End If
            End If
        End If
    Next i
End Function
Private Function GetTagSize(ByVal fh As Long, ByVal fp As Long) As Long
    Dim TagHeader As v2TagHeader
    
    'Search for an ID3v2 tag.
    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
    If Data2String(VarPtr(TagHeader.Identifier(0)), 3, ENC_ISO) = "ID3" Then
        'The size stored in the header excludes itself, and excludes the footer (if present).
        GetTagSize = Data2Long(VarPtr(TagHeader.Size(0)), True) + Len(TagHeader)

        'v 2.4 (or later?) flags: %abcd0000 abc = ignored, d = footer present
        If TagHeader.Version(0) >= 4 Then
            If TagHeader.Flags And &H10& Then
                'Add the size of the footer (which is the same size than the header) to the existing size.
                GetTagSize = GetTagSize + Len(TagHeader)
            End If
        End If
    End If
End Function

Public Property Get Version() As String
    Version = MPG.Version
End Property
Public Property Get Layer() As String
    Layer = MPG.Layer
End Property

Public Property Get Bitrate() As Long
    Bitrate = MPG.Bitrate
End Property
Public Property Get Frequency() As Long
    Frequency = MPG.Frequency
End Property

Public Property Get Length() As Long
    Length = MPG.Length
End Property
Public Property Get FileSize() As Long
    FileSize = MPG.FileSize
End Property
Public Property Get HasCRC() As Boolean
    HasCRC = MPG.HasCRC
End Property
Public Property Get HasVBR() As Boolean
    HasVBR = MPG.HasVBR
End Property
Public Property Get Copyrighted() As Boolean
    Copyrighted = MPG.Copyrighted
End Property
Public Property Get Original() As Boolean
    Original = MPG.Original
End Property
Public Property Get HasEmphasis() As Boolean
    HasEmphasis = MPG.HasEmphasis
End Property
Public Property Get ChannelMode() As String
    ChannelMode = MPG.ChannelMode
End Property













