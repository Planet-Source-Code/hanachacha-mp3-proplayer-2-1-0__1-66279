' BASSWMA 2.3 Visual Basic API, copyright (c) 2002-2006 Ian Luck.
' Requires BASS 2.3 - available @ www.un4seen.com

' See the BASSWMA.CHM file for more complete documentation

Attribute VB_Name = "BASSWMA"

' Additional error codes returned by BASS_ErrorGetCode
Global Const BASS_ERROR_WMA_LICENSE = 1000     ' the file is protected
Global Const BASS_ERROR_WMA = 1001             ' Windows Media (9 or above) is not installed
Global Const BASS_ERROR_WMA_WM9 = BASS_ERROR_WMA
Global Const BASS_ERROR_WMA_DENIED = 1002      ' access denied (user/pass is invalid)
Global Const BASS_ERROR_WMA_INDIVIDUAL = 1004  ' individualization is needed

' Additional config options
Global Const BASS_CONFIG_WMA_PRECHECK = &H10100
Global Const BASS_CONFIG_WMA_PREBUF = &H10101

' additional WMA sync type
Global Const BASS_SYNC_WMA_CHANGE = 1001

' additional BASS_StreamGetFilePosition WMA mode
Global Const BASS_FILEPOS_WMA_BUFFER = 1000 ' internet buffering progress (0-100%)

' Additional flags for use with BASS_WMA_EncodeOpenFile/Network
Global Const BASS_WMA_ENCODE_SCRIPT = &H20000  ' set script (mid-stream tags) in the WMA encoding

' Additional flag for use with BASS_WMA_EncodeGetRates
Global Const BASS_WMA_ENCODE_RATES_VBR = &H10000 ' get available VBR quality settings

' WMENCODEPROC "type" values
Global Const BASS_WMA_ENCODE_HEAD = 0
Global Const BASS_WMA_ENCODE_DATA = 1
Global Const BASS_WMA_ENCODE_DONE = 2

' BASS_WMA_EncodeSetTag "type" values
Global Const BASS_WMA_TAG_ANSI    = 0
Global Const BASS_WMA_TAG_UNICODE = 1
Global Const BASS_WMA_TAG_UTF8    = 2

' BASS_CHANNELINFO type
Global Const BASS_CTYPE_STREAM_WMA = &H10300
Global Const BASS_CTYPE_STREAM_WMA_MP3 = &H10301

' Additional BASS_StreamGetTags type
Global Const BASS_TAG_WMA = 8 ' WMA tags : array of null-terminated UTF-8 strings


Declare Function BASS_WMA_StreamCreateFile Lib "basswma.dll" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
Declare Function BASS_WMA_StreamCreateFileAuth Lib "basswma.dll" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long, ByVal user As String, ByVal pass As String) As Long
Declare Function BASS_WMA_StreamCreateFileUser Lib "basswma.dll" (ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long

Declare Function BASS_WMA_EncodeGetRates Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long) As Long
Declare Function BASS_WMA_EncodeOpen Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal bitrate As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_WMA_EncodeOpenFile Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal bitrate As Long, ByVal file As String) As Long
Declare Function BASS_WMA_EncodeOpenNetwork Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal bitrate As Long, ByVal port As Long, ByVal clients As Long) As Long
Declare Function BASS_WMA_EncodeOpenNetworkMulti Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByRef bitrates As Long, ByVal port As Long, ByVal clients As Long) As Long
Declare Function BASS_WMA_EncodeOpenPublish Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal bitrate As Long, ByVal url As String, ByVal user As String, ByVal pass As String) As Long
Declare Function BASS_WMA_EncodeOpenPublishMulti Lib "basswma.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByRef bitrates As Long, ByVal url As String, ByVal user As String, ByVal pass As String) As Long
Declare Function BASS_WMA_EncodeGetPort Lib "basswma.dll" (ByVal handle As Long) As Long
Declare Function BASS_WMA_EncodeSetNotify Lib "basswma.dll" (ByVal handle As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_WMA_EncodeGetClients Lib "basswma.dll" (ByVal handle As Long) As Long
Declare Function BASS_WMA_EncodeSetTag Lib "basswma.dll" (ByVal handle As Long, ByVal tag As String, ByVal text As String, ByVal ttype As Long) As Long
Declare Function BASS_WMA_EncodeWrite Lib "basswma.dll" (ByVal handle As Long, ByVal buffer As Long, ByVal length As Long) As Long
Declare Function BASS_WMA_EncodeClose Lib "basswma.dll" (ByVal handle As Long) As Long

Declare Function BASS_WMA_GetWMObject Lib "basswma.dll" (ByVal handle As Long) As Long


Sub CLIENTCONNECTPROC(ByVal handle As Long, ByVal connect As Long, ByVal ip As Long, ByVal user As Long)

    'CALLBACK FUNCTION !!!

    ' Client connection notification callback function.
    ' handle : The encoder
    ' connect: TRUE=client is connecting, FALSE=disconnecting
    ' ip     : The client's IP (xxx.xxx.xxx.xxx:port)
    ' user   : The 'user' parameter value given when calling BASS_EncodeSetNotify

End Sub

Sub WMENCODEPROC(ByVal handle As Long, ByVal dtype As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long)

    'CALLBACK FUNCTION !!!

    ' Encoder callback function.
    ' handle : The encoder handle
    ' dtype  : The type of data, one of BASS_WMA_ENCODE_xxx values
    ' buffer : The encoded data
    ' length : Length of the data
    ' user   : The 'user' parameter value given when calling BASS_WMA_EncodeOpen

End Sub
