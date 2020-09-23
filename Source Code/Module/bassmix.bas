' BASSmix 2.3 Visual Basic API, (c) 2005-2006 Ian Luck.
' Requires BASS - available @ www.un4seen.com

' See the BASSMIX.CHM file for detailed documentation

Attribute VB_Name = "BASSmix"

' BASS_Mixer_StreamCreate flags
Global Const BASS_MIXER_END     = &H10000  ' end the stream when there are no sources
Global Const BASS_MIXER_NONSTOP = &H20000  ' don't stall when there are no sources

' BASS_Mixer_StreamAddChannel flags
Global Const BASS_MIXER_MATRIX  = &H10000  ' matrix mixing
Global Const BASS_MIXER_DOWNMIX = &H400000 ' downmix to stereo/mono

Declare Function BASS_Mixer_StreamCreate Lib "bassmix.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long) As Long
Declare Function BASS_Mixer_StreamAddChannel Lib "bassmix.dll" (ByVal handle As Long, ByVal channel As Long, ByVal flags As Long) As Long
Declare Function BASS_Mixer_StreamAddChannelEx Lib "bassmix.dll" (ByVal handle As Long, ByVal channel As Long, ByVal flags As Long, ByVal start As Long, ByVal length As Long) As Long

Declare Function BASS_Mixer_ChannelRemove Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Mixer_ChannelGetMixer Lib "bassmix.dll" (ByVal handle As Long) As Long
Private Declare Function BASS_Mixer_ChannelSetPosition64 Lib "bassmix.dll" Alias "BASS_Mixer_ChannelSetPosition" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long) As Long
Declare Function BASS_Mixer_ChannelGetPosition Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Mixer_ChannelSetMatrix Lib "bassmix.dll" (ByVal handle As Long, ByRef matrix As Single) As Long
Declare Function BASS_Mixer_ChannelGetMatrix Lib "bassmix.dll" (ByVal handle As Long, ByRef matrix As Single) As Long

Function BASS_Mixer_ChannelSetPosition(ByVal handle As Long, ByVal pos As Long) As Long
BASS_Mixer_ChannelSetPosition = BASS_Mixer_ChannelSetPosition64(handle, pos, 0)
End Function
