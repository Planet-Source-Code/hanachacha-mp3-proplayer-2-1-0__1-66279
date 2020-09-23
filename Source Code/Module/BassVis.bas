Attribute VB_Name = "modBassVis"
Option Explicit

' BASS_SONIQUEVIS_CreateVis flags
Global Const BASS_VIS_NOINIT = 1

' BASS_SONIQUEVIS_SetConfig flags
Global Const BASS_VIS_CONFIG_FFTAMP = 1
Global Const BASS_VIS_CONFIG_FFT_SKIPCOUNT = 2 ' Skip count range is from 0 to 3 (because of limited FFT request size)
Global Const BASS_VIS_CONFIG_WAVE_SKIPCOUNT = 3 ' Skip count range is from 0 to (...) try it out, whenever Bass crashes or does not return enough sample data
Global Const BASS_VIS_CONFIG_SONIQUE_SLOWFADE = 4 ' Dim light colors to less than half, then slowly fade them out

' Bass FFT Amplification values
Global Const BASS_VIS_FFTAMP_NORMAL = 1
Global Const BASS_VIS_FFTAMP_HIGH = 2
Global Const BASS_VIS_FFTAMP_HIGHER = 3
Global Const BASS_VIS_FFTAMP_HIGHEST = 4

' BASS_VIS_FindPlugin flags
Global Const BASS_VIS_FIND_SONIQUE = 1
Global Const BASS_VIS_FIND_WINAMP = 2
Global Const BASS_VIS_FIND_RECURSIVE = 4
' return value type
Global Const BASS_VIS_FIND_COMMALIST = 8
  ' Delphi's comma list style (item1,item2,"item 3",item4,"item with space")
  ' the list ends with single NULL character


Declare Function BASS_SONIQUEVIS_CreateVis Lib "bass_vis.dll" (ByVal f As Any, ByVal visconfig As Any, ByVal flags As Long, ByVal w As Long, ByVal h As Long) As Long
Declare Function BASS_SONIQUEVIS_Render Lib "bass_vis.dll" (ByVal handle As Long, ByVal channel As Long, ByVal canvas As Long) As Long
Declare Function BASS_SONIQUEVIS_Render2 Lib "bass_vis.dll" (ByVal handle As Long, ByVal data As Long, ByVal fft As Long, ByVal canvas As Long, ByVal flags As Long, ByVal pos As Long) As Long
Declare Function BASS_SONIQUEVIS_Free Lib "bass_vis.dll" (ByVal handle As Long) As Long
Declare Function BASS_SONIQUEVIS_GetName Lib "bass_vis.dll" (ByVal handle As Long) As Long
Declare Function BASS_SONIQUEVIS_Clicked Lib "bass_vis.dll" (ByVal handle As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function BASS_SONIQUEVIS_Resize Lib "bass_vis.dll" (ByVal handle As Long, ByVal nw As Long, ByVal nh As Long) As Long
Declare Sub BASS_SONIQUEVIS_CreateFakeSoniqueWnd Lib "bass_vis.dll" ()
Declare Sub BASS_SONIQUEVIS_DestroyFakeSoniqueWnd Lib "bass_vis.dll" ()

Declare Function BASS_WINAMPVIS_CreateVis Lib "bass_vis.dll" (ByVal f As Any, ByVal module As Long, ByVal flags As Long) As Long
Declare Function BASS_WINAMPVIS_Render Lib "bass_vis.dll" (ByVal handle As Long, ByVal channel As Long) As Long
Declare Function BASS_WINAMPVIS_Render2 Lib "bass_vis.dll" (ByVal handle As Long, ByVal data As Long, ByVal fft As Long, ByVal flags As Long, ByVal rate As Long) As Long
Declare Function BASS_WINAMPVIS_Free Lib "bass_vis.dll" (ByVal handle As Long) As Long
Declare Function BASS_WINAMPVIS_GetName Lib "bass_vis.dll" (ByVal handle As Long) As Long
Declare Function BASS_WINAMPVIS_GetModuleName Lib "bass_vis.dll" (ByVal handle As Long) As Long
Declare Function BASS_WINAMPVIS_Config Lib "bass_vis.dll" (ByVal handle As Long) As Long
Declare Function BASS_WINAMPVIS_SetChanInfo Lib "bass_vis.dll" (ByVal handle As Long, ByVal title As Any, ByVal pos As Long, ByVal length As Long) As Long

Declare Function BASS_VIS_GetConfig Lib "bass_vis.dll" (ByVal opt As Long) As Long
Declare Sub BASS_VIS_SetConfig Lib "bass_vis.dll" (ByVal opt As Long, ByVal value As Long)
Declare Function BASS_VIS_FindPlugins Lib "bass_vis.dll" (ByVal vispath As Any, ByVal flags As Long) As Long
