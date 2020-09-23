Attribute VB_Name = "modXPTheme"
'+++++++++++++++++++++++++++++++++++++++++++
'+ Author : Phuc.H Truong aka <Hanachacha> +
'+++++++++++++++++++++++++++++++++++++++++++
'From MSDN - Microsoft
Option Explicit

Function XPStyle(Optional AutoRestart As Boolean = True, Optional CreateNew As Boolean) As Boolean
    InitCommonControls

    On Error Resume Next
         Dim XML             As String
         Dim ManifestCheck   As String
         Dim strManifest     As String
         Dim FreeFileNo      As Integer
        
        If AutoRestart = True Then CreateNew = False
        
        XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & _
          "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & _
          "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & _
          "    processorArchitecture=""x86"" " & vbCrLf & _
          "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & _
          "    name=""EXENAME""/>" & vbCrLf & _
          "    <description>EXEDESCRIBTION</description>" & vbCrLf & _
          "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & _
          "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & _
          "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
          "         version=""6.0.0.0""" & vbCrLf & _
          "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
          "         language=""*""" & vbCrLf & _
          "         processorArchitecture=""x86""/>" & vbCrLf & _
          "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & _
          "</assembly>" & vbCrLf & "")
        
        strManifest = App.path & "\" & App.EXEName & ".exe.manifest"
        
        ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)
        
        If ManifestCheck = "" Or CreateNew = True Then
          XML = Replace(XML, "EXENAME", App.EXEName & ".exe")
          XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
          XML = Replace(XML, "EXEDESCRIBTION", App.FileDescription)
          
          FreeFileNo = FreeFile
          
          If ManifestCheck <> "" Then
            SetAttr strManifest, vbNormal
            Kill strManifest
          End If
          
          Open strManifest For Binary As #(FreeFileNo)
             Put #(FreeFileNo), , XML
          Close #(FreeFileNo)
          
          SetAttr strManifest, vbHidden + vbSystem
          
          If ManifestCheck = "" Then
            XPStyle = False
          Else
            XPStyle = True
          End If
          
          If AutoRestart = True Then
            Shell App.path & "\" & App.EXEName & ".exe", vbNormalFocus
            End
          End If
          
        Else
          XPStyle = True
        End If

End Function
