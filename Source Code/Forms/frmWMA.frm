VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmWMA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M3P : WMA - Tag Editor"
   ClientHeight    =   4905
   ClientLeft      =   225
   ClientTop       =   360
   ClientWidth     =   5055
   Icon            =   "frmWMA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmHeader 
      Caption         =   "Mics"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   4815
      Begin VB.Label lblDetail 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "U&ndo"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   13
      Top             =   2213
      Width           =   435
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2213
      Width           =   330
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1733
      Width           =   300
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1253
      Width           =   435
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   773
      Width           =   390
   End
   Begin MSForms.TextBox txtGenre 
      Height          =   300
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2990;529"
      BorderColor     =   12164479
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtYear 
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2355;529"
      BorderColor     =   12164479
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTitle 
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "7435;529"
      BorderColor     =   12164479
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtAlbum 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "7435;529"
      BorderColor     =   12164479
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtArtist 
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "7435;529"
      BorderColor     =   12164479
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmWMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WMA As New clsWMA
Dim tmpWMA As New clsWMA

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUndo_Click()
    Set WMA = tmpWMA
    Call WMA.WriteTag(txtFileName.text)
    txtArtist.text = WMA.Artist
    txtAlbum.text = WMA.Album
    txtTitle.text = WMA.Title
    txtYear.text = WMA.Year
    txtGenre.text = WMA.Genre
End Sub

Private Sub cmdUpdate_Click()
    With WMA
        .Album = txtAlbum.text
        .Artist = txtArtist.text
        .Title = txtTitle.text
        .Genre = txtGenre.text
        .Year = txtYear.text
    End With
    Call WMA.WriteTag(txtFileName.text)
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Icon = LoadResPicture(112, vbResIcon)
    txtFileName.text = NowPlaying(frmPlayList.List.Key(currentRIndex)).Infor.FullName
    Set fFile = fso.GetFile(txtFileName.text)
    
    Dim strMsg As String
    Call WMA.Get_WMA_Header(txtFileName.text)
    Set tmpWMA = WMA
    
    txtArtist.text = WMA.Artist
    txtAlbum.text = WMA.Album
    txtTitle.text = WMA.Title
    txtYear.text = WMA.Year
    txtGenre.text = WMA.Genre
    Debug.Print WMA.Genre & "AAAA"
    strMsg = strMsg & "WM/GenreID : " & WMA.GenreID & vbCrLf
    strMsg = strMsg & "Bitrate : " & WMA.bitrate & " Kbps" & vbCrLf
    strMsg = strMsg & "Frequency : " & WMA.Frequency & " Khz" & vbCrLf
    strMsg = strMsg & "Lenght : " & WMA.length & " secconds" & vbCrLf
    strMsg = strMsg & "WMFSDKVersion : " & WMA.SDKVer & vbCrLf
    strMsg = strMsg & "WMFSDKNeeded : " & WMA.SDKNeeded & vbCrLf
    lblDetail.Caption = strMsg
End Sub
