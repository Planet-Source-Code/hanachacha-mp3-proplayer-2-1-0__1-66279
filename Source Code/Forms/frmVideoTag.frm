VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{E36A7FF8-53F4-4B09-B66A-40EEA5290C09}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmVideo_Infor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M3P : DirectShow Information"
   ClientHeight    =   4200
   ClientLeft      =   285
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmVideoTag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUndo 
      Caption         =   "U&ndo"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin M3P_Control.Video Video 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtFilename 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Frame frmBorder 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   653
         Width           =   300
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   293
         Width           =   495
      End
      Begin MSForms.TextBox txtTitle 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   4575
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "8070;529"
         BorderColor     =   12164479
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtArtist 
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   4575
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "8070;529"
         BorderColor     =   12164479
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblDetail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Infor"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmVideo_Infor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject
Dim fFile As file
Dim VID As New clsAVI
Dim tmpVID As New clsAVI

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdUndo_Click()
    Set VID = tmpVID
    Call VID.ReadVideo(txtFileName.text)
    txtArtist.text = VID.Artist
    txtTitle.text = VID.Title
End Sub

Private Sub cmdUpdate_Click()
    If NowPlaying(frmPlayList.List.Key(currentRIndex)).Infor.ExtType <> ".mpg" Or NowPlaying(frmPlayList.List.Key(currentRIndex)).Infor.ExtType <> ".mpeg" Then
        With VID
            .Artist = txtArtist.text
            .Title = txtTitle.text
        End With
        Call VID.WriteVideoTag(txtFileName.text)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Icon = LoadResPicture(112, vbResIcon)
    txtFileName.text = NowPlaying(frmPlayList.List.Key(currentRIndex)).Infor.FullName
    Set fFile = fso.GetFile(txtFileName.text)
    
    Dim strMsg As String
    
    Video.OpenVideo txtFileName.text
    VID.ReadVideo (txtFileName.text)
    Set tmpVID = VID
    
    txtArtist.text = VID.Artist
    txtTitle.text = VID.Title
    
    strMsg = "Video size : " & Video.ScrW & " x " & Video.ScrH & vbCrLf
    strMsg = strMsg & "Audio : " & VID.bitrate \ 1000 & " Kbps" & vbCrLf
    strMsg = strMsg & "Lenght : " & Video.Duration & " secconds" & vbCrLf
    
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "File type : " & fFile.Type & vbCrLf
    strMsg = strMsg & "File size : " & fFile.Size & " bytes" & vbCrLf
    strMsg = strMsg & "Date created : " & fFile.DateCreated & vbCrLf
    strMsg = strMsg & "Date last accessed : " & fFile.DateLastAccessed & vbCrLf
    strMsg = strMsg & "Date last modified : " & fFile.DateLastModified & vbCrLf
    lblDetail.Caption = strMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fFile = Nothing
    Set fso = Nothing
    Video.CloseVideo
End Sub
