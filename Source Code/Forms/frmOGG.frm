VERSION 5.00
Begin VB.Form frmOGG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M3P : OGG Vorbis Tag"
   ClientHeight    =   4920
   ClientLeft      =   195
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "frmOGG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDetail 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmOGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fso As New FileSystemObject
Dim fFile As file

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strMsg As String
    Me.Icon = LoadResPicture(112, vbResIcon)
    txtFileName.text = NowPlaying(frmPlayList.List.Key(currentRIndex)).Infor.FullName
    
    Dim Ogg As New clsOGG
    Call Ogg.GetOGGTag(txtFileName.text)
    strMsg = "**************************OGG Informations**************************" & vbCrLf
    strMsg = strMsg & "Filename : " & txtFileName.text & vbCrLf
    strMsg = strMsg & "Author : " & Ogg.oArtist & vbCrLf
    strMsg = strMsg & "Album : " & Ogg.oAlbum & vbCrLf
    strMsg = strMsg & "Title : " & Ogg.oTitle & vbCrLf
    strMsg = strMsg & "Genre : " & Ogg.oGenre & vbCrLf
    strMsg = strMsg & "Track : " & Ogg.oTrackNumber & vbCrLf
    strMsg = strMsg & "Comment : " & Ogg.oComment & vbCrLf
    strMsg = strMsg & "Encoded : " & Ogg.oEncodedUsing & vbCrLf
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "*******************************File Mics*******************************" & vbCrLf
    strMsg = strMsg & "File type : " & fFile.Type & vbCrLf
    strMsg = strMsg & "File size : " & fFile.Size & " bytes" & vbCrLf
    strMsg = strMsg & "Date created : " & fFile.DateCreated & vbCrLf
    strMsg = strMsg & "Date last accessed : " & fFile.DateLastAccessed & vbCrLf
    strMsg = strMsg & "Date last modified : " & fFile.DateLastModified & vbCrLf
    txtDetail.text = strMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fFile = Nothing
End Sub
