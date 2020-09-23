VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Track"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5520
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkUpdate 
      Caption         =   "Update tag if suported"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   4
      Left            =   3600
      TabIndex        =   12
      Top             =   1830
      Width           =   330
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1830
      Width           =   435
   End
   Begin VB.Label lblCaption 
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
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1035
      Width           =   435
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   630
      Width           =   390
   End
   Begin MSForms.TextBox txtTrack 
      Height          =   300
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2355;529"
      BorderColor     =   -2147483645
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTrack 
      Height          =   300
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2778;529"
      BorderColor     =   -2147483645
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTrack 
      Height          =   300
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1400
      Width           =   4695
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "8281;529"
      BorderColor     =   -2147483645
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTrack 
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1000
      Width           =   4695
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "8281;529"
      BorderColor     =   -2147483645
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTrack 
      Height          =   300
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "8281;529"
      BorderColor     =   -2147483645
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k() As Long
Dim bolMultiEdit As Boolean


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    If bolMultiEdit = False Then
        Library(CurrentTrack).Infor.Artist = txtTrack(0).text
        Library(CurrentTrack).Infor.Album = txtTrack(1).text
        Library(CurrentTrack).Infor.Title = txtTrack(2).text
        Library(CurrentTrack).Infor.Genre = txtTrack(3).text
        Library(CurrentTrack).Infor.Year = txtTrack(4).text
        Library(CurrentTrack).strDayUpdate = date
        
        Call WriteDataFile(CurrentTrack)
        If chkUpdate.value = vbChecked Then
            Select Case LCase(Library(CurrentTrack).Infor.ExtType)
                Case ".mp3"
                    Dim tagID3v1 As New clsID3v1
                    Dim tagID3v2 As New clsID3v2
                    With tagID3v1
                        .Artist = txtTrack(0).text
                        .Album = txtTrack(1).text
                        .Genre = ReturnGenreID(txtTrack(3).text)
                        .Title = txtTrack(2).text
                        .Year = txtTrack(4).text
                    End With
                    Call tagID3v1.WriteTag(Library(CurrentTrack).Infor.FullName)
                    
                    With tagID3v2
                        .Artist = txtTrack(0).text
                        .Album = txtTrack(1).text
                        .Genre = txtTrack(3).text
                        .Title = txtTrack(2).text
                        .Year = txtTrack(4).text
                    End With
                    Call tagID3v2.WriteID3v2Tag(Library(CurrentTrack).Infor.FullName)
                Case ".wma"
                    Dim tagWma As New clsWMA
                    With tagWma
                        .Artist = txtTrack(0).text
                        .Album = txtTrack(1).text
                        .Genre = txtTrack(3).text
                        .Title = txtTrack(2).text
                        .Year = txtTrack(4).text
                    End With
                    Call tagWma.WriteTag(Library(CurrentTrack).Infor.FullName)
                Case ".wmv", ".avi", ".mpg", ".mpe", "mpeg", ".asf"
                    Dim tagVID As New clsAVI
                    With tagVID
                        .Artist = txtTrack(0).text
                        .Title = txtTrack(2).text
                    End With
                    Call tagVID.WriteVideoTag(Library(CurrentTrack).Infor.FullName)
            End Select
        End If
    Else
        Dim i As Long
        For i = 0 To UBound(k) - 1
            If txtTrack(0).text <> "(multiple value)" And txtTrack(0).text <> "" Then
                Library(k(i)).Infor.Artist = txtTrack(0).text
            End If
            If txtTrack(1).text <> "(multiple value)" And txtTrack(1).text <> "" Then
                Library(k(i)).Infor.Album = txtTrack(1).text
            End If
            If txtTrack(2).text <> "(multiple value)" And txtTrack(2).text <> "" Then
                Library(k(i)).Infor.Title = txtTrack(2).text
            End If
            If txtTrack(3).text <> "(multiple value)" And txtTrack(3).text <> "" Then
                Library(k(i)).Infor.Genre = txtTrack(3).text
            End If
            If txtTrack(4).text <> "(multiple value)" And txtTrack(4).text <> "" Then
                Library(k(i)).Infor.Year = txtTrack(4).text
            End If
            Library(k(i)).strDayUpdate = date
            Call WriteDataFile(k(i))
        Next i
    End If
    
    Call frmLibrary.ReadLibrary
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo beep
    Dim i As Long
    Dim x As Long
    Dim lngCount As Long
    
    
    'AlwaysOnTop Me, True
    frmLibrary.Enabled = False
    
    lngCount = 0
    ReDim k(0)
    
    For i = 1 To frmLibrary.lvwFilter.ListItems.Count
        If frmLibrary.lvwFilter.ListItems(i).Selected = True Then
            lngCount = lngCount + 1
            ReDim Preserve k(lngCount)
            k(lngCount - 1) = TrackIndex(frmLibrary.lvwFilter.ListItems(i).SubItems(13))
        End If
    Next i
    If lngCount = 1 Then
        bolMultiEdit = False
    Else
        bolMultiEdit = True
    End If
    If bolMultiEdit = False Then
        txtTrack(0).text = Library(CurrentTrack).Infor.Artist
        txtTrack(1).text = Library(CurrentTrack).Infor.Album
        txtTrack(2).text = Library(CurrentTrack).Infor.Title
        txtTrack(3).text = Library(CurrentTrack).Infor.Genre
        txtTrack(4).text = Library(CurrentTrack).Infor.Year
    Else
        i = 0
        For i = 0 To UBound(k) - 1
            For x = UBound(k) - 1 To i + 1 Step -1
                If Library(k(x)).Infor.Artist <> Library(k(i)).Infor.Artist Then
                    txtTrack(0).text = "(multiple value)"
                    Exit For
                Else
                    txtTrack(0).text = Library(CurrentTrack).Infor.Artist
                End If
            Next x
        Next i
        For i = 0 To UBound(k) - 1
            For x = UBound(k) - 1 To i + 1 Step -1
                If Library(k(x)).Infor.Album <> Library(k(i)).Infor.Album Then
                    txtTrack(1).text = "(multiple value)"
                    Exit For
                Else
                    txtTrack(1).text = Library(CurrentTrack).Infor.Album
                End If
            Next x
        Next i
        For i = 0 To UBound(k) - 1
            For x = UBound(k) - 1 To i + 1 Step -1
                If Library(k(x)).Infor.Title <> Library(k(i)).Infor.Title Then
                    txtTrack(2).text = "(multiple value)"
                    Exit For
                Else
                    txtTrack(2).text = Library(CurrentTrack).Infor.Title
                End If
            Next x
        Next i
        For i = 0 To UBound(k) - 1
            For x = UBound(k) - 1 To i + 1 Step -1
                If Library(k(x)).Infor.Genre <> Library(k(i)).Infor.Genre Then
                    txtTrack(3).text = "(multiple value)"
                    Exit For
                Else
                    txtTrack(3).text = Library(CurrentTrack).Infor.Genre
                End If
            Next x
        Next i
        For i = 0 To UBound(k) - 1
            For x = UBound(k) - 1 To i + 1 Step -1
                If Library(k(x)).Infor.Year <> Library(k(i)).Infor.Year Then
                    txtTrack(4).text = "(multiple value)"
                    Exit For
                Else
                    txtTrack(4).text = Library(CurrentTrack).Infor.Year
                End If
            Next x
        Next i
    End If
beep:
    If Err.Number <> 0 Then
        Debug.Print Err.Description & " " & Err.Source
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AlwaysOnTop Me, False
    frmLibrary.Enabled = True
End Sub
