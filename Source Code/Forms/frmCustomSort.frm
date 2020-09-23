VERSION 5.00
Begin VB.Form frmCustomSort 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P _ Advanced sort"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5160
   Icon            =   "frmCustomSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtSortString 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCustomSort.frx":000C
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "frmCustomSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub cmdCancel_Click()
    txtSortString.text = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim strSorted As String
    Dim strFilePlaying As String
    
    strSorted = txtSortString.text
    Call SetSortString(strSorted)
    
    frmPlayList.List.IsSort = True
    
    If currentPlayIndex <> 0 Then
        currentPlayIndex = frmPlayList.List.CurrentPlayItem
    End If
    
    tPlaylistConfig.strSortString = strSorted
    Unload Me
End Sub

Private Sub Form_Load()
    Call AlwaysOnTop(Me, True)
    Dim tmp As String
    tmp = tPlaylistConfig.strSortString
    If tmp <> "" Then txtSortString.text = tmp
End Sub
