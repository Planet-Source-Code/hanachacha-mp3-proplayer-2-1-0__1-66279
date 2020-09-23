VERSION 5.00
Begin VB.Form frmEQSavePreset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P _ Save EQ Preset"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtPreset 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3855
   End
   Begin VB.ListBox lstPreset 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmEQSavePreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject
Public bolShow As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim strPreset As String
    
    strPreset = txtPreset.text
    Call SaveEQPreset(strPreset)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strEQ As String
    bolShow = True
    Me.Left = frmMedia.Left + frmMedia.height
    Me.Top = frmMedia.Top
    
    strEQ = App.Path & "\EQ\EqualizerPreset.epr"
    
    If Not FileExists(strEQ) Then
        fso.CreateTextFile strEQ
    End If
    
    Dim i As Integer
    Dim fn As Long
    Dim preset As String
    'Add Equalizer preset
    fn = FreeFile
    Open strEQ For Input As #fn
        Do Until EOF(fn)
            Line Input #fn, preset
            If Mid(preset, 1, 1) = "[" And Right(preset, 1) = "]" And Len(preset) > 2 Then
                lstPreset.AddItem Mid(preset, 2, Len(preset) - 2)
            End If
        Loop
    Close #fn
    
    For i = 0 To lstPreset.ListCount - 1
        If lstPreset.List(i) = strCurrentEQPreset Then
            lstPreset.Selected(i) = True
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bolShow = False
End Sub

Private Sub lstPreset_Click()
    txtPreset.text = lstPreset.List(lstPreset.ListIndex)
End Sub
