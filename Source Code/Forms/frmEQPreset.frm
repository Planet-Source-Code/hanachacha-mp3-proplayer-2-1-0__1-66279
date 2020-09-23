VERSION 5.00
Begin VB.Form frmEQLoadPreset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P _ Load EQ Preset"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.ListBox lstPreset 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmEQLoadPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTempPreset As String
Public bolShow As Boolean
Private Sub cmdCancel_Click()
    If strTempPreset <> "" Then
        Call LoadEQPreset(strTempPreset)
    End If
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim strPreset As String
        strPreset = lstPreset.List(lstPreset.ListIndex)
        Call LoadEQPreset(strPreset)
    Unload Me
End Sub
Private Sub Form_Load()
    Dim strEQ As String
    
    bolShow = True
    Me.Left = frmMedia.Left + frmMedia.height
    Me.Top = frmMedia.Top
    
    strEQ = App.path & "\EQ\EqualizerPreset.epr"
    If Not FileExists(strEQ) Then
        MsgBox "Program not found Equalizer preset file !!!" & vbCrLf _
                & "You should use save Equalizer preset to make new file", _
                vbCritical, "MP3_proPlayer"
        Unload Me
    Else
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
        
        strTempPreset = strCurrentEQPreset
        For i = 0 To lstPreset.ListCount - 1
            If lstPreset.List(i) = strCurrentEQPreset Then
                lstPreset.Selected(i) = True
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bolShow = False
End Sub

Private Sub lstPreset_DblClick()
    Dim strPreset As String
    
    strPreset = lstPreset.List(lstPreset.ListIndex)
    Call LoadEQPreset(strPreset)
End Sub
