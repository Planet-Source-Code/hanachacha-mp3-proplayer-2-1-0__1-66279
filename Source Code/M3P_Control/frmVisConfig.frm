VERSION 5.00
Begin VB.Form frmVisConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M3P _ Visualization"
   ClientHeight    =   225
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuVisual 
      Caption         =   "Visualization"
      Visible         =   0   'False
      Begin VB.Menu mnuVisStyle 
         Caption         =   "Oscilliscope"
         Index           =   0
      End
      Begin VB.Menu mnuVisStyle 
         Caption         =   "Spectrum"
         Index           =   1
      End
      Begin VB.Menu mnuVisStyle 
         Caption         =   "None"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmVisConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuVisStyle_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 2
        mnuVisStyle(i).Checked = False
    Next i
    intStyle = Index
    mnuVisStyle(Index).Checked = True
End Sub
