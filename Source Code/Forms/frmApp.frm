VERSION 5.00
Begin VB.Form frmApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3_ProPlayer 2.1.0"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   ControlBox      =   0   'False
   Icon            =   "frmApp.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmApp"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Call Hook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Unhook(Me.hwnd)
End Sub

