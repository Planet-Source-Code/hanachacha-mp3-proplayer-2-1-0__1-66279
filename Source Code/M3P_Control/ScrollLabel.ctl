VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.UserControl ScrollLabel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ToolboxBitmap   =   "ScrollLabel.ctx":0000
   Begin VB.Timer tmrInvert 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   4080
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   4080
   End
   Begin MSForms.Label lblText 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   225
      BackColor       =   -2147483643
      VariousPropertyBits=   19
      Size            =   "397;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "ScrollLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim bolScroll As Boolean

Dim strText As String
Dim pos As Integer

Private Sub lblText_Click()
    tmrScroll.Enabled = False
    tmrInvert.Enabled = False
    bolScroll = Not bolScroll
    If bolScroll Then
        lblText.Caption = "*** " & strText & " "
    Else
        lblText.Caption = strText
        lblText.Left = 0
    End If
    Dim tmp As String
    Dim x As Long
    
    For x = 0 To Len(strText)
        tmp = tmp & "_"
    Next x
    
    lblText.Width = UserControl.TextWidth(tmp)
    If lblText.Width < UserControl.ScaleWidth Then lblText.Width = UserControl.ScaleWidth
    If lblText.Height > UserControl.ScaleHeight Then lblText.Height = UserControl.ScaleHeight
    lblText.top = UserControl.ScaleHeight - lblText.Height
    If lblText.Width > UserControl.ScaleWidth Then
        tmrScroll.Enabled = True
    End If
End Sub

Private Sub tmrInvert_Timer()
    tmrInvert.Interval = 100
    lblText.Left = lblText.Left + 1
    If lblText.Left >= 0 Then
        tmrScroll.Interval = 2000
        tmrScroll.Enabled = True
        tmrInvert.Enabled = False
    End If
End Sub

Private Sub tmrScroll_Timer()
    On Error Resume Next
    If bolScroll Then
        tmrScroll.Interval = 250
        Dim x As String
        Dim Y As String
        x = Left(lblText.Caption, 1)
        Y = Right(lblText.Caption, Len(lblText.Caption) - 1)
        lblText.Caption = Y + x
    Else
        tmrScroll.Interval = 100
        lblText.Left = lblText.Left - 1
        If (0 - lblText.Left) + UserControl.ScaleWidth <= lblText.Width Then
            tmrInvert.Interval = 2000
            tmrInvert.Enabled = True
            tmrScroll.Enabled = False
        End If
    End If
End Sub
Public Property Let Title(strTitle As String)
    Dim x As Long
    Dim tmp As String
    
    tmrScroll.Enabled = False
    tmrInvert.Enabled = False
    strText = strTitle
    
    For x = 0 To Len(strText)
        tmp = tmp & "_"
    Next x
    
    lblText.Width = UserControl.TextWidth(tmp)
    If lblText.Width < UserControl.ScaleWidth Then lblText.Width = UserControl.ScaleWidth
    lblText.Caption = strTitle
    lblText.Left = 0
    If lblText.Width > UserControl.ScaleWidth Then
        lblText.Caption = "*** " & lblText.Caption & " "
        tmrScroll.Enabled = True
    End If
    PropertyChanged "Title"
End Property
Public Property Get Title() As String
    Title = strText
End Property
Public Property Set Font(ByVal NewFont As StdFont)
    Set lblText.Font = NewFont
    Set UserControl.Font = NewFont
    PropertyChanged "Font"
End Property
Public Property Get Font() As StdFont
    Set Font = lblText.Font
End Property
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    lblText.ForeColor = NewColor
    PropertyChanged "ForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblText.ForeColor
End Property
Public Property Let FontBold(ByVal bolBold As Boolean)
    lblText.Font.Bold = bolBold
    UserControl.Font.Bold = bolBold
    PropertyChanged "FontBold"
End Property
Public Property Get FontBold() As Boolean
    FontBold = lblText.Font.Bold
End Property
Public Property Let Size(ByVal NewSize As Long)
    lblText.Font.Size = NewSize
    UserControl.Font.Size = NewSize
    PropertyChanged "Size"
End Property
Private Sub UserControl_Click()
    bolScroll = Not bolScroll
    lblText.Caption = strText
End Sub

Private Sub UserControl_Resize()
    If lblText.Height > UserControl.ScaleHeight Then lblText.Height = UserControl.ScaleHeight
    lblText.top = UserControl.ScaleHeight - lblText.Height
End Sub

Public Sub Draw(pic As StdPicture, xS As Long, yS As Long, widthS As Long, heightS As Long)
    UserControl.Cls
    UserControl.PaintPicture pic, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, xS, yS, widthS, heightS, vbSrcCopy
End Sub
Public Property Let Scroll(bol As Boolean)
    bolScroll = bol
    PropertyChanged "Scroll"
End Property
Public Property Get Scroll() As Boolean
    Scroll = bolScroll
End Property

