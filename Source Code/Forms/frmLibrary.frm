VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0179B2D7-CD62-439D-BE78-CF820F5A4B44}#1.0#0"; "M3P_Control.ocx"
Begin VB.Form frmLibrary 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "M3P _ Library"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   Icon            =   "frmLibrary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00E35400&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9495
      TabIndex        =   48
      Top             =   8040
      Width           =   9495
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Library"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList imgLibrary 
      Left            =   720
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin VB.PictureBox picLibrary 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   0
      ScaleHeight     =   537
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.PictureBox picRate 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   6480
         ScaleHeight     =   135
         ScaleWidth      =   255
         TabIndex        =   53
         Top             =   6960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   3120
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton cmdInput 
            Caption         =   "&Cancel"
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   37
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdInput 
            Caption         =   "&Ok"
            Height          =   375
            Index           =   0
            Left            =   3240
            TabIndex        =   36
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00DD8055&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter new playlist name :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DD8055&
            Caption         =   " Create New Playlist"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   5175
         End
         Begin MSForms.TextBox txtInput 
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   4935
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "8705;661"
            BorderColor     =   -2147483645
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.PictureBox picToolbar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E35400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   633
         TabIndex        =   2
         Top             =   0
         Width           =   9495
         Begin VB.CommandButton cmdFind 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Search"
            Height          =   375
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdRemove 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Remove"
            Height          =   375
            Left            =   1335
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   1320
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Add"
            Height          =   375
            Left            =   0
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   1335
         End
         Begin MSForms.TextBox txtSearch 
            Height          =   360
            Left            =   4200
            TabIndex        =   38
            Top             =   120
            Width           =   4575
            VariousPropertyBits=   747653147
            BackColor       =   16777215
            ForeColor       =   14898176
            BorderStyle     =   1
            Size            =   "8070;635"
            BorderColor     =   12582912
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.HScrollBar sldListHor 
         Height          =   255
         Left            =   2640
         TabIndex        =   52
         Top             =   6360
         Width           =   6135
      End
      Begin VB.VScrollBar sldListVer 
         Height          =   5775
         Left            =   12120
         TabIndex        =   51
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H00E35400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   2640
         MousePointer    =   9  'Size W E
         ScaleHeight     =   385
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   2
         TabIndex        =   3
         Top             =   600
         Width           =   30
      End
      Begin M3P_Control.Video VideoPreview 
         Height          =   0
         Left            =   2640
         TabIndex        =   45
         Top             =   6720
         Visible         =   0   'False
         Width           =   0
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin VB.CommandButton cmdTree 
         BackColor       =   &H00F0F0F0&
         Caption         =   "My Library"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         MaskColor       =   &H80000016&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   2655
      End
      Begin VB.PictureBox picMaster 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   5775
         Left            =   2640
         ScaleHeight     =   385
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   633
         TabIndex        =   4
         Top             =   600
         Width           =   9495
         Begin VB.PictureBox picList 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   4575
            Left            =   0
            ScaleHeight     =   305
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   633
            TabIndex        =   5
            Top             =   0
            Width           =   9495
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   9000
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   3240
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   41
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   3120
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   34
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   3000
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   7
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   2760
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   8
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   2520
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   9
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   2280
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   10
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   2040
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   6
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   1800
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   11
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   1560
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   12
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   1320
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   13
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   1080
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   14
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   840
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   15
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   600
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   16
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   360
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   17
               Top             =   1080
               Width           =   15
            End
            Begin VB.PictureBox picColWidth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E35400&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   120
               MousePointer    =   9  'Size W E
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   18
               Top             =   1080
               Width           =   15
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   0
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   600
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   1200
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   1800
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   2400
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   3000
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   3600
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   4200
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   4800
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   5400
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   6000
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   6600
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   7200
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   7800
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CommandButton cmdColumnHeader 
               BackColor       =   &H00F0F0F0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   8400
               MaskColor       =   &H00FF8080&
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   615
            End
         End
         Begin MSComctlLib.ListView lvwFilter 
            Height          =   3135
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   5530
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Artist"
               Object.Width           =   3969
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Album"
               Object.Width           =   3969
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Title"
               Object.Width           =   3969
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Genre"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Year"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Bitrate"
               Object.Width           =   1984
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Duration"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Frequency"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Rate"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Date add"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Date Update"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   11
               Text            =   "Play Count"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Name"
               Object.Width           =   3969
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Filepath"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   14
               Text            =   "Size"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView tvwLibrary 
         Height          =   5535
         Left            =   0
         TabIndex        =   1
         Top             =   855
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   9763
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         SingleSel       =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bolShow As Boolean
Public bolPlayingNow As Boolean
Public CurrentNode As Node

'Arr for Library Data
Private tmpData() As String
Private arrPL() As String
Private Artist() As String
Private Album() As String
Private Genre() As String
Private Year() As String
Private ArtistV() As String
Private YearV() As String
Private GenreV() As String

'Var for Resize TreeView
Private bolResize As Boolean
Private lngOldX As Long
'Var for Resize Column(s) in Library
Private bolResizeCol(14) As Boolean
Private bolAutoResize As Boolean
Private lngOldXCol(14) As Long

'Var test MouseRDown in Library
Private bolLibRdown As Boolean

'This type useful to check ? node selected
Private Type NodeLibrary
    Refer As String
    Name As String
    Type As Integer '0 is Audio,1 is Video,2 is other
End Type

Private OldInputDownX As Long
Private OldInputDownY As Long

Private bolTreeRDown As Boolean
Private CurrentSelect As Long
Private intStart As Long, intEnd As Long

Private Function FindNode(strNode As String) As Long
    On Error GoTo beep
    Dim i As Long
    For i = 1 To tvwLibrary.Nodes.Count
        If tvwLibrary.Nodes(i).text = strNode Then
            FindNode = i
            Exit For
        End If
    Next i
beep:
    If Err.Number <> 0 Then
        FindNode = 0
    End If
End Function
Public Sub ReadLibrary()
    On Error Resume Next
    
    ReDim Artist(0)
    ReDim Album(0)
    ReDim Genre(0)
    ReDim Year(0)
    ReDim ArtistV(0)
    ReDim GenreV(0)
    ReDim YearV(0)
    
    Dim x As Long
    Dim i As Long
    Dim y As Long
    Dim lngArtist As Long
    Dim lngArtistV As Long
    Dim lngAlbum As Long
    Dim lngGenre As Long
    Dim lngGenreV As Long
    Dim lngYear As Long
    Dim lngYearV As Long
    
    lngArtist = 0
    lngArtistV = 0
    lngAlbum = 0
    lngGenre = 0
    lngGenreV = 0
    lngYear = 0
    lngYearV = 0
    
    Dim f As Boolean
    Dim tStr As String

    'Set array Artist Audio & Artist Video
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Audio" Then
            lngArtist = lngArtist + 1
            ReDim Preserve tmpData(lngArtist)
            tmpData(lngArtist - 1) = Library(i).Infor.Artist
        End If
    Next i
    i = 0
    lngArtist = 0
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(Artist) - 1
            If tStr = Artist(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If Artist was found in list
            lngArtist = lngArtist + 1
            ReDim Preserve Artist(lngArtist)
            Artist(lngArtist - 1) = tStr
        End If
        i = i + 1
    Wend
    
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Video" Then
            lngArtistV = lngArtistV + 1
            ReDim Preserve tmpData(lngArtistV)
            tmpData(lngArtistV - 1) = Library(i).Infor.Artist
        End If
    Next i
    i = 0
    lngArtistV = 0
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(ArtistV) - 1
            If tStr = ArtistV(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If ArtistV was found in list
            lngArtistV = lngArtistV + 1
            ReDim Preserve ArtistV(lngArtistV)
            ArtistV(lngArtistV - 1) = tStr
        End If
        i = i + 1
    Wend
    
    
    'Set array Album Audio
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Audio" Then
            lngAlbum = lngAlbum + 1
            ReDim Preserve tmpData(lngAlbum)
            tmpData(lngAlbum - 1) = Library(i).Infor.Album
        End If
    Next i
    lngAlbum = 0
    i = 0
    'ReDim Album
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(Album) - 1
            If tStr = Album(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If album was found in list
            lngAlbum = lngAlbum + 1
            ReDim Preserve Album(lngAlbum)
            Album(lngAlbum - 1) = tStr
        End If
        i = i + 1
    Wend
    
    
    'Set array Genre Audio & Genre Video
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Audio" Then
            lngGenre = lngGenre + 1
            ReDim Preserve tmpData(lngGenre)
            tmpData(lngGenre - 1) = Library(i).Infor.Genre
        End If
    Next i
    lngGenre = 0
    i = 0
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(Genre) - 1
            If tStr = Genre(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If Genre was found in list
            lngGenre = lngGenre + 1
            ReDim Preserve Genre(lngGenre)
            Genre(lngGenre - 1) = tStr
        End If
        i = i + 1
    Wend
    
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Video" Then
            lngGenreV = lngGenreV + 1
            ReDim Preserve tmpData(lngGenreV)
            tmpData(lngGenreV - 1) = Library(i).Infor.Genre
        End If
    Next i
    lngGenreV = 0
    i = 0
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(GenreV) - 1
            If tStr = GenreV(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If GenreV was found in list
            lngGenreV = lngGenreV + 1
            ReDim Preserve GenreV(lngGenreV)
            GenreV(lngGenreV - 1) = tStr
        End If
        i = i + 1
    Wend

    'Set array Year Audio & Year Video
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Audio" Then
            lngYear = lngYear + 1
            ReDim Preserve tmpData(lngYear)
            tmpData(lngYear - 1) = Library(i).Infor.Year
        End If
    Next i
    lngYear = 0
    i = 0
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(Year) - 1
            If tStr = Year(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If year was found in list
            lngYear = lngYear + 1
            ReDim Preserve Year(lngYear)
            Year(lngYear - 1) = tStr
        End If
        i = i + 1
    Wend
    
    ReDim tmpData(0)
    For i = 0 To UBound(Library) - 1
        If Library(i).strType = "Video" Then
            lngYearV = lngYearV + 1
            ReDim Preserve tmpData(lngYearV)
            tmpData(lngYearV - 1) = Library(i).Infor.Year
        End If
    Next i
    lngYearV = 0
    i = 0
    While i < UBound(tmpData)
        f = False
        tStr = tmpData(i)
        For y = 0 To UBound(YearV) - 1
            If tStr = YearV(y) Then
                f = True
                Exit For
            End If
        Next y
        If Not f Then   'If yearV was found in list
            lngYearV = lngYearV + 1
            ReDim Preserve YearV(lngYearV)
            YearV(lngYearV - 1) = tStr
        End If
        i = i + 1
    Wend
    
    
    Dim dblSize As Double
    Dim lngLenght As Long
    dblSize = 0
    lngLenght = 0
    For i = 0 To UBound(Library) - 1
        dblSize = dblSize + ((Library(i).Infor.Size / 1024) / 1024)
        lngLenght = lngLenght + Library(i).Infor.Duration
    Next i
    lblStatus.Caption = "Library :" & UBound(Library) & " Item(s)"
    lblStatus.Caption = lblStatus.Caption & " --- Total: " & Time2String(lngLenght)
    lblStatus.Caption = lblStatus.Caption & " / " & Round(dblSize, 2) & " MB"
    Call AddNodeLibrary
End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PopupMenu frmMenu.mnuLibAdd, vbPopupMenuLeftAlign, 0, 33
    End If
End Sub

Private Sub cmdColumnHeader_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        bolLibRdown = False
    End If
    If Button = vbRightButton Then
        bolLibRdown = True
    End If
End Sub

Private Sub cmdColumnHeader_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If bolAutoResize = False Then
        If bolLibRdown = False Then
            If lvwFilter.SortKey = Index Then
                If lvwFilter.SortOrder = lvwAscending Then
                    lvwFilter.SortOrder = lvwDescending
                Else
                    lvwFilter.SortOrder = lvwAscending
                End If
            Else
                lvwFilter.SortKey = Index
            End If
            lvwFilter.Sorted = True
            Call DrawList
            Dim i As Long
            Dim z As Long
            z = (lvwFilter.GetFirstVisible.Index)
            For i = z To z + ItemPage
                If lvwFilter.ListItems(i).Selected = True Then
                    Call DrawSelected(i)
                End If
            Next i
        Else
            PopupMenu frmMenu.mnuView, vbPopupMenuLeftAlign
        End If
    End If
End Sub

Private Sub cmdFind_Click()
    Dim i As Long
    Dim x As Long
    Dim tStr As String
    bolPlayingNow = False
    lvwFilter.ListItems.Clear
    tStr = LCase(txtSearch.text)
    If tStr <> "" Then
        For i = 0 To UBound(Library) - 1
            If InStr(1, LCase(Library(i).Infor.Artist), tStr) Then
                Call ViewData(i)
            Else
                If InStr(1, LCase(Library(i).Infor.Album), tStr) Then
                    Call ViewData(i)
                Else
                    If InStr(1, LCase(Library(i).Infor.Title), tStr) Then
                        Call ViewData(i)
                    Else
                        If InStr(1, LCase(Library(i).Infor.Genre), tStr) Then
                            Call ViewData(i)
                        Else
                            If InStr(1, LCase(Library(i).Infor.Year), tStr) Then
                                Call ViewData(i)
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If
    Call DrawList
    Call UpdateStatus
End Sub

Private Sub cmdInput_Click(Index As Integer)
    Select Case Index
        Case 0
            Call CreateNewPlaylist(txtInput.text)
            picInput.Visible = False
        Case 1
            picInput.Visible = False
    End Select
End Sub

Private Sub cmdRemove_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PopupMenu frmMenu.mnuLibSub, vbPopupMenuLeftAlign, 89, 33
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim firstVisible As Long
    Dim tmp As Integer
    
    firstVisible = lvwFilter.GetFirstVisible.Index
    tmp = firstVisible
    If tmp > 0 And tmp <= lvwFilter.ListItems.Count Then
        Select Case KeyCode
            Case 40
                If sldListVer.value < sldListVer.max Then
                    sldListVer.value = tmp + ItemPage
                Else
                    sldListVer.value = sldListVer.max
                End If
                Call sldListVer_Change
            Case 39
                If sldListHor.value < sldListHor.max Then
                    sldListHor.value = sldListHor.value + 1
                    Call sldListHor_Change
                End If
            Case 38
                If sldListVer.value > sldListVer.min Then
                    sldListVer.value = tmp - 1
                Else
                    sldListVer.value = sldListVer.min
                End If
                Call sldListVer_Change
            Case 37
                If sldListHor.value > sldListHor.min Then
                    sldListHor.value = sldListHor.value - 1
                    Call sldListHor_Change
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        If txtSearch.text <> "" Then
            Dim i As Long
            Dim x As Long
            Dim tStr As String
            
            lvwFilter.ListItems.Clear
            tStr = LCase(txtSearch.text)
            
            For i = 0 To UBound(Library) - 1
                If InStr(1, LCase(Library(i).Infor.Artist), tStr) Then
                    Call ViewData(i)
                Else
                    If InStr(1, LCase(Library(i).Infor.Album), tStr) Then
                        Call ViewData(i)
                    Else
                        If InStr(1, LCase(Library(i).Infor.Title), tStr) Then
                            Call ViewData(i)
                        Else
                            If InStr(1, LCase(Library(i).Infor.Genre), tStr) Then
                                Call ViewData(i)
                            Else
                                If InStr(1, LCase(Library(i).Infor.Year), tStr) Then
                                    Call ViewData(i)
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            Call DrawList
            Call UpdateStatus
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim firstVisible As Long
    Dim tmp As Integer
    
    firstVisible = lvwFilter.GetFirstVisible.Index
    tmp = firstVisible
    If tmp > 0 And tmp <= lvwFilter.ListItems.Count Then
        Select Case KeyCode
            Case 40
                If sldListVer.value < sldListVer.max Then
                    sldListVer.value = tmp + ItemPage
                Else
                    sldListVer.value = sldListVer.max
                End If
                Call sldListVer_Change
            Case 39
                If sldListHor.value < sldListHor.max Then
                    sldListHor.value = sldListHor.value + 1
                    Call sldListHor_Change
                End If
            Case 38
                If sldListVer.value > sldListVer.min Then
                    sldListVer.value = tmp - 1
                Else
                    sldListVer.value = sldListVer.min
                End If
                Call sldListVer_Change
            Case 37
                If sldListHor.value > sldListHor.min Then
                    sldListHor.value = sldListHor.value - 1
                    Call sldListHor_Change
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim i As Long
    
    Me.Icon = LoadResPicture(112, vbResIcon)
    bolShow = True
    
    For i = 0 To 14
        LibOption.ColWidth(i) = ReadINI("library", "Column" & i, strFileconfig, 200)
        LibOption.ColView(i) = ReadINI("library", "ColumnView" & i, strFileconfig, True)
        If LibOption.ColView(i) = True Then
            lvwFilter.ColumnHeaders(i + 1).width = LibOption.ColWidth(i)
            If LibOption.ColWidth(i) < 1 Then LibOption.ColWidth(i) = 100: lvwFilter.ColumnHeaders(i + 1).width = 100
        Else
            lvwFilter.ColumnHeaders(i + 1).width = 0
        End If
        cmdColumnHeader(i).Caption = lvwFilter.ColumnHeaders(i + 1).text
    Next i
    
    For i = 1 To 14 Step 1
        imgLibrary.ListImages.Add i, , LoadResPicture(112 + i, vbResIcon)
    Next i
    
    picRate.Picture = LoadResPicture(101, vbResBitmap)

    Me.Left = ReadINI("Demension", "LibraryLeft", strFileconfig)
    Me.Top = ReadINI("Demension", "LibraryTop", strFileconfig)
    Me.height = ReadINI("Demension", "LibraryHeight", strFileconfig)
    Me.width = ReadINI("Demension", "LibraryWidth", strFileconfig)
    
    Call DrawList
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Long
    frmMenu.mnuLibrary.Checked = False
    WriteINI "Demension", "LibraryTop", Me.Top, strFileconfig
    WriteINI "Demension", "LibraryLeft", Me.Left, strFileconfig
    WriteINI "Demension", "LibraryWidth", Me.width, strFileconfig
    WriteINI "Demension", "LibraryHeight", Me.height, strFileconfig
    For i = 0 To 14
        WriteINI "Library", "Column" & i, LibOption.ColWidth(i), strFileconfig
        WriteINI "Library", "ColumnView" & i, LibOption.ColView(i), strFileconfig
    Next i
    If VideoPreview.State = playing Or Paused Then
        VideoPreview.StopVideo
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleHeight < 200 Then Me.height = 200 * Screen.TwipsPerPixelY
    If Me.ScaleWidth < 400 Then Me.width = 600 * Screen.TwipsPerPixelX
    picLibrary.Move 0, 0, (Me.ScaleWidth - 0), (Me.ScaleHeight - 25)
    picStatus.Move 0, Me.ScaleHeight - 25, (Me.ScaleWidth - 0), 25
    If txtSearch.width < 305 Then
        txtSearch.width = picToolbar.ScaleWidth - txtSearch.Left - 17
    Else
        txtSearch.width = 305
    End If
End Sub

Private Sub picColWidth_DblClick(Index As Integer)
    On Error Resume Next
    If Index <> 8 Then
        Dim i As Long
        Dim z As Long
        Dim maxWidth As Long
        
        bolAutoResize = True
        maxWidth = 1
        z = lvwFilter.GetFirstVisible.Index
        For i = z To z + ItemPage
            If i > lvwFilter.ListItems.Count Then Exit For
            If Index <> 0 Then
                If picList.TextWidth(lvwFilter.ListItems(i).SubItems(Index)) > maxWidth Then
                    maxWidth = picList.TextWidth(lvwFilter.ListItems(i).SubItems(Index))
                Else
                    maxWidth = maxWidth
                End If
            Else
                If picList.TextWidth(lvwFilter.ListItems(i).text) > maxWidth Then
                    maxWidth = picList.TextWidth(lvwFilter.ListItems(i).text)
                Else
                    maxWidth = maxWidth
                End If
            End If
        Next i
        
        lvwFilter.ColumnHeaders(Index + 1).width = maxWidth + 4
        Call DrawList
        For i = z To z + ItemPage
            If i > lvwFilter.ListItems.Count Then Exit For
            Call DrawSelected(i)
        Next i
    End If
End Sub

Private Sub picColWidth_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        bolResizeCol(Index) = True
        lngOldXCol(Index) = x
    End If
End Sub

Private Sub picColWidth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim tmp As Long
    bolAutoResize = False
    If Button = vbLeftButton Then
        If bolResizeCol(Index) Then
            Dim i As Long
            
            i = picColWidth(Index).Left
            i = i + (x - lngOldXCol(Index))
            If i < 0 Then i = 0
            picColWidth(Index).Left = i
            cmdColumnHeader(Index).width = picColWidth(Index).Left - cmdColumnHeader(Index).Left
            lvwFilter.ColumnHeaders(Index + 1).width = cmdColumnHeader(Index).width
            For i = 0 To 13
                tmp = tmp + cmdColumnHeader(i).width
            Next i
            sldListHor.max = tmp / 10
            Call DrawList
            Dim z As Long
            z = (lvwFilter.GetFirstVisible.Index)
            For i = z To z + ItemPage
                If i > lvwFilter.ListItems.Count Then Exit For
                Call DrawSelected(i)
            Next i
        End If
    End If
End Sub

Private Sub picColWidth_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        bolAutoResize = False
    End If
End Sub


Private Sub picInput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        OldInputDownX = x
        OldInputDownY = y
    End If
End Sub

Private Sub picInput_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Dim z As Long
        Dim s As Long
        z = picInput.Left
        s = picInput.Top
        picInput.Left = z + (x - OldInputDownX)
        picInput.Top = s + (y - OldInputDownY)
    End If
End Sub

Private Sub picLibrary_Resize()
    On Error Resume Next
    picToolbar.width = picLibrary.ScaleWidth
    tvwLibrary.height = picLibrary.ScaleHeight - 57
    picWidth.Move tvwLibrary.width, 40, 2, picLibrary.ScaleHeight - 40
    picMaster.Move tvwLibrary.width + 2, 40, picLibrary.ScaleWidth - (tvwLibrary.width + 1), picLibrary.ScaleHeight - 40
    sldListVer.Left = picLibrary.ScaleWidth - sldListVer.width
    sldListHor.Top = picLibrary.ScaleHeight - sldListHor.height
    sldListHor.Left = picMaster.Left
    sldListHor.width = picLibrary.ScaleWidth - (picMaster.Left - sldListVer.width) - 17
    Call DrawList
    Dim i As Long, z As Long
    For i = z To z + ItemPage
        If i > lvwFilter.ListItems.Count Then Exit For
        Call DrawSelected(i)
    Next i
End Sub
Private Sub SetupTree()
    On Error Resume Next
    Dim nodRoot As Node
    Dim nodSub As Node
    Dim nodSubChild As Node
    
    tvwLibrary.ImageList = imgLibrary
    
    ' All Audio
    Set nodRoot = tvwLibrary.Nodes.Add(, , , , 1, 1)
    nodRoot.Key = "Audio Root"
    nodRoot.text = "All Audio"
        'Audio Child
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 2, 2)
        nodSub.Key = "Audio Artist"
        nodSub.text = "Artist"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 3, 3)
        nodSub.Key = "Audio Album"
        nodSub.text = "Album"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 4, 4)
        nodSub.Key = "Audio Genre"
        nodSub.text = "Genre"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 5, 5)
        nodSub.Key = "Audio Year"
        nodSub.text = "Year"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 6, 6)
        nodSub.Key = "Audio Rated"
        nodSub.text = "Rated Song"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 6, 6)
            nodSubChild.Key = "Audio Rated 1"
            nodSubChild.text = "5 Star Rated Song"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 6, 6)
            nodSubChild.Key = "Audio Rated 2"
            nodSubChild.text = "4 Star Rated Song"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 6, 6)
            nodSubChild.Key = "Audio Rated 3"
            nodSubChild.text = "3 Star Rated Song"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 6, 6)
            nodSubChild.Key = "Audio Rated 4"
            nodSubChild.text = "2 Star Rated Song"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 6, 6)
            nodSubChild.Key = "Audio Rated 5"
            nodSubChild.text = "1 Star Rated Song"
        
    'All Video
    Set nodRoot = tvwLibrary.Nodes.Add(, , , , 7, 7)
    nodRoot.Key = "Video Root"
    nodRoot.text = "All Video"
        'Video Child
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 8, 8)
        nodSub.Key = "Video Artist"
        nodSub.text = "Artist"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 14, 14)
        nodSub.Key = "Video Genre"
        nodSub.text = "Genre"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 9, 9)
        nodSub.Key = "Video Year"
        nodSub.text = "Year"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 10, 10)
        nodSub.Key = "Video Rated"
        nodSub.text = "Rated Video"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 10, 10)
            nodSubChild.Key = "Video Rated 1"
            nodSubChild.text = "5 Star Rated Video"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 10, 10)
            nodSubChild.Key = "Video Rated 2"
            nodSubChild.text = "4 Star Rated Video"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 10, 10)
            nodSubChild.Key = "Video Rated 3"
            nodSubChild.text = "3 Star Rated Video"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 10, 10)
            nodSubChild.Key = "Video Rated 4"
            nodSubChild.text = "2 Star Rated Video"
            Set nodSubChild = tvwLibrary.Nodes.Add(nodSub, tvwChild, , , 10, 10)
            nodSubChild.Key = "Video Rated 5"
            nodSubChild.text = "1 Star Rated Video"
            
    'Playlist
    Set nodRoot = tvwLibrary.Nodes.Add(, , , , 11, 11)
    nodRoot.Key = "My Playlist"
    nodRoot.text = "My Playlist"
    
    'Auto playlist
    Set nodRoot = tvwLibrary.Nodes.Add(, , , , 12, 12)
    nodRoot.Key = "Auto Playlist"
    nodRoot.text = "Auto Playlist"
        'Auto playlist child
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 12, 12)
        nodSub.Key = "Auto Playlist 1"
        nodSub.text = "Favorite - 4 and 5 star rated"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 12, 12)
        nodSub.Key = "Auto Playlist 2"
        nodSub.text = "Favorite - Top 10 most played"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 12, 12)
        nodSub.Key = "Auto Playlist 3"
        nodSub.text = "Fresh - yet played"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 12, 12)
        nodSub.Key = "Auto Playlist 4"
        nodSub.text = "Fresh - yet rated"
        Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 12, 12)
        nodSub.Key = "Auto Playlist 5"
        nodSub.text = "Fresh - new added"

    'Now Playing
    Set nodRoot = tvwLibrary.Nodes.Add(, , , , 13, 13)
    nodRoot.Key = "Now Playing"
    nodRoot.text = "Now Playing"
End Sub
Private Sub AddNodeLibrary()
    On Error Resume Next
    Dim nodRoot As Node
    Dim i As Long
    Dim y As Long
    
    tvwLibrary.Nodes.Clear
    Call SetupTree
    
    
    ' All Audio
        'Artist
        For i = 0 To UBound(Artist) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(1).Child, tvwChild, , , 2, 2)
            nodRoot.Key = "Audio Artist " & i
            nodRoot.text = Artist(i)
        Next i
        tvwLibrary.Nodes(1).Child.Sorted = True
        
        'Album
        For i = 0 To UBound(Album) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(1).Child.Next, tvwChild, , , 3, 3)
            nodRoot.Key = "Audio Album " & i
            nodRoot.text = Album(i)
        Next i
        tvwLibrary.Nodes(1).Child.Next.Sorted = True
        
        'Genre
        For i = 0 To UBound(Genre) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(1).Child.Next.Next, tvwChild, , , 4, 4)
            nodRoot.Key = "Audio Genre " & i
            nodRoot.text = Genre(i)
        Next i
        tvwLibrary.Nodes(1).Child.Next.Next.Sorted = True
        
        'Year
        For i = 0 To UBound(Year) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(1).Child.Next.Next.Next, tvwChild, , , 5, 5)
            nodRoot.Key = "Audio Year " & i
            nodRoot.text = Year(i)
        Next i
        tvwLibrary.Nodes(1).Child.Next.Next.Next.Sorted = True
    'End All Audio
    
    'All Video
    y = FindNode("All Video")
        'Artist
        For i = 0 To UBound(ArtistV) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(y).Child, tvwChild, , , 8, 8)
            nodRoot.Key = "Video Artist " & i
            nodRoot.text = ArtistV(i)
        Next i
        tvwLibrary.Nodes(y).Child.Sorted = True
        
        'Genre
        For i = 0 To UBound(GenreV) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(y).Child.Next, tvwChild, , , 4, 4)
            nodRoot.Key = "Video Genre " & i
            nodRoot.text = GenreV(i)
        Next i
        tvwLibrary.Nodes(y).Child.Next.Sorted = True
        
        'Year
        For i = 0 To UBound(YearV) - 1
            Set nodRoot = tvwLibrary.Nodes.Add(tvwLibrary.Nodes(y).Child.Next.Next, tvwChild, , , 9, 9)
            nodRoot.Key = "Video Year " & i
            nodRoot.text = YearV(i)
        Next i
        tvwLibrary.Nodes(y).Child.Next.Next.Sorted = True
    'End All Video
    
    'My Playlist
    Dim nodSub As Node
    y = FindNode("My Playlist")
        For i = 0 To UBound(Playlist) - 1
            If Trim(Playlist(i).Name) > 0 Then
                Set nodRoot = tvwLibrary.Nodes(y)
                Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 11, 11)
                nodSub.Key = "My Playlist " & i
                nodSub.text = Playlist(i).Name
                tvwLibrary.Nodes(y).Child.Sorted = True
            End If
        Next i
        tvwLibrary.Nodes(y).Child.Sorted = True
End Sub

Private Sub picList_DblClick()
    On Error Resume Next
    Dim tStr As String
    tStr = Library(CurrentTrack).Infor.FullName
    If bolPlayingNow = False Then
        Select Case LibOption.intDblClick
            Case 0
                Call AddFile(tStr)
                frmPlayList.List.DisplayList
            Case 1
                Call AddFile(tStr)
                frmPlayList.List.DisplayList
                Dim z As Long
                For z = 1 To frmPlayList.List.ListItemCount
                    If tStr = NowPlaying(frmPlayList.List.Key(z)).Infor.FullName Then
                        Play (z)
                        Exit For
                    End If
                Next z
            Case 2
                Select Case LCase(Library(CurrentTrack).strType)
                    Case "audio"
                        VideoPreview.OpenVideo tStr
                        VideoPreview.PlayVideo
                    Case "video"
                        VideoPreview.OpenVideo tStr
                        VideoPreview.Visible = True
                        VideoPreview.PlayVideo
                        VideoPreview.Display = CustomizeSize
                End Select
        End Select
    Else
        Dim i As Long
        For i = 1 To frmPlayList.List.ListItemCount
            If lvwFilter.ListItems(CurrentSelect).SubItems(13) = NowPlaying(frmPlayList.List.Key(i)).Infor.FullName Then
                Call Play(CInt(i))
                Exit For
            End If
        Next i
    End If
End Sub


Private Sub picList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    If lvwFilter.ListItems.Count = 0 Then
        Exit Sub
    Else
        Dim firstVisible As Long
        Dim tmp As Integer
        firstVisible = lvwFilter.GetFirstVisible.Index
        tmp = (firstVisible + (y - lvwFilter.ListItems(firstVisible).Top) \ lvwFilter.ListItems(1).height)
        If tmp > 0 And tmp <= lvwFilter.ListItems.Count Then
            If Button = vbLeftButton Then
                Select Case Shift
                    Case 0 'none
                        For i = 1 To lvwFilter.ListItems.Count
                            lvwFilter.ListItems(i).Selected = False
                        Next i
                        CurrentSelect = 0
                        intStart = 0
                        intEnd = 0
                    Case 1 'shift
                    Case 2 'Ctrl
                        lvwFilter.ListItems(tmp).Selected = Not lvwFilter.ListItems(tmp).Selected
                        intStart = 0
                        intEnd = 0
                End Select
            End If
        End If
    End If
End Sub

Private Sub picList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim i As Long
    If lvwFilter.ListItems.Count = 0 Then
        Exit Sub
    Else
        If Button = vbRightButton Then
            Dim z As Long
            For z = 0 To 5
                frmMenu.mnuRateC(z).Checked = False
            Next z
            frmMenu.mnuRateC(5 - Library(CurrentTrack).intRate).Checked = True
            PopupMenu frmMenu.mnuLibEdit, vbPopupMenuLeftAlign
            For z = 0 To UBound(Playlist) - 1
                If Playlist(x).Name <> "" Then
                    Call AddNewPlaylist(z)
                End If
            Next z
        End If
        If Button = vbLeftButton Then
            Dim firstVisible As Long
            Dim tmp As Integer
            firstVisible = lvwFilter.GetFirstVisible.Index
        tmp = (firstVisible + (y - lvwFilter.ListItems(firstVisible).Top) \ lvwFilter.ListItems(1).height)
            Select Case Shift
                Case 0
                    intStart = 0
                    intEnd = 0
                    CurrentSelect = tmp
                    lvwFilter.ListItems(CurrentSelect).Selected = True
                Case 1 'shift
                    intEnd = tmp
                    intStart = CurrentSelect
                    If intStart < intEnd Then
                        For i = intStart To intEnd Step 1
                                lvwFilter.ListItems(i).Selected = True
                        Next i
                    Else
                        For i = intStart To intEnd Step -1
                                lvwFilter.ListItems(i).Selected = True
                        Next i
                    End If
                Case 2
                    intStart = 0
                    intEnd = 0
                    If lvwFilter.ListItems(tmp).Selected = True Then
                        CurrentSelect = tmp
                    Else
                        CurrentSelect = -1
                    End If
                End Select
        End If
        
        Call DrawList
        z = (lvwFilter.GetFirstVisible.Index)
        For i = z To z + ItemPage
            If i > lvwFilter.ListItems.Count Then Exit For
            Call DrawSelected(i)
        Next i
        If CurrentSelect <> 0 Then
            sldListVer.value = CurrentSelect
        End If
        picList.Refresh
        If CurrentSelect <> -1 Then
            CurrentTrack = TrackIndex(lvwFilter.ListItems(CurrentSelect).SubItems(13))
        End If
    End If
End Sub


Private Sub picList_Resize()
    Dim i As Long
    
    For i = 0 To 14
        picColWidth(i).height = picList.height
    Next i
    picList.BackColor = &HFFFFFF
    Call DrawGird
    Set picList.Picture = picList.Image

End Sub

Private Sub picMaster_Resize()
    On Error Resume Next
    Dim i As Long
    Dim tmp As Long
    
    For i = 0 To 12
        tmp = tmp + cmdColumnHeader(i).width
    Next i
    
    lvwFilter.Move 0, 0, picMaster.ScaleWidth, picMaster.ScaleHeight
    picList.Move picList.Left, 0, (picMaster.ScaleWidth + tmp), picMaster.ScaleHeight
    VideoPreview.Move tvwLibrary.width + 1, 40, picLibrary.ScaleWidth - VideoPreview.Left, picMaster.ScaleHeight
    sldListVer.height = picMaster.ScaleHeight - 16
    sldListHor.max = tmp / 100

End Sub
Private Sub picToolbar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMenu.mnuLibOpt, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub picWidth_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        bolResize = True
        lngOldX = x
    End If
End Sub

Private Sub picWidth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If bolResize Then
            Dim i As Long
            
            i = picWidth.Left
            i = i + (x - lngOldX)
            If i < 178 Then i = 178
            If i > 400 Then i = 400
            picWidth.Left = i
            
            cmdTree.Move 0, 40, picWidth.Left, 17
            tvwLibrary.Move 0, 57, picWidth.Left, picLibrary.ScaleHeight - 57
            picMaster.Move tvwLibrary.width + 1, 40
            sldListHor.Left = picMaster.Left
            sldListHor.width = picLibrary.ScaleWidth - (picMaster.Left - sldListVer.width)
            VideoPreview.Move tvwLibrary.width + 1, 40, picLibrary.ScaleWidth - VideoPreview.Left, picMaster.ScaleHeight
        End If
    End If
End Sub

Private Sub sldListHor_Change()
    On Error Resume Next
    picList.Left = -(sldListHor.value * 100)
End Sub


Private Sub sldListVer_Change()
    On Error GoTo beep
    Dim z As Long
    Dim i As Long
    
    If lvwFilter.ListItems.Count > 1 Then
        If sldListVer.value <= lvwFilter.ListItems.Count Then
            lvwFilter.ListItems.Item(sldListVer.value).EnsureVisible
        End If
    End If
    
    Call DrawList
    z = (lvwFilter.GetFirstVisible.Index)
    For i = z To z + ItemPage
        If i > lvwFilter.ListItems.Count Then Exit For
            Call DrawSelected(i)
    Next i
    Exit Sub
beep:
    If sldListVer.value = 0 Then sldListVer.value = 1
    If sldListVer.value > lvwFilter.ListItems.Count Then sldListVer.value = lvwFilter.ListItems.Count

End Sub


Private Sub tvwLibrary_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim y As Long
    Dim z As Long
    Dim x As Long
    Dim tFilter As String
    Dim tStr As String
    
    tStr = Mid(CurrentNode.Key, 1, InStrRev(CurrentNode.Key, " ", -1, vbBinaryCompare) - 1)
    tFilter = Mid(tStr, InStrRev(tStr, " ", -1, vbBinaryCompare) + 1)
    
    y = lvwFilter.ListItems.Count
    
    Select Case LCase(tFilter)
        Case "artist"
            For z = 1 To y
                x = TrackIndex(lvwFilter.ListItems(z).SubItems(13))
                Library(x).Infor.Artist = NewString
                Library(x).strDayUpdate = date
            Next z
        Case "album"
            For z = 1 To y
                x = TrackIndex(lvwFilter.ListItems(z).SubItems(13))
                Library(x).Infor.Album = NewString
                Library(x).strDayUpdate = date
            Next z
        Case "genre"
            For z = 1 To y
                x = TrackIndex(lvwFilter.ListItems(z).SubItems(13))
                Library(x).Infor.Genre = NewString
                Library(x).strDayUpdate = date
            Next z
        Case "year"
            For z = 1 To y
                x = TrackIndex(lvwFilter.ListItems(z).SubItems(13))
                Library(x).Infor.Year = NewString
                Library(x).strDayUpdate = date
            Next z
        Case "playlist"
            Playlist(CurrentPlaylist).Name = NewString
            Call WriteDataList(CurrentPlaylist)
    End Select
    lvwFilter.ListItems.Clear
    Call ReadLibrary
    Call DrawList
End Sub

Private Sub tvwLibrary_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        bolTreeRDown = True
    End If
    If Button = vbLeftButton Then
        bolTreeRDown = False
    End If
End Sub

Private Sub tvwLibrary_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim tmp As NodeLibrary
    Dim i As Long
       
    bolPlayingNow = False
    For i = 1 To lvwFilter.ListItems.Count
        lvwFilter.ListItems(i).Selected = False
    Next i
    Set CurrentNode = Node
    
    Select Case LCase(Mid(Node.Key, 1, InStr(1, Node.Key, " ", vbBinaryCompare) - 1))
        Case "audio"
            tmp.Type = 0
        Case "video"
            tmp.Type = 1
        Case "my", "auto", "now"
            tmp.Type = 2
    End Select
    With Node
        tmp.Name = .text
        tmp.Refer = .Parent.text
    End With
    
    cmdTree.Caption = IIf(tmp.Refer <> "", tmp.Refer & " -- " & tmp.Name, tmp.Name)
    
    If tmp.Type = 0 Or tmp.Type = 1 Then 'If type is audio or video get ? node is child or parent or root
        Select Case LCase(tmp.Refer)
            Case ""
                tmp.Refer = "NodeRoot"
            Case LCase(Node.Root.text), LCase(Node.Root.Next.text)
                tmp.Refer = "NodeParent"
        End Select
    End If
    
    
    Select Case tmp.Type
        Case 0, 1
            Call Filter(tmp.Type, tmp.Refer, tmp.Name)
        Case 2
            Call FilterPlaylist(tmp.Refer, tmp.Name)
            If tmp.Refer = "My Playlist" Then
                Dim s As Long
                Dim x As Long
                s = -1
                For x = 0 To UBound(Playlist) - 1
                    If Len(tmp.Name) = Len(Playlist(x).Name) Then
                        If LCase(tmp.Name) = LCase(Playlist(x).Name) Then
                            s = x
                            Exit For
                        End If
                    End If
                    CurrentPlaylist = s
                Next x
            End If
    End Select
    
    Call DrawList
    
    If bolTreeRDown Then
        If tmp.Type = 0 Or tmp.Type = 1 Then
            If tmp.Refer = "NodeRoot" Or tmp.Refer = "NodeParent" Then
                frmMenu.mnuTreeViewC(2).Visible = False
                frmMenu.mnuTreeViewC(3).Visible = False
                frmMenu.mnuTreeViewC(4).Visible = False
            Else
                frmMenu.mnuTreeViewC(2).Visible = True
                frmMenu.mnuTreeViewC(3).Visible = True
                frmMenu.mnuTreeViewC(4).Visible = True
            End If
        Else
            If tmp.Refer = "" Then
                frmMenu.mnuTreeViewC(2).Visible = False
                frmMenu.mnuTreeViewC(3).Visible = False
                frmMenu.mnuTreeViewC(4).Visible = False
            Else
                If tmp.Refer = "My Playlist" Then
                    frmMenu.mnuTreeViewC(2).Visible = True
                    frmMenu.mnuTreeViewC(3).Visible = True
                    frmMenu.mnuTreeViewC(4).Visible = True
                End If
            End If
        End If
        bolTreeRDown = False
        PopupMenu frmMenu.mnuTreeView, vbPopupMenuLeftAlign
    End If
End Sub
Private Sub Filter(intType As Integer, strRefer As String, strFilter As String)
    On Error Resume Next
    Dim tStr As String
    Dim i As Long
    Dim tmpSort As Long
    
    tmpSort = lvwFilter.SortKey
    
    lvwFilter.ListItems.Clear
    lvwFilter.Sorted = False
    
    If intType = 0 Then
        tStr = "Audio"
    End If
    If intType = 1 Then
        tStr = "Video"
    End If
    
    Select Case strRefer
        Case "NodeRoot", "NodeParent"
            For i = 0 To UBound(Library) - 1
                If Library(i).strType = tStr Then
                    Select Case strFilter
                        Case "Rated Song", "Rated Video"
                            If Library(i).intRate > 0 Then
                                Call ViewData(i)
                            End If
                        Case Else
                            Call ViewData(i)
                    End Select
                End If
            Next i
        Case "Artist"
            For i = 0 To UBound(Library) - 1
                If Library(i).strType = tStr Then
                    If Library(i).Infor.Artist = strFilter Then
                        Call ViewData(i)
                    End If
                End If
            Next i
        Case "Album"
            For i = 0 To UBound(Library) - 1
                If Library(i).strType = tStr Then
                    If Library(i).Infor.Album = strFilter Then
                        Call ViewData(i)
                    End If
                End If
            Next i
        Case "Genre"
            For i = 0 To UBound(Library) - 1
                If Library(i).strType = tStr Then
                    If Library(i).Infor.Genre = strFilter Then
                        Call ViewData(i)
                    End If
                End If
            Next i
        Case "Year"
            For i = 0 To UBound(Library) - 1
                If Library(i).strType = tStr Then
                    If Library(i).Infor.Year = strFilter Then
                        Call ViewData(i)
                    End If
                End If
            Next i
        Case "Rated Song", "Rated Video"
            For i = 0 To UBound(Library) - 1
                If Library(i).strType = tStr Then
                    If Library(i).intRate = CInt(Mid(strFilter, 1, 1)) Then
                        Call ViewData(i)
                    End If
                End If
            Next i
    End Select
    
    lvwFilter.Sorted = True
    lvwFilter.SortKey = tmpSort
    
    Call UpdateStatus
End Sub

Private Sub FilterPlaylist(strRefer As String, strFilter As String)
    Dim i As Long
    Dim x As Long
    Dim lvw As ListItem
    
    lvwFilter.ListItems.Clear
    Select Case strRefer
        Case ""
            Select Case strFilter
                Case "Now Playing"
                    Dim tmpSort As Long
                    Dim tmpOrder As Long
    
                    tmpSort = lvwFilter.SortKey
                    tmpOrder = lvwFilter.SortOrder
                    lvwFilter.Sorted = False

                    If frmPlayList.List.ListItemCount > 0 Then
                        For i = 1 To frmPlayList.List.ListItemCount
                            Dim tTrack As track
                            Dim z As Long
                            tTrack = NowPlaying(frmPlayList.List.Key(i)).Infor
                            If TrackIndex(NowPlaying(frmPlayList.List.Key(i)).Infor.FullName) <> -1 Then
                                ViewData (TrackIndex(NowPlaying(frmPlayList.List.Key(i)).Infor.FullName))
                            Else
                                z = lvwFilter.ListItems.Count + 1
                                lvwFilter.ListItems.Add z, , tTrack.Artist
                                z = lvwFilter.ListItems.Count
                                lvwFilter.ListItems(z).SubItems(1) = tTrack.Album
                                lvwFilter.ListItems(z).SubItems(2) = tTrack.Title
                                lvwFilter.ListItems(z).SubItems(3) = tTrack.Genre
                                lvwFilter.ListItems(z).SubItems(4) = tTrack.Year
                                lvwFilter.ListItems(z).SubItems(5) = tTrack.bitrate
                                lvwFilter.ListItems(z).SubItems(6) = Time2String(tTrack.Duration)
                                lvwFilter.ListItems(z).SubItems(7) = tTrack.Frequency
                                lvwFilter.ListItems(z).SubItems(8) = ""
                                lvwFilter.ListItems(z).SubItems(9) = date
                                lvwFilter.ListItems(z).SubItems(10) = date
                                lvwFilter.ListItems(z).SubItems(11) = 0
                                lvwFilter.ListItems(z).SubItems(12) = tTrack.Filename
                                lvwFilter.ListItems(z).SubItems(13) = tTrack.FullName
                                lvwFilter.ListItems(z).SubItems(14) = tTrack.Size
                            End If
                        Next i
                        
                        lvwFilter.SortKey = tmpSort
                        lvwFilter.SortOrder = tmpOrder
                        lvwFilter.Sorted = True
                        bolPlayingNow = True
                    End If
                    lblStatus.Caption = "Now Playing have :" & lvwFilter.ListItems.Count & " Item(s)"
                Case "My Playlist"
                    lblStatus.Caption = "My Playlist have " & UBound(Playlist) & " Item(s)"
                Case "Auto Playlist"
                    lblStatus.Caption = "Please select one child sub item"
            End Select
        Case "Auto Playlist"
            Select Case strFilter
                Case "Favorite - 4 and 5 star rated"
                    Call LoadAutoPlaylist(0)
                Case "Favorite - Top 10 most played"
                    Call LoadAutoPlaylist(1)
                Case "Fresh - yet played"
                    Call LoadAutoPlaylist(2)
                Case "Fresh - yet rated"
                    Call LoadAutoPlaylist(3)
                Case "Fresh - new added"
                    Call LoadAutoPlaylist(4)
            End Select
        Case "My Playlist"
            Call LoadMyPlaylist(strFilter)
    End Select
End Sub
Private Sub ViewData(Index As Long)
    Me.MousePointer = vbHourglass
    
    DoEvents
    Dim i As Long
    i = lvwFilter.ListItems.Count + 1
    lvwFilter.ListItems.Add i, , Library(Index).Infor.Artist
    i = lvwFilter.ListItems.Count
    lvwFilter.ListItems(i).SubItems(1) = Library(Index).Infor.Album
    lvwFilter.ListItems(i).SubItems(2) = Library(Index).Infor.Title
    lvwFilter.ListItems(i).SubItems(3) = Library(Index).Infor.Genre
    lvwFilter.ListItems(i).SubItems(4) = Library(Index).Infor.Year
    lvwFilter.ListItems(i).SubItems(5) = Library(Index).Infor.bitrate
    lvwFilter.ListItems(i).SubItems(6) = Time2String(Library(Index).Infor.Duration)
    lvwFilter.ListItems(i).SubItems(7) = Library(Index).Infor.Frequency
    lvwFilter.ListItems(i).SubItems(8) = Library(Index).intRate
    lvwFilter.ListItems(i).SubItems(9) = Library(Index).strDay
    lvwFilter.ListItems(i).SubItems(10) = Library(Index).strDayUpdate
    lvwFilter.ListItems(i).SubItems(11) = Library(Index).intPlaycount
    lvwFilter.ListItems(i).SubItems(12) = Library(Index).Infor.Filename
    lvwFilter.ListItems(i).SubItems(13) = Library(Index).Infor.FullName
    lvwFilter.ListItems(i).SubItems(14) = Library(Index).Infor.Size
    
    sldListVer.max = lvwFilter.ListItems.Count
    sldListVer.min = 1
    Me.MousePointer = vbNormal
End Sub

Private Sub LoadMyPlaylist(strPlaylist As String)
    Dim i As Long
    Dim x As Long
    
    For i = 0 To UBound(Playlist) - 1
        If Playlist(i).Name = strPlaylist Then
            x = i
            Exit For
        End If
    Next i
    
    arrPL = Split(Playlist(x).file & ",", ",")
    For i = 0 To UBound(arrPL) - 1
        Call ViewData(CLng(arrPL(i)))
    Next i
    Call DrawList
    Call UpdateStatus
End Sub
Private Sub LoadAutoPlaylist(AutoPlaylist As Integer)
    Dim tStr As String
    Dim i As Long
    
    Dim tmpSort As Long
    Dim tmpOrder As Long
    
    tmpSort = lvwFilter.SortKey
    tmpOrder = lvwFilter.SortOrder
    lvwFilter.Sorted = False
    
    Select Case AutoPlaylist
        Case 0
            For i = 0 To UBound(Library) - 1
                If Library(i).intRate = 4 Or Library(i).intRate = 5 Then
                    Call ViewData(i)
                End If
            Next i
        Case 1
            If UBound(Library) <= 10 Then
                For i = 0 To UBound(Library) - 1
                    ViewData (i)
                Next i
                Exit Sub
            End If
            Dim arrTop(9) As Long
            Dim x As Long
            For i = 0 To UBound(Library) - 1
                For x = 0 To Library(i).intPlaycount
                    tStr = tStr & "a"
                Next x
                lvwFilter.ListItems.Add i + 1, , tStr
                lvwFilter.ListItems(i + 1).SubItems(1) = i
                tStr = ""
            Next i
            lvwFilter.Sorted = True
            lvwFilter.SortKey = 0
            lvwFilter.SortOrder = lvwDescending
            For i = 0 To 9
                arrTop(i) = (CLng(lvwFilter.ListItems(i + 1).SubItems(1)))
            Next i
            lvwFilter.ListItems.Clear
            lvwFilter.Sorted = False
            For i = 0 To 9
                ViewData (arrTop(i))
            Next i
        Case 2
            For i = 0 To UBound(Library) - 1
                If Library(i).intPlaycount = 0 Then
                    Call ViewData(i)
                End If
            Next i
        Case 3
            For i = 0 To UBound(Library) - 1
                If Library(i).intRate = 0 Then
                    Call ViewData(i)
                End If
            Next i
        Case 4
            For i = 0 To UBound(Library) - 1
                If CDate(Library(i).strDay) = date Then
                    Call ViewData(i)
                End If
            Next i
    End Select
    lvwFilter.SortKey = tmpSort
    lvwFilter.SortOrder = tmpOrder
    lvwFilter.Sorted = True
    Call UpdateStatus
End Sub

Public Function ItemPage() As Integer
    On Error Resume Next
    If lvwFilter.ListItems.Count = 0 Then Exit Function
    ItemPage = lvwFilter.height / lvwFilter.ListItems(1).height - 1
End Function
Private Sub DrawGird()
    On Error Resume Next
    Dim rec As RECT
    Dim lngColor As Long
    Dim x As Long
    Dim y As Long
    
    lngColor = &HE35400
    y = picList.TextHeight("I") + 1
    For x = y + 1 To picList.ScaleHeight Step y
        picList.Line (0, x + 1)-(picList.ScaleWidth, x + 1), lngColor
    Next x
End Sub
Private Sub DrawRate(Index As Long)
    On Error Resume Next
    Dim i As Integer
    Dim intRate As Integer
    
    intRate = CInt(lvwFilter.ListItems(Index).SubItems(8))
    If lvwFilter.ListItems(Index).Selected = True Then
        picList.Line (cmdColumnHeader(8).Left, lvwFilter.ListItems(Index).Top - 1)-(cmdColumnHeader(8).Left + cmdColumnHeader(8).width, lvwFilter.ListItems(Index).Top + lvwFilter.ListItems(Index).height - 1), &HE35400, BF
    End If
    If intRate = 0 Then Exit Sub
    For i = 1 To intRate
        TransparentBlt picList.hDC, cmdColumnHeader(8).Left + 13 * (i - 1), lvwFilter.ListItems(Index).Top, 12, 12, picRate.hDC, 0, 0, 48, 48, &HFF00FF
    Next i
End Sub
Public Sub DrawList() 'Support MultiSelect
    On Error Resume Next
    Dim rec As RECT
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim tStr As String
    
    
    If VideoPreview.State = Stopped Then
        VideoPreview.Visible = False
    End If
    
    picList.Cls
    picList.BackColor = &HFFFFFF
    picList.ForeColor = &HE35400
    'Draw ColumnHeader
    
    For i = 0 To 14
        If lvwFilter.ColumnHeaders(i + 1).width = 0 Then
            LibOption.ColView(i) = False
            cmdColumnHeader(i).Visible = False
            picColWidth(i).Visible = False
            cmdColumnHeader(i).Left = lvwFilter.ColumnHeaders(i + 1).Left
            cmdColumnHeader(i).width = 1
        Else
            LibOption.ColView(i) = True
            cmdColumnHeader(i).Left = lvwFilter.ColumnHeaders(i + 1).Left
            cmdColumnHeader(i).width = lvwFilter.ColumnHeaders(i + 1).width
            cmdColumnHeader(i).Visible = True
            picColWidth(i).Visible = True
            LibOption.ColWidth(i) = cmdColumnHeader(i).width
        End If
        frmMenu.mnuViewC(i).Checked = LibOption.ColView(i)
    Next i
    cmdColumnHeader(15).Left = cmdColumnHeader(14).Left + cmdColumnHeader(14).width + 1
    cmdColumnHeader(15).width = picList.width - (cmdColumnHeader(14).Left + cmdColumnHeader(14).width + 1)
    
    For i = 0 To 14
        picColWidth(i).Top = 0
        picColWidth(i).Left = cmdColumnHeader(i).Left + cmdColumnHeader(i).width
    Next i
            
    If lvwFilter.ListItems.Count <= ItemPage Then
        sldListVer.Visible = False
    Else
        sldListVer.Visible = True
    End If
    
    Dim z As Long
    z = 0
    For i = 0 To cmdColumnHeader.Count - 1
        z = z + cmdColumnHeader(i).width
    Next i
    If z <= sldListHor.width Then
        sldListHor.Enabled = False
    Else
        sldListHor.Enabled = True
    End If
    
    'Draw List
    Dim DT As Long
    
    If lvwFilter.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    y = (lvwFilter.GetFirstVisible.Index)
    For x = y To y + ItemPage
        If x > lvwFilter.ListItems.Count Then Exit For
        tStr = lvwFilter.ListItems(x).text
        rec.Top = lvwFilter.ListItems(x).Top
        rec.Bottom = rec.Top + lvwFilter.ListItems(x).height
        rec.Left = 0
        rec.Right = lvwFilter.ColumnHeaders(1).width
        Call DrawText(picList.hDC, tStr, Len(tStr), rec, DT_LEFT)
        For i = 2 To lvwFilter.ColumnHeaders.Count + 1
            If lvwFilter.ColumnHeaders(i).width <> 0 Then
                If i <> 9 Then
                    tStr = lvwFilter.ListItems(x).SubItems(i - 1)
                    rec.Left = 2 + lvwFilter.ColumnHeaders(i - 1).Left + lvwFilter.ColumnHeaders(i - 1).width
                    rec.Right = rec.Left + lvwFilter.ColumnHeaders(i).width - 4
                    Select Case lvwFilter.ColumnHeaders(i).Alignment
                        Case 0
                            DT = DT_LEFT
                        Case 1
                            DT = DT_RIGHT
                        Case 2
                            DT = DT_CENTER
                    End Select
                    Call DrawText(picList.hDC, tStr, Len(tStr), rec, DT)
                Else
                    Call DrawRate(x)
                End If
            End If
        Next i
        If FileExists(lvwFilter.ListItems(x).SubItems(13)) = False Then
            picList.Line (0, rec.Top + lvwFilter.ListItems(x).height / 2)-(lvwFilter.width, rec.Top + lvwFilter.ListItems(x).height / 2), &HFF
        End If
    Next x
End Sub
Public Sub DrawSelected(Index As Long)
    On Error Resume Next
    
    Dim rec As RECT
    Dim lngColor As Long
    Dim i As Long
    Dim DT As Long
    Dim tStr As String
    
    lngColor = &HE35400
    picList.ForeColor = &HFFFFFF
    
    'Draw
    If lvwFilter.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lvwFilter.ListItems(Index).Selected = False Then
        Exit Sub
    End If
    
    tStr = lvwFilter.ListItems(Index).text
    rec.Top = lvwFilter.ListItems(Index).Top
    rec.Left = 0
    rec.Bottom = rec.Top + lvwFilter.ListItems(Index).height
    rec.Right = lvwFilter.ColumnHeaders(1).width
    picList.Line (rec.Left, rec.Top)-(rec.Right, rec.Bottom - 2), lngColor, BF
    Call DrawText(picList.hDC, tStr, Len(tStr), rec, DT_LEFT)
    For i = 2 To lvwFilter.ColumnHeaders.Count + 1
        If i <> 9 Then
            Select Case lvwFilter.ColumnHeaders(i).Alignment
                Case 0
                    DT = DT_LEFT
                Case 1
                    DT = DT_RIGHT
                Case 2
                    DT = DT_CENTER
            End Select
            tStr = lvwFilter.ListItems(Index).SubItems(i - 1)
            rec.Top = lvwFilter.ListItems(Index).Top
            rec.Left = lvwFilter.ColumnHeaders(i - 1).Left + lvwFilter.ColumnHeaders(i - 1).width
            rec.Bottom = rec.Top + lvwFilter.ListItems(Index).height
            rec.Right = rec.Left + lvwFilter.ColumnHeaders(i).width
            picList.Line (rec.Left, rec.Top)-(rec.Right, rec.Bottom - 2), lngColor, BF
            Call DrawText(picList.hDC, tStr, Len(tStr), rec, DT)
        Else
            Call DrawRate(Index)
        End If
    Next i
End Sub

Private Sub CreateNewPlaylist(strName As String)
    On Error Resume Next
    Dim i As Long
    Dim x As Long
    Dim tStr As String
    
    If Trim(strName) = "" Then Exit Sub
    
    For i = 0 To UBound(Playlist) - 1
        If strName = Playlist(i).Name Then
            MsgBox strName & " is already exists", vbCritical, "M3P - Error"
            Exit Sub
        End If
    Next i
    
    ReDim Preserve Playlist(UBound(Playlist) + 1)
    Playlist(UBound(Playlist) - 1).Name = strName
    
    For x = 1 To lvwFilter.ListItems.Count
        If lvwFilter.ListItems(x).Selected = True Then
            i = TrackIndex(lvwFilter.ListItems(x).SubItems(13))
            Playlist(UBound(Playlist) - 1).file = Playlist(UBound(Playlist) - 1).file & "," & i
        End If
    Next x
    
    Call WriteDataList(UBound(Playlist) - 1)
    'My Playlist
    Call AddNewPlaylist(UBound(Playlist) - 1)
    
    Dim nodRoot As Node
    Dim nodSub As Node
    x = FindNode("My Playlist")
    Set nodRoot = tvwLibrary.Nodes(x)
    Set nodSub = tvwLibrary.Nodes.Add(nodRoot, tvwChild, , , 11, 11)
    nodSub.Key = "My Playlist " & UBound(Playlist) - 1
    nodSub.text = Playlist(UBound(Playlist) - 1).Name
    tvwLibrary.Nodes(x).Child.Sorted = True
End Sub

Private Sub VideoPreview_EndOfStream()
    VideoPreview.StopVideo
    VideoPreview.CloseVideo
    VideoPreview.Visible = False
End Sub

Private Sub VideoPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        VideoPreview.StopVideo
        VideoPreview.CloseVideo
        VideoPreview.Visible = False
    End If
End Sub
Private Sub UpdateStatus()
    On Error GoTo beep
    Dim i As Long
    Dim x As Long
    Dim dblSize As Double
    Dim lngLenght As Long
    
    dblSize = 0
    lngLenght = 0
    For i = 1 To lvwFilter.ListItems.Count
        x = TrackIndex(lvwFilter.ListItems(i).SubItems(13))
        dblSize = dblSize + (Library(x).Infor.Size / 1024) / 1024
        lngLenght = lngLenght + Library(x).Infor.Duration
    Next i
    lblStatus.Caption = "Found :" & lvwFilter.ListItems.Count & " Item(s)"
    lblStatus.Caption = lblStatus.Caption & " --- Total: " & Time2String(lngLenght)
    lblStatus.Caption = lblStatus.Caption & " / " & Round(dblSize, 2) & " MB"
    Exit Sub
beep:
    dblSize = 0
    lngLenght = 0
    lblStatus.Caption = "Found : 0 Item(s)"
    lblStatus.Caption = lblStatus.Caption & " --- Total: 00:00:00"
    lblStatus.Caption = lblStatus.Caption & " / 0" & " MB"
End Sub
