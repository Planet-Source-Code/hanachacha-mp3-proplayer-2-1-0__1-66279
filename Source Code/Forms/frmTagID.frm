VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTagID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M3P : MPEG Information - ID3 Editor"
   ClientHeight    =   6120
   ClientLeft      =   315
   ClientTop       =   795
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTagID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   5640
      Width           =   855
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
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Un&do Changes"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy from ID3v1"
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
      Index           =   1
      Left            =   5640
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy To ID3v1"
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
      Index           =   0
      Left            =   4200
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame frmTag 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   9255
      Begin VB.Frame frmMPG 
         ClipControls    =   0   'False
         ForeColor       =   &H80000002&
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   3975
         Begin VB.Label lblCaptionFrame 
            AutoSize        =   -1  'True
            Caption         =   "Mpeg"
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblMpg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame frmID3v2 
         ClipControls    =   0   'False
         ForeColor       =   &H80000002&
         Height          =   5055
         Left            =   4200
         TabIndex        =   17
         Top             =   0
         Width           =   4935
         Begin VB.ComboBox cboID3v2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox chkID3v2 
            Caption         =   "ID3v2 Tag"
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
            Left            =   1080
            TabIndex        =   32
            Top             =   263
            Width           =   1095
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   10
            Left            =   4200
            TabIndex        =   54
            Top             =   240
            Width           =   615
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "1085;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   9
            Left            =   1080
            TabIndex        =   53
            Top             =   4560
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   8
            Left            =   1080
            TabIndex        =   52
            Top             =   4200
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   7
            Left            =   1080
            TabIndex        =   51
            Top             =   3840
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   6
            Left            =   1080
            TabIndex        =   50
            Top             =   3480
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   5
            Left            =   1080
            TabIndex        =   49
            Top             =   3120
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   1020
            Index           =   4
            Left            =   1080
            TabIndex        =   48
            Top             =   2040
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;1799"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   3
            Left            =   1080
            TabIndex        =   47
            Top             =   1680
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "2355;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   2
            Left            =   1080
            TabIndex        =   46
            Top             =   1320
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   1
            Left            =   1080
            TabIndex        =   45
            Top             =   960
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v2 
            Height          =   300
            Index           =   0
            Left            =   1080
            TabIndex        =   44
            Top             =   585
            Width           =   3735
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6588;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblCaptionFrame 
            AutoSize        =   -1  'True
            Caption         =   "ID3v2"
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Artist :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   31
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Album :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   495
            TabIndex        =   30
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Title :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   630
            TabIndex        =   29
            Top             =   1380
            Width           =   405
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Year :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   28
            Top             =   1740
            Width           =   435
         End
         Begin VB.Label lblID3v2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Genre :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   2520
            TabIndex        =   27
            Top             =   1740
            Width           =   525
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comment :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   26
            Top             =   2340
            Width           =   780
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Composer :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   210
            TabIndex        =   25
            Top             =   3180
            Width           =   825
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Orig. Artist :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   24
            Top             =   3540
            Width           =   900
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   225
            TabIndex        =   23
            Top             =   3900
            Width           =   810
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "URL :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   645
            TabIndex        =   22
            Top             =   4230
            Width           =   390
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Encoded by :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   21
            Top             =   4620
            Width           =   945
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Track # :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   3480
            TabIndex        =   20
            Top             =   300
            Width           =   660
         End
      End
      Begin VB.Frame frmID3v1 
         ClipControls    =   0   'False
         ForeColor       =   &H80000002&
         Height          =   2415
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   3975
         Begin VB.ComboBox cboID3v1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkID3v1 
            Caption         =   "ID3v1 Tag"
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
            Left            =   840
            TabIndex        =   16
            Top             =   263
            Width           =   1095
         End
         Begin MSForms.TextBox txtID3v1 
            Height          =   300
            Index           =   5
            Left            =   3360
            TabIndex        =   42
            Top             =   240
            Width           =   495
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "873;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v1 
            Height          =   300
            Index           =   4
            Left            =   840
            TabIndex        =   41
            Top             =   2040
            Width           =   3015
            VariousPropertyBits=   746604571
            MaxLength       =   28
            BorderStyle     =   1
            Size            =   "5318;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v1 
            Height          =   300
            Index           =   3
            Left            =   840
            TabIndex        =   39
            Top             =   1680
            Width           =   735
            VariousPropertyBits=   746604571
            MaxLength       =   4
            BorderStyle     =   1
            Size            =   "1296;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v1 
            Height          =   300
            Index           =   2
            Left            =   840
            TabIndex        =   38
            Top             =   1320
            Width           =   3015
            VariousPropertyBits=   746604571
            MaxLength       =   30
            BorderStyle     =   1
            Size            =   "5318;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v1 
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   37
            Top             =   960
            Width           =   3015
            VariousPropertyBits=   746604571
            MaxLength       =   30
            BorderStyle     =   1
            Size            =   "5318;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtID3v1 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   36
            Top             =   585
            Width           =   3015
            VariousPropertyBits=   746604571
            MaxLength       =   30
            BorderStyle     =   1
            Size            =   "5318;529"
            BorderColor     =   12164479
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblCaptionFrame 
            AutoSize        =   -1  'True
            Caption         =   "ID3v1"
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Artist :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   300
            TabIndex        =   15
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Album :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   255
            TabIndex        =   14
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Title :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   390
            TabIndex        =   13
            Top             =   1380
            Width           =   405
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Year :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   12
            Top             =   1740
            Width           =   435
         End
         Begin VB.Label lblID3v2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Genre :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   1800
            TabIndex        =   11
            Top             =   1740
            Width           =   525
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comment :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   15
            TabIndex        =   10
            Top             =   2100
            Width           =   780
         End
         Begin VB.Label lblID3v2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Track # :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   2640
            TabIndex        =   9
            Top             =   300
            Width           =   660
         End
      End
   End
   Begin VB.TextBox txtFileName 
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
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   80
      Width           =   9015
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmTagID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bolShow As Boolean

Dim i As Integer
Dim strFilePath As String
Dim TempID3v1 As New clsID3v1
Dim tagID3v1 As New clsID3v1
Dim tagID3v2 As New clsID3v2
Dim TempID3v2 As New clsID3v2


Private Sub chkID3v1_Click()
    If chkID3v1.value = 0 Then
        For i = 0 To 5
            txtID3v1(i).Enabled = False
        Next i
        cboID3v1.Enabled = False
        cmdCopy(1).Enabled = False
    Else
        For i = 0 To 5
            txtID3v1(i).Enabled = True
        Next i
        cboID3v1.Enabled = True
        cmdCopy(1).Enabled = True
    End If
End Sub

Private Sub chkID3v2_Click()
    If chkID3v2.value = 0 Then
        For i = 0 To 10
            txtID3v2(i).Enabled = False
        Next i
        cboID3v2.Enabled = False
        cmdCopy(0).Enabled = False
    Else
        For i = 0 To 10
            txtID3v2(i).Enabled = True
        Next i
        cboID3v2.Enabled = True
        cmdCopy(0).Enabled = True
    End If
End Sub

Private Sub cmdCopy_Click(Index As Integer)
    Select Case Index
        Case 0
            For i = 0 To 4
                txtID3v1(i).text = txtID3v2(i).text
            Next
            txtID3v1(5).text = txtID3v2(10).text
            cboID3v1.ListIndex = cboID3v2.ListIndex
        Case 1
            For i = 0 To 4
                txtID3v2(i).text = txtID3v1(i).text
            Next
            txtID3v2(10).text = txtID3v1(5).text
            cboID3v2.ListIndex = cboID3v1.ListIndex
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUndo_Click()
    On Error Resume Next
'Undo ID3v1
    Set tagID3v1 = TempID3v1
    Call tagID3v1.WriteTag(strFilePath)
    If tagID3v1.ReadTag(strFilePath) = False Then
        chkID3v1.value = 0
    Else
        chkID3v1.value = 1
        txtID3v1(0).text = Trim(tagID3v1.Artist)
        txtID3v1(2).text = Trim(tagID3v1.Title)
        txtID3v1(1).text = Trim(tagID3v1.Album)
        txtID3v1(3).text = Trim(tagID3v1.Year)
        txtID3v1(4).text = Trim(tagID3v1.Comment)
        txtID3v1(5).text = tagID3v1.track
        cboID3v1.ListIndex = tagID3v1.Genre
    End If
'Undo ID3v2
    Set tagID3v2 = TempID3v2
    Call tagID3v2.WriteID3v2Tag(strFilePath)
    If tagID3v2.ReadID3v2Tag(strFilePath) = False Then
        chkID3v2.value = 0
    Else
        chkID3v2.value = 1
        txtID3v2(0).text = tagID3v2.Artist
        txtID3v2(1).text = tagID3v2.Album
        txtID3v2(2).text = tagID3v2.Title
        txtID3v2(3).text = tagID3v2.Year
        txtID3v2(4).text = tagID3v2.Comments
        txtID3v2(5).text = tagID3v2.Composer
        txtID3v2(6).text = tagID3v2.OrigArtist
        txtID3v2(7).text = tagID3v2.Copyright
        txtID3v2(8).text = tagID3v2.url
        txtID3v2(9).text = tagID3v2.EncodedBy
        txtID3v2(10).text = tagID3v2.track
        If Mid(tagID3v2.Genre, 1, 1) = "(" Then
            cboID3v2.ListIndex = CLng(Mid(tagID3v2.Genre, 2, Len(tagID3v2.Genre) - 2))
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo beep
    
        If chkID3v1.value = 1 Then
            With tagID3v1
                .Artist = txtID3v1(0).text
                .Title = txtID3v1(2).text
                .Album = txtID3v1(1).text
                .Year = txtID3v1(3).text
                .Comment = txtID3v1(4).text
                .Genre = CByte(cboID3v1.ListIndex)
                .track = txtID3v1(5).text
            End With
            Call tagID3v1.WriteTag(strFilePath)
        Else
            Call tagID3v1.RemoveID3v1(strFilePath)
        End If
        
        If chkID3v2.value = 1 Then
            With tagID3v2
                .Artist = txtID3v2(0).text
                .Album = txtID3v2(1).text
                .Title = txtID3v2(2).text
                .Year = txtID3v2(3).text
                .Comments = txtID3v2(4).text
                .Composer = txtID3v2(5).text
                .OrigArtist = txtID3v2(6).text
                .Copyright = txtID3v2(7).text
                .url = txtID3v2(8).text
                .EncodedBy = txtID3v2(9).text
                .track = txtID3v2(10).text
                .Genre = cboID3v2.List(cboID3v2.ListIndex)
            End With
            Call tagID3v2.WriteID3v2Tag(strFilePath)
        Else
            tagID3v2.RemoveID3v2tag strFilePath
        End If
        
        Call tagID3v1.ReadTag(strFilePath)
        Unload Me
beep:
    If Err.Number <> 0 Then
        MsgBox "Can not update file is playing !!!", vbInformation, "Update Error"
        Exit Sub
    End If
End Sub


Private Sub Form_Load()
    On Error Resume Next
    Dim hasID3v1 As Boolean
    Dim hasID3v2 As Boolean
    
    bolShow = True
    Me.Icon = LoadResPicture(112, vbResIcon)
    
    For i = LBound(GenreArray) To UBound(GenreArray)
        cboID3v1.AddItem GenreArray(i)
        cboID3v2.AddItem GenreArray(i)
    Next i
    
    txtFileName.text = NowPlaying(frmPlayList.List.Key(currentRIndex)).Infor.FullName
    strFilePath = txtFileName.text
    'ID3v1
        hasID3v1 = tagID3v1.ReadTag(strFilePath)
        Set TempID3v1 = tagID3v1 'save to temp
        
        If hasID3v1 = False Then
            chkID3v1.value = 0
            For i = 0 To 5
                txtID3v1(i).Enabled = False
            Next i
            cboID3v1.Enabled = False
        Else
            chkID3v1.value = 1
            txtID3v1(0).text = Trim(tagID3v1.Artist)
            txtID3v1(2).text = Trim(tagID3v1.Title)
            txtID3v1(1).text = Trim(tagID3v1.Album)
            txtID3v1(3).text = Trim(tagID3v1.Year)
            txtID3v1(4).text = Trim(tagID3v1.Comment)
            txtID3v1(5).text = tagID3v1.track
            cboID3v1.ListIndex = tagID3v1.Genre
        End If
        
    'ID3v2
        hasID3v2 = tagID3v2.ReadID3v2Tag(strFilePath)
        Set TempID3v2 = tagID3v2
        
        If hasID3v2 = False Then
            chkID3v2.value = 0
            For i = 0 To 10
                txtID3v2(i).Enabled = False
            Next i
            cboID3v2.Enabled = False
        Else
            chkID3v2.value = 1
            txtID3v2(0).text = tagID3v2.Artist
            txtID3v2(1).text = tagID3v2.Album
            txtID3v2(2).text = tagID3v2.Title
            txtID3v2(3).text = tagID3v2.Year
            txtID3v2(4).text = tagID3v2.Comments
            txtID3v2(5).text = tagID3v2.Composer
            txtID3v2(6).text = tagID3v2.OrigArtist
            txtID3v2(7).text = tagID3v2.Copyright
            txtID3v2(8).text = tagID3v2.url
            txtID3v2(9).text = tagID3v2.EncodedBy
            txtID3v2(10).text = tagID3v2.track
            If Mid(tagID3v2.Genre, 1, 1) = "(" Then
                cboID3v2.ListIndex = CLng(Mid(tagID3v2.Genre, 2, Len(tagID3v2.Genre) - 2))
            Else
                cboID3v2.ListIndex = ReturnGenreID(tagID3v2.Genre)
            End If
        End If
    'Mpeg Infor
        Dim MPG As New clsMPEG
        Call MPG.ReadMPEGHeader(txtFileName)
        
        lblMpg.Caption = "Size : " & MPG.FileSize & " bytes" & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Version : " & MPG.version
        lblMpg.Caption = lblMpg.Caption & "  _  Layer " & IIf(MPG.Layer = "1", "I", IIf(MPG.version = "2", "II", "III")) & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Bitrate : " & MPG.bitrate & " Kbps" & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Frequency : " & MPG.Frequency & " Khz" & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "CCR : " & Bol2String(MPG.HasCRC) & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "VBR : " & Bol2String(MPG.HasVBR) & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Mode : " & MPG.ChannelMode & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Time lenght : " & MPG.length & " seconds" & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Copyrighted : " & Bol2String(MPG.Copyrighted) & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Emphasis : " & Bol2String(MPG.HasEmphasis) & vbNewLine
        lblMpg.Caption = lblMpg.Caption & "Original : " & Bol2String(MPG.Original)
        Set MPG = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bolShow = False
    Set TempID3v1 = Nothing
    Set tagID3v1 = Nothing
    Set tagID3v2 = Nothing
    Set TempID3v2 = Nothing
End Sub
