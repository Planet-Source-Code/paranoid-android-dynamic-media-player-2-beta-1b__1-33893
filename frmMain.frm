VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1E5ED54D-4BB2-11D6-8DC1-90B225C3E54F}#1.0#0"; "GRADBUTTON.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H007D3F00&
   BorderStyle     =   0  'None
   Caption         =   "DMP2 BETA 1b"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMenuSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00E37200&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   760
      Left            =   2530
      ScaleHeight     =   765
      ScaleWidth      =   1110
      TabIndex        =   27
      Top             =   2200
      Visible         =   0   'False
      Width           =   1110
      Begin GradButton.GradientButton cmdFullscreen 
         Height          =   255
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "Add"
         Top             =   0
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   450
         Appearance      =   0
         BackColor       =   13128960
         BorderColor     =   16777215
         ButtonType      =   1
         Caption         =   "Fullscreen"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   16773055
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777152
         GradientAngle   =   100
         GradientColor1  =   6946816
         GradientColor2  =   16743194
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16777152
         UseHover        =   0   'False
      End
      Begin GradButton.GradientButton cmdAddSize 
         Height          =   255
         Left            =   0
         TabIndex        =   29
         ToolTipText     =   "Add"
         Top             =   255
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Appearance      =   0
         BackColor       =   13128960
         BorderColor     =   16777215
         Caption         =   "+"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   16773055
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777152
         GradientAngle   =   100
         GradientColor1  =   6946816
         GradientColor2  =   16743194
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16777152
         UseHover        =   0   'False
      End
      Begin GradButton.GradientButton cmdRemoveSize 
         Height          =   255
         Left            =   550
         TabIndex        =   30
         ToolTipText     =   "Remove"
         Top             =   255
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Appearance      =   0
         BackColor       =   13128960
         BorderColor     =   16777215
         Caption         =   "-"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   16773055
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777152
         GradientAngle   =   100
         GradientColor1  =   6946816
         GradientColor2  =   16743194
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16777152
         UseHover        =   0   'False
      End
      Begin GradButton.GradientButton cmdActual 
         Height          =   255
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "Add"
         Top             =   510
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   450
         Appearance      =   0
         BackColor       =   13128960
         BorderColor     =   16777215
         Caption         =   "Actual Size"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   16773055
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777152
         GradientAngle   =   100
         GradientColor1  =   6946816
         GradientColor2  =   16743194
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16777152
         UseHover        =   0   'False
      End
   End
   Begin VB.PictureBox picMenuRemove 
      Appearance      =   0  'Flat
      BackColor       =   &H00CE6700&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   385
      ScaleHeight     =   255
      ScaleWidth      =   1215
      TabIndex        =   18
      Top             =   2410
      Visible         =   0   'False
      Width           =   1215
      Begin GradButton.GradientButton cmdClearPlaylist 
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Appearance      =   0
         BackColor       =   13128960
         BorderColor     =   16777215
         Caption         =   "Clear Playlist"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   16773055
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777152
         GradientAngle   =   100
         GradientColor1  =   6946816
         GradientColor2  =   16743194
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseHover        =   0   'False
      End
   End
   Begin VB.PictureBox picScroll2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H007D3F00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   2855
      MousePointer    =   9  'Size W E
      ScaleHeight     =   225
      ScaleWidth      =   30
      TabIndex        =   39
      Top             =   250
      Width           =   30
   End
   Begin VB.PictureBox picScroll1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H007D3F00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   350
      MousePointer    =   9  'Size W E
      ScaleHeight     =   225
      ScaleWidth      =   30
      TabIndex        =   40
      Top             =   250
      Width           =   30
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H007D3F00&
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5865
      TabIndex        =   1
      Top             =   3240
      Width           =   5895
      Begin prjDMPBETA.SliderControl sldMedia 
         Height          =   135
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   238
      End
      Begin GradButton.GradientButton cmdAdvanced 
         Height          =   255
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "Volume"
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Appearance      =   0
         BackColor       =   13128960
         BorderColor     =   16777215
         Caption         =   "Advanced"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   16773055
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777152
         GradientAngle   =   100
         GradientColor1  =   6946816
         GradientColor2  =   16743194
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16777152
         HoverMode       =   1
         UseHover        =   0   'False
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H009F5000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3225
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.PictureBox picMenuAdd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CE6700&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1015
         Left            =   120
         ScaleHeight     =   1020
         ScaleWidth      =   1215
         TabIndex        =   54
         Top             =   1630
         Visible         =   0   'False
         Width           =   1215
         Begin GradButton.GradientButton cmdLoadPlaylist 
            Height          =   255
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Load Playlist"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdSavePlaylist 
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   250
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Save Playlist"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdAppend 
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   510
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Append"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdAddFiles 
            Height          =   255
            Left            =   0
            TabIndex        =   58
            Top             =   760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Add Files"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseHover        =   0   'False
         End
      End
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H009F5000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2860
         ScaleHeight     =   225
         ScaleWidth      =   1965
         TabIndex        =   35
         Top             =   220
         Width           =   2000
         Begin VB.PictureBox picScroll3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H007D3F00&
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   1920
            MousePointer    =   9  'Size W E
            ScaleHeight     =   225
            ScaleWidth      =   30
            TabIndex        =   53
            Top             =   -10
            Width           =   55
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   45
            TabIndex        =   38
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.PictureBox picArtist 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H009F5000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   2445
         TabIndex        =   34
         Top             =   220
         Width           =   2470
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Artist"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   45
            TabIndex        =   37
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.PictureBox picNumbers 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H009F5000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -10
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   33
         Top             =   220
         Width           =   350
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   50
            TabIndex        =   36
            Top             =   0
            Width           =   105
         End
      End
      Begin VB.Timer tmrEndOfFile 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5400
         Top             =   1080
      End
      Begin VB.Timer tmrMisc 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   5400
         Top             =   600
      End
      Begin VB.Timer tmrMenu 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   5400
         Top             =   1560
      End
      Begin MSComDlg.CommonDialog PlaylistDialog 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   4
      End
      Begin VB.ListBox lstFilenames 
         Appearance      =   0  'Flat
         Height          =   225
         ItemData        =   "frmMain.frx":0482
         Left            =   5520
         List            =   "frmMain.frx":0484
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H007D3F00&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   220
         Left            =   -10
         ScaleHeight     =   225
         ScaleWidth      =   5910
         TabIndex        =   3
         Top             =   0
         Width           =   5910
         Begin GradButton.GradientButton cmdExit 
            Height          =   225
            Left            =   5625
            TabIndex        =   17
            ToolTipText     =   "Exit"
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            Appearance      =   0
            BackColor       =   8208128
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings 2"
               Size            =   11.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientColor1  =   14905856
            GradientColor2  =   8665088
            GradientType    =   1
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":0486
            Style           =   1
         End
         Begin GradButton.GradientButton cmdOnTop 
            Height          =   225
            Left            =   5115
            TabIndex        =   22
            ToolTipText     =   "Always On Top"
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            Appearance      =   0
            BackColor       =   8208128
            BorderColor     =   16777215
            ButtonType      =   1
            Caption         =   "T"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   5.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientColor1  =   14905856
            GradientColor2  =   8665088
            GradientType    =   1
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
         End
         Begin GradButton.GradientButton cmdMinimize 
            Height          =   225
            Left            =   5370
            TabIndex        =   24
            ToolTipText     =   "Minimize"
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            Alignment       =   8
            Appearance      =   0
            BackColor       =   8208128
            BorderColor     =   16777215
            Caption         =   "_"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientColor1  =   14905856
            GradientColor2  =   8665088
            GradientType    =   1
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dynamic Media Player 2 BETA 1b"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H009F5000&
         ForeColor       =   &H80000008&
         Height          =   310
         Left            =   -10
         ScaleHeight     =   285
         ScaleWidth      =   5865
         TabIndex        =   2
         Top             =   2920
         Width           =   5895
         Begin GradButton.GradientButton chkRepeat1 
            Height          =   255
            Left            =   105
            TabIndex        =   20
            ToolTipText     =   "Repeat One"
            Top             =   15
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            ButtonType      =   1
            Caption         =   "Repeat One"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            HoverMode       =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton chkRepeatAll 
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            ToolTipText     =   "Repeat All"
            Top             =   15
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            ButtonType      =   1
            Caption         =   "Repeat All"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            HoverMode       =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdVolume 
            Height          =   255
            Left            =   4920
            TabIndex        =   23
            ToolTipText     =   "Volume"
            Top             =   15
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Volume"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            HoverMode       =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdSizeMenu 
            Height          =   255
            Left            =   2520
            TabIndex        =   26
            ToolTipText     =   "Video Size Menu"
            Top             =   15
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Video Size"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            HoverMode       =   1
            UseHover        =   0   'False
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H009F5000&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -10
         ScaleHeight     =   285
         ScaleWidth      =   5865
         TabIndex        =   4
         Top             =   2620
         Width           =   5895
         Begin GradButton.GradientButton cmdAdd 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Add"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "+"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdRemove 
            Height          =   255
            Left            =   375
            TabIndex        =   8
            ToolTipText     =   "Remove"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "-"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdID3 
            Height          =   255
            Left            =   630
            TabIndex        =   9
            ToolTipText     =   "Tag Editor (ID3)"
            Top             =   10
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            Caption         =   "Tag Editor"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16773055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverForeColor  =   16777152
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdPrev 
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            ToolTipText     =   "Previous File"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":0564
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdPlay 
            Height          =   255
            Left            =   2295
            TabIndex        =   11
            ToolTipText     =   "Play"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":066A
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdPause 
            Height          =   255
            Left            =   2550
            TabIndex        =   12
            ToolTipText     =   "Pause"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":074C
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdStop 
            Height          =   255
            Left            =   2805
            TabIndex        =   13
            ToolTipText     =   "Stop"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":0876
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdNext 
            Height          =   255
            Left            =   3060
            TabIndex        =   14
            ToolTipText     =   "Next File"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":0958
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton cmdCloseFile 
            Height          =   255
            Left            =   3480
            TabIndex        =   25
            ToolTipText     =   "Close File"
            Top             =   10
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Appearance      =   0
            BackColor       =   13128960
            BorderColor     =   16777215
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownForeColor   =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777152
            GradientAngle   =   100
            GradientColor1  =   6946816
            GradientColor2  =   16743194
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            Picture         =   "frmMain.frx":0A5E
            Style           =   1
            UseHover        =   0   'False
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "C00:00:00.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   255
            Left            =   3840
            TabIndex        =   15
            Top             =   35
            Width           =   1995
         End
      End
      Begin MSComctlLib.ListView lstPlaylist 
         Height          =   2175
         Left            =   -15
         TabIndex        =   6
         Top             =   470
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   12541952
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   5880
         Y1              =   225
         Y2              =   225
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H009F5000&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   100
      ScaleHeight     =   1215
      ScaleWidth      =   5655
      TabIndex        =   41
      Top             =   3720
      Width           =   5680
      Begin VB.Timer tmrAdvanced 
         Interval        =   100
         Left            =   5160
         Top             =   720
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   5640
         X2              =   5640
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5640
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Label lblDefaultDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Default Device: None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2760
         TabIndex        =   52
         Top             =   0
         Width           =   1710
      End
      Begin VB.Label lblActualVideoH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Actual Video Height: -1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2760
         TabIndex        =   51
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label lblCurrentVideoH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Current Video Height: -1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2760
         TabIndex        =   50
         Top             =   480
         Width           =   2025
      End
      Begin VB.Label lblActualVideoW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Actual Video Width: -1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2760
         TabIndex        =   49
         Top             =   720
         Width           =   1830
      End
      Begin VB.Label lblCurrentVideoW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Current Video Width: -1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2760
         TabIndex        =   48
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Device Used: None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label lblTotalFrames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Total Frames: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblFPS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " FPS: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   44
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Multimedia Status: None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   43
         Top             =   960
         Width           =   2040
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Percent Completed: 0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   42
         Top             =   240
         Width           =   1965
      End
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   5880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line7 
      X1              =   5880
      X2              =   5880
      Y1              =   3600
      Y2              =   5040
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   0
      Y1              =   3600
      Y2              =   5040
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'* Dynamic Media Player 2 BETA 1b *
'*     Build Date: 5/27/02        *
'*     By Paranoid Android        *
'**********************************

'If you use any part of my program, I am going to
'Have to insist that you write me into the credits
'Because, you wouldn't have been the one who
'Wrote it, thanks.

'I want to thank Abdullah Al-Ahdal for his
'Wonderfully coded Multimedia module
'Without it this program would really suck.

'For any request you can contact me at:
'crazygoat8@hotmail.com

Option Explicit 'Error handle!
Public PlaylistFilename As String 'The current playlist filename
Public ID3FileName As String 'The current filename to be processed through the ID3
Const AliasName As String = "DMPMedia" 'The alias of the player, used later on
Public typeDevice As String 'The device used to play the file
Public Result As String 'The result, whether it had an error or not
Public Percent As Long 'The percent completed
Public PlaylistIndex As Integer 'The current playlist ListIndex
Private TimeMode As String 'The time mode, current or remaining
Private CurrentTime As String 'The current time
Private CurrentPosition As String 'The current frame
Public FramesPerSecond As String 'The current FPS
Private TotalTime As String 'The total ammount of time
Private TotalFrames As String 'The total ammount of frames
Private Paused As Boolean 'Whether the file is paused or not
Public CurrentFilename As String 'The current filename opened
Public tmpWidth, tmpHeight As Integer 'The temporary width and height of the video
Public AddedSize As Integer 'The ammount added to the size of the video
Public ActualWidth As Integer 'The actual width of the video
Public ActualHeight As Integer 'The actual height of the video
Public FullscreenWidth As Integer 'The width of the video in fullscreen mode
Public FullscreenHeight As Integer 'The height of the video in fullscreen mode
Public Status As String 'The current status of the device, playing, paused, etc.
Public CDAudioEnabled As Boolean 'Tells whether it is in CD Audio mode

Private Sub cmdActual_Click()
    
    'If the current open filename is a video then...
    If LCase$(Right$(CurrentFilename, 4)) = ".avi" Or LCase$(Right$(CurrentFilename, 4)) = ".mpg" Or LCase$(Right$(CurrentFilename, 5)) = ".mpeg" Or LCase$(Right$(CurrentFilename, 4)) = ".mpe" Or LCase$(Right$(CurrentFilename, 4)) = ".m1v" Or LCase$(Right$(CurrentFilename, 4)) = ".mp2" Or LCase$(Right$(CurrentFilename, 5)) = ".mpv2" Or LCase$(Right$(CurrentFilename, 4)) = ".mpa" Then
        ResizeVideo (0) 'Resize the video with the actual width and height
        AddedSize = 0
    End If
    
    picMenuSize.Visible = False 'Hide the menu
    
End Sub

Private Sub cmdActual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    tmrMenu.Enabled = False 'Disable the menu timer.
    
End Sub

Private Sub cmdAdd_Click()

    'Shows or hides the add menu
    If picMenuAdd.Visible = False Then
        picMenuAdd.Visible = True
        picMenuRemove.Visible = False
        picMenuSize.Visible = False
    Else
        picMenuAdd.Visible = False
        picMenuRemove.Visible = False
        picMenuSize.Visible = False
    End If

End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False 'Disable the menu timer

End Sub

Private Sub cmdAddFiles_Click()

    On Error Resume Next 'Dont want any errors

    picMenuAdd.Visible = False 'Hide the add menu
    
    With PlaylistDialog 'Using the Playlists Common Dialog Control
    
    'To be able to select multiple files we need to set the flags
    .Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNLongNames + &O4
    
    .MaxFileSize = 20000

    'Clear the filename out of the box
    .filename = ""
    
    'We want to add files not playlists...
    .DialogTitle = "Add Files"
    
    'Change the filters
    .Filter = "All Media Types|*.avi;*.mpg;*.dat;*.mpeg;*.mpe;*.mp3;*.mp2;*.mp1;*.wav;*.aif;*.aiff;*.aifc;*.au;*.m1v;*.mov;*.mpa;*.qt;*.snd;*.mpm;*.mpv;*.enc;*.mid;*.rmi;*.vob;*.wma;*.wmv;*.wmp;*.wmx;*.wax|Windows Video (avi)|*.avi|Windows Audio (wma,wax)|*.wma;*.wax|Windows Audio/Video (wmp,wmv,wmx)|*.wmp;*.wmv;*.wmx|Wav Audio (wav)|*.wav|MPEG Video (mpeg,mpg,mpe,m1v,mp2,mpv2,mpa)|*.mpeg;*.mpg;*.mpe;*.m1v;*.mp2;*.mpv2;*.mpa|MPEG Audio (mp3)|*.mp3|MIDI Music (mid,midi,rmi)|*.mid;*.midi;*.rmi|AIFF Sound (aif,aifc,aiff)|*.aif;*.aifc;*.aiff|AU Sound (au,snd)|*.au;*.snd"
    .ShowOpen

    'If you select nothing then exit sub
    If .filename = "" Then Exit Sub

    Dim i As Integer

    'Parse all of the files in the list
    For i = 1 To CountFilesInList(.filename)
        ParseFiles (GetFileFromList(.filename, i))
    Next
    End With
    
    On Error GoTo 0
    
End Sub

Private Sub cmdAddFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False 'Disable the menu timer

End Sub

Private Sub cmdAddSize_Click()

    'If the current filename is a video
    If LCase$(Right$(CurrentFilename, 4)) = ".avi" Or LCase$(Right$(CurrentFilename, 4)) = ".mpg" Or LCase$(Right$(CurrentFilename, 5)) = ".mpeg" Or LCase$(Right$(CurrentFilename, 4)) = ".mpe" Or LCase$(Right$(CurrentFilename, 4)) = ".m1v" Or LCase$(Right$(CurrentFilename, 4)) = ".mp2" Or LCase$(Right$(CurrentFilename, 5)) = ".mpv2" Or LCase$(Right$(CurrentFilename, 4)) = ".mpa" Then
        If AddedSize >= 0 And AddedSize < 10000 Then
            AddedSize = AddedSize + 200 'Add 200 in width and height, so it's symmetrical
            ResizeVideo (AddedSize) 'Process the new size
            picMenuSize.Visible = False 'Hide the size menu
        End If
    End If
    
End Sub

Private Sub cmdAddSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    tmrMenu.Enabled = False 'Disable the menu timer
    
End Sub

Private Sub cmdAdvanced_Click()

    Dim X
    
    'This just increases or decreases the height of the form
    'So that you can see the advanced information section
    If Me.Height = 3645 Then
        For X = 0 To 1410 Step 5
            Me.Height = 3645 + X
        Next X
    Else
        For X = 0 To 1410 Step 5
            Me.Height = 5055 - X
        Next X
    End If

End Sub

Private Sub cmdAppend_Click()

    'This just adds a playlist to your current playlist
    On Error Resume Next
    Dim FilePath$, tmpString$, i%, FindComma%
    Dim cM3U As New clsOpenM3U, Count&, X&
    
    picMenuAdd.Visible = False 'Hide the add menu
    
    'Change the filters again
    PlaylistDialog.Filter = "M3U Playlist (*.m3u)|*.m3u"
    PlaylistDialog.DialogTitle = "Append Playlist"
    PlaylistDialog.ShowOpen
    PlaylistDialog.CancelError = True
    
    If PlaylistDialog.filename = "" Then Exit Sub
      
    cM3U.filename = PlaylistDialog.filename
    PlaylistFilename = cM3U.filename
    Count = cM3U.Refresh

    'Parse the filenames in the playlist
    For X = 0 To cM3U.Count - 1
        ParseFiles (cM3U.FileTitle(X))
    Next X

    On Error GoTo 0
    
End Sub

Private Sub cmdAppend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False 'Disable the menu timer

End Sub

Private Sub cmdClearPlaylist_Click()

    'This clears the items out of the playlist
    picMenuRemove.Visible = False
    
    lstFilenames.Clear
    lstPlaylist.ListItems.Clear
    PlaylistFilename = ""
    
    StopMedia 'If something is playing then it is stopped
    CloseMedia 'The current file is closed

End Sub

Private Sub cmdClearPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False  'Disable the menu timer

End Sub

Private Sub cmdCloseFile_Click()
    
    'This closes the current file
    StopMedia 'Stop it
    CloseMedia 'Close it
    
    If frmMovieView.Visible = True Then
        Unload frmMovieView
        frmMovieView.Hide
    End If

    CurrentFilename = "" 'Reset the current filename to nothing
    
End Sub

Private Sub cmdExit_Click()

    'This closes all forms in the project
    Dim Frm As Form
    
    AlwaysOnTop frmMain, False
    
    For Each Frm In Forms
        Unload Frm
        Set Frm = Nothing
    Next Frm
    
    CloseAll 'This closes all multimedia devices open

End Sub

Private Sub cmdFullscreen_Click()

    'This checks whether the file is a video and then it...
    'Will set it to fullscreen mode
    If LCase$(Right$(CurrentFilename, 4)) = ".avi" Or LCase$(Right$(CurrentFilename, 4)) = ".mpg" Or LCase$(Right$(CurrentFilename, 5)) = ".mpeg" Or LCase$(Right$(CurrentFilename, 4)) = ".mpe" Or LCase$(Right$(CurrentFilename, 4)) = ".m1v" Or LCase$(Right$(CurrentFilename, 4)) = ".mp2" Or LCase$(Right$(CurrentFilename, 5)) = ".mpv2" Or LCase$(Right$(CurrentFilename, 4)) = ".mpa" Then
        If cmdFullscreen.Value = True Then
            frmMovieView.Caption = ""
            frmMovieView.left = 0
            frmMovieView.top = 0
            frmMovieView.Width = Screen.Width
            frmMovieView.Height = Screen.Height
            Result = PutMultimedia(frmMovieView.hwnd, AliasName, Val(0), Val(0), Val(0), Val(0))
            FullscreenWidth = GetSize(AliasName, "cx")
            FullscreenHeight = GetSize(AliasName, "cy")
            AlwaysOnTop frmMovieView, True
            ShowCursor (False)
        Else
            frmMain.cmdFullscreen.Value = False
            frmMovieView.Caption = "DMP2 BETA 1b"
            AlwaysOnTop frmMovieView, False
            frmMovieView.Width = (ActualWidth * 15) + AddedSize
            frmMovieView.Height = (ActualHeight * 15) + AddedSize
            Result = PutMultimedia(frmMovieView.hwnd, AliasName, Val(0), Val(0), Val(0), Val(0))
            frmMovieView.Height = frmMovieView.Height + 350
            frmMovieView.left = frmMain.left
            frmMovieView.top = frmMain.top
            ShowCursor (True)
         End If
    Else
        cmdFullscreen.Value = False
    End If
    picMenuSize.Visible = False

End Sub

Private Sub cmdFullscreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    tmrMenu.Enabled = False 'Disable the menu timer
    
End Sub

Private Sub cmdID3_Click()

'This will open the built-in ID3 editor for you to...
'Edit the contents of the information stored in the media
On Error Resume Next
    
    If ID3FileName <> "" Then
    
        Me.Enabled = False
        Load frmID3
        frmID3.Show
        
    End If
       
On Error GoTo 0

End Sub

Private Sub cmdLoadPlaylist_Click()

    'This will load a playlist and parse the files in it
    On Error Resume Next
    Dim FilePath$, tmpString$, i%, FindComma%
    Dim cM3U As New clsOpenM3U, Count&, X&
    
    picMenuAdd.Visible = False

    lstFilenames.Clear
    lstPlaylist.ListItems.Clear
    
    PlaylistDialog.Filter = "M3U Playlist (*.m3u)|*.m3u"
    PlaylistDialog.DialogTitle = "Load Playlist"
    PlaylistDialog.ShowOpen
    PlaylistDialog.CancelError = True
    
    If PlaylistDialog.filename = "" Then Exit Sub
      
    cM3U.filename = PlaylistDialog.filename
    PlaylistFilename = cM3U.filename
    Count = cM3U.Refresh

    For X = 0 To cM3U.Count - 1
        ParseFiles (cM3U.FileTitle(X))
    Next X

    On Error GoTo 0
    
End Sub

Private Sub cmdLoadPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False 'Disable the menu timer

End Sub

Private Sub cmdMinimize_Click()

    'It simply minimizes the window
    frmMain.WindowState = vbMinimized

End Sub

Private Sub cmdNext_Click()

    'This will select the next file in the list...
    'And if it is the last file it will select the first one...
    'In the list
    AcquireNextFile

End Sub

Private Sub cmdOnTop_Click()

    'This will set the form ontop of every other form
    If cmdOnTop.Value = True Then
        AlwaysOnTop frmMain, True
    Else
        AlwaysOnTop frmMain, False
    End If

End Sub

Private Sub cmdPause_Click()
    
    'This will pause the media
    If Paused = True Then
        Paused = False
        ResumeMedia
    Else
        Paused = True
        PauseMedia
    End If

End Sub

Private Sub cmdPlay_Click()

    'This will play the media
    If Status = "playing" Then Exit Sub

    If Paused = True Then
        Paused = False
        ResumeMedia
    Else
        PlayMedia
        
        If LCase$(Right$(CurrentFilename, 4)) = ".avi" Or LCase$(Right$(CurrentFilename, 4)) = ".mpg" Or LCase$(Right$(CurrentFilename, 5)) = ".mpeg" Or LCase$(Right$(CurrentFilename, 4)) = ".mpe" Or LCase$(Right$(CurrentFilename, 4)) = ".m1v" Or LCase$(Right$(CurrentFilename, 4)) = ".mp2" Or LCase$(Right$(CurrentFilename, 5)) = ".mpv2" Or LCase$(Right$(CurrentFilename, 4)) = ".mpa" Then
            Load frmMovieView
            frmMovieView.Show
        End If
        
        If Paused = True Then
            ResumeMedia
        End If
    End If

End Sub

Private Sub cmdPrev_Click()

    'This does the same as AcquireNextFile except it...
    'Selects the previous file.
    AcquirePreviousFile

End Sub

Private Sub cmdRemove_Click()

    'This will bring up the remove menu
    If picMenuRemove.Visible = False Then
        picMenuRemove.Visible = True
        picMenuAdd.Visible = False
        picMenuSize.Visible = False
    Else
        picMenuRemove.Visible = False
        picMenuAdd.Visible = False
        picMenuSize.Visible = False
    End If
    
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False 'Disable the menu timer

End Sub

Private Sub cmdRemoveSize_Click()
    
    'This will subtrack from the size of the video
    If LCase$(Right$(CurrentFilename, 4)) = ".avi" Or LCase$(Right$(CurrentFilename, 4)) = ".mpg" Or LCase$(Right$(CurrentFilename, 5)) = ".mpeg" Or LCase$(Right$(CurrentFilename, 4)) = ".mpe" Or LCase$(Right$(CurrentFilename, 4)) = ".m1v" Or LCase$(Right$(CurrentFilename, 4)) = ".mp2" Or LCase$(Right$(CurrentFilename, 5)) = ".mpv2" Or LCase$(Right$(CurrentFilename, 4)) = ".mpa" Then
        If AddedSize >= 0 And AddedSize < 10000 Then
            AddedSize = AddedSize - 200
            ResizeVideo (AddedSize)
            picMenuSize.Visible = False
        End If
    End If
    
End Sub

Private Sub cmdRemoveSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    tmrMenu.Enabled = False 'Disable the menu timer
    
End Sub

Private Sub cmdSavePlaylist_Click()

    'This will save the current playlist
    picMenuAdd.Visible = False
    
    On Error Resume Next
    
    'If your playlist is empty, bring up a message box
    If lstFilenames.ListCount = 0 Then
        MsgBox "You can't save an empty playlist." & vbCrLf & "Please create your playlist before saving.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    'Change the filters
    PlaylistDialog.Filter = "M3U Playlist (*.m3u)|*.m3u"
    PlaylistDialog.ShowSave
    PlaylistDialog.CancelError = True
    PlaylistDialog.DialogTitle = "Save Playlist"
    
    If PlaylistDialog.filename = "" Then Exit Sub
      
    'And finally the saving part
    SavePlaylist PlaylistDialog.filename

    On Error GoTo 0

End Sub

Private Sub cmdSavePlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = False 'Disable the menu timer

End Sub

Private Sub cmdSizeMenu_Click()

    'This brings up the video size menu
    If picMenuSize.Visible = False Then
        picMenuAdd.Visible = False
        picMenuRemove.Visible = False
        picMenuSize.Visible = True
    Else
        picMenuAdd.Visible = False
        picMenuRemove.Visible = False
        picMenuSize.Visible = False
    End If
    
End Sub

Private Sub cmdSizeMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   tmrMenu.Enabled = False
    
End Sub

Private Sub cmdStop_Click()

    'This will stop the current media
    StopMedia
    frmMovieView.Hide

End Sub

Private Sub cmdVolume_Click()
    
    'This brings up the volume form
    frmVolumeControl.Show

End Sub

Private Sub Form_Activate()
    
    'Makes the form not always ontop
    AlwaysOnTop frmMain, False
    
End Sub

Public Sub FormDrag(TheForm As Form)

    'Allows for the form to be moved without a titlebar
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Form_Load()

    CloseMedia
    
    AddedSize = 0
    
    TimeMode = "Current"
    
    'Disable the screensaver
    Call ScreenSaverActive(False)
    
    'Sets all of the default devices up
    If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
        SetDefaultDevice "MPEGVideo", "mciqtz.drv"
    End If

    If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
        SetDefaultDevice "avivideo", "mciavi.drv"
    End If
    
    ActualWidth = -1
    ActualHeight = -1
    
End Sub

Private Sub Form_Paint()

    If cmdOnTop.Value = True Then
        AlwaysOnTop frmMain, True
    Else
        AlwaysOnTop frmMain, False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Re-enable the screensaver
    Call ScreenSaverActive(True)

End Sub

Private Sub Label1_DblClick()

    picScroll1.left = 350
    picNumbers.Width = 350
    picArtist.Width = 2470
    picArtist.left = 360
    picScroll2.left = 2855
    picTitle.left = 2860
    picTitle.Width = 2000
    picScroll3.left = 1920
    lstPlaylist.ColumnHeaders(1).Width = 349.79
    lstPlaylist.ColumnHeaders(2).Width = 2500.15
    lstPlaylist.ColumnHeaders(3).Width = 2000.12
   
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        FormDrag Me
    End If
    
End Sub

Private Sub lblTime_Click()

    'Changes the time mode
    If TimeMode = "Remaining" Then
        TimeMode = "Current"
        If lblTime.Caption = "R00:00:00.00" Then
            lblTime.Caption = "C00:00:00.00"
        End If
      Else
        TimeMode = "Remaining"
        If lblTime.Caption = "C00:00:00.00" Then
            lblTime.Caption = "R00:00:00.00"
        End If
    End If
    
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Changes the tooltiptext according to the timemode
    If TimeMode = "Current" Then
        lblTime.ToolTipText = "Current Time"
    Else
        lblTime.ToolTipText = "Remaining Time"
    End If

End Sub

Private Sub lstPlaylist_Click()

    'Changes the ID3 filename
    On Error Resume Next
    If lstFilenames.ListCount = 0 Then Exit Sub
    ID3FileName = lstFilenames.List(lstPlaylist.SelectedItem - 1)
    On Error GoTo 0

End Sub

Private Sub lstPlaylist_DblClick()

    'Loads the selected filename
    On Error Resume Next
    If lstFilenames.ListCount = 0 Then Exit Sub
    CurrentFilename = lstFilenames.List(lstPlaylist.SelectedItem - 1)
    PlaylistIndex = lstPlaylist.SelectedItem
    
    OpenMedia CurrentFilename, True
    On Error GoTo 0
    
End Sub

Private Sub lstPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)

    'If you press delete it deletes the selected item...
    'And resets the number column to satisfy the new list
    Dim X As Integer
    
    On Error Resume Next 'Starts then error handle
    
    If KeyCode = vbKeyDelete Then
        
        If frmMain.Caption = lstPlaylist.SelectedItem.ListSubItems(2) Then
            StopMedia
            CloseMedia
        End If
        
        'Removes the item
        lstFilenames.RemoveItem CInt(lstPlaylist.SelectedItem.Text) - 1
        lstPlaylist.ListItems.Remove CInt(lstPlaylist.SelectedItem.Text)
    
        'Resets the numbers
        For X = 1 To lstPlaylist.ListItems.Count
            lstPlaylist.ListItems(X).Text = X
        Next X
        
        'If it's an M3U playlist then it automatically saves
        If PlaylistFilename <> "" Then
            SavePlaylist PlaylistFilename
        End If
        
    End If
    
    On Error GoTo 0 'Dont you hate errors, this line ignores them
    
End Sub

Private Sub lstPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Changes the tooltiptext
    On Error Resume Next
    
    tmrMenu.Enabled = True 'Enable the menu timer
    If lstFilenames.ListCount <> 0 Then
        lstPlaylist.ToolTipText = lstPlaylist.SelectedItem.ListSubItems(2).Text & " Selected"
    Else
        lstPlaylist.ToolTipText = ""
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub picBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = True 'Enable the menu timer

End Sub

Public Function ParseFiles(sFilename As String) As Boolean

'This is the main function to pares files into the playlist...
'First it runs them through the ID3 to get all of the info...
'E.G. Title and Artist
Dim ID3 As New MP3Util

    'If it's the right filename
    If LCase$(Right$(sFilename, 3)) = "wma" Or LCase$(Right$(sFilename, 3)) = "wmv" Or LCase$(Right$(sFilename, 3)) = "vob" Or LCase$(Right$(sFilename, 3)) = "rmi" Or LCase$(Right$(sFilename, 3)) = "mid" Or LCase$(Right$(sFilename, 3)) = "enc" Or LCase$(Right$(sFilename, 3)) = "mpv" Or LCase$(Right$(sFilename, 3)) = "mpm" Or LCase$(Right$(sFilename, 3)) = "snd" Or LCase$(Right$(sFilename, 2)) = "qt" Or LCase$(Right$(sFilename, 3)) = "mpa" Or LCase$(Right$(sFilename, 3)) = "mov" Or LCase$(Right$(sFilename, 3)) = "m1v" _
    Or LCase$(Right$(sFilename, 2)) = "au" Or LCase$(Right$(sFilename, 4)) = "aifc" Or LCase$(Right$(sFilename, 4)) = "aiff" Or LCase$(Right$(sFilename, 3)) = "aif" Or LCase$(Right$(sFilename, 3)) = "wav" Or LCase$(Right$(sFilename, 3)) = "mp1" Or LCase$(Right$(sFilename, 3)) = "mp2" Or LCase$(Right$(sFilename, 3)) = "mpe" Or LCase$(Right$(sFilename, 4)) = "mpeg" Or LCase$(Right$(sFilename, 3)) = "dat" Or LCase$(Right$(sFilename, 3)) = "mpg" Or LCase$(Right$(sFilename, 3)) = "mpg" Or LCase$(Right$(sFilename, 3)) = "avi" _
    Or LCase$(Right$(sFilename, 3)) = "wmx" Or LCase$(Right$(sFilename, 3)) = "wax" Or LCase$(Right$(sFilename, 3)) = "wmp" Or LCase$(Right$(sFilename, 3)) = "mp3" Then
        ParseFiles = True
    Else
        ParseFiles = False
    End If
    
    If ParseFiles = True Then
        'Runs it through the ID3
        ID3.filename = sFilename
        ID3.readTag
        lstPlaylist.ListItems.Add lstPlaylist.ListItems.Count + 1, , lstPlaylist.ListItems.Count + 1
        If FileExists(sFilename) = True Then
            If ID3.artist <> "" Then
                lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 1, , ID3.artist
            Else
                lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 1, , "Unknown"
            End If
            lstFilenames.AddItem sFilename
        End If
        
        If ID3.title <> "" Then
            lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 2, , ID3.title
        Else
            lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 2, , GetFileTitle(sFilename)
        End If
    End If
    
End Function

Private Function FileExists(FullFileName As String) As Boolean

    'To make sure that the files actually exist
    
    On Error GoTo MakeF
    Open FullFileName For Input As #1
    Close #1 'Closes the file so we dont get an error
    FileExists = True
    'Very simple, if theres an error then the file does not exist
    
Exit Function

MakeF:
    FileExists = False

Exit Function

End Function

Private Sub SavePlaylist(sFilename As String)
  
  'Saves the playlist
  Dim iFilenum As Integer
  Dim iCnt As Integer
    
    iFilenum = FreeFile
    
    Open sFilename For Output As #iFilenum
    
    For iCnt = 1 To lstFilenames.ListCount
        Print #iFilenum, lstFilenames.List(iCnt - 1)
    Next iCnt
    
    Close #iFilenum

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = True 'Enables the menu timer

End Sub

Private Sub Picture5_DblClick()

    'Resets the column headers if you dbl click the titlebar
    picScroll1.left = 350
    picNumbers.Width = 350
    picArtist.Width = 2470
    picArtist.left = 360
    picScroll2.left = 2855
    picTitle.left = 2860
    picTitle.Width = 2000
    picScroll3.left = 1920
    lstPlaylist.ColumnHeaders(1).Width = 349.79
    lstPlaylist.ColumnHeaders(2).Width = 2500.15
    lstPlaylist.ColumnHeaders(3).Width = 2000.12

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        FormDrag Me
    End If
    
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = True 'Enable the menu timer

End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrMenu.Enabled = True 'Enable the menu timer

End Sub

Private Sub sldMedia_MouseDown()

    tmrMisc.Enabled = False

End Sub

Private Sub sldMedia_MouseUp()

'Changes the position of the media based on the slider
Dim Pos As Long
  
    If FramesPerSecond = "" Then Exit Sub

    Pos = sldMedia.Value * (FramesPerSecond * 2)
    Result = MoveMultimedia(AliasName, Pos)

    If Result = "Success" Then
    
        tmrMisc.Enabled = True
        
    End If
End Sub

Private Sub tmrAdvanced_Timer()
    
    'Updates the advanced info
    Dim Device As String
    Device = typeDevice
    If Device = "" Then Device = "None"
    lblDevice.Caption = " Device Used: " & Device
    lblPercent.Caption = " Percent Completed: " & Percent & "%"
    lblFPS.Caption = " FPS: " & FramesPerSecond
    lblTotalFrames.Caption = " Total Frames: " & TotalFrames
    Status = GetStatusMultimedia(AliasName)
    If Status = "ERROR" Then Status = "No File"
    lblStatus.Caption = " Multimedia Status: " & UCase(Status)
    lblCurrentVideoW.Caption = " Current Video Width: " & GetSize(AliasName, "cx")
    lblCurrentVideoH.Caption = " Current Video Height: " & GetSize(AliasName, "cy")
    lblActualVideoW.Caption = " Actual Video Width: " & ActualWidth
    lblActualVideoH.Caption = " Actual Video Height: " & ActualHeight
    lblDefaultDevice.Caption = " Default Device: " & GetDefaultDevice(Device)

End Sub

Private Sub tmrEndOfFile_Timer()

    'Does all of the effects if the multimedia is at the end...
    'E.G. Repeat one and Repeat all
    If AreMultimediaAtEnd(AliasName, Val(0)) = True Then
        
        If chkRepeat1.Value = True And chkRepeatAll.Value = False Then
            PlayMedia
        End If
        If chkRepeatAll.Value = True Then
            If PlaylistIndex = lstPlaylist.ListItems.Count Then
                AcquireNextFile
            End If
        End If
        If chkRepeatAll.Value = False And chkRepeat1.Value = False Then
            StopMedia
            CloseMedia
            Exit Sub
        End If
        
    End If
            
End Sub

Private Sub tmrMenu_Timer()
    
    'Closes all of the menus
    picMenuAdd.Visible = False
    picMenuRemove.Visible = False
    picMenuSize.Visible = False
    tmrMenu.Enabled = False

End Sub

Public Sub OpenMedia(filename As String, Play As Boolean)
   
    'Used to open the selected media into then
    'Program and run it through the device
    If lstPlaylist.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    StopMedia
    CloseMedia
        
    If LCase$(Right$(filename, 4)) = ".avi" Then
        typeDevice = "AviVideo"
    Else
        typeDevice = "MPEGVideo"
    End If
    
    Result = OpenMultimedia(frmMovieView.hwnd, AliasName, filename, typeDevice)

    If Result = "Success" Then
    
        ActualWidth = GetSize(AliasName, "cx")
        ActualHeight = GetSize(AliasName, "cy")
    
        tmrMisc.Enabled = True
        
        TotalFrames = GetTotalframes(AliasName)
        FramesPerSecond = GetFramesPerSecond(AliasName)
        TotalTime = GetTotalTimeByMS(AliasName) / 1000
        sldMedia.Max = TotalFrames / (FramesPerSecond * 2)
        frmMain.Caption = lstPlaylist.ListItems(PlaylistIndex).ListSubItems(2)

        If Play = True Then
            PlayMedia
        End If
        
        If LCase$(Right$(filename, 4)) = ".avi" Or LCase$(Right$(filename, 4)) = ".mpg" Or LCase$(Right$(filename, 5)) = ".mpeg" Or LCase$(Right$(filename, 4)) = ".mpe" Or LCase$(Right$(filename, 4)) = ".m1v" Or LCase$(Right$(filename, 4)) = ".mp2" Or LCase$(Right$(filename, 5)) = ".mpv2" Or LCase$(Right$(filename, 4)) = ".mpa" Then
            
            ResizeVideo (0)
            Load frmMovieView
            frmMovieView.Show
            
        Else
            
            frmMovieView.Hide
            
        End If
        
    End If


End Sub

Public Sub ResizeVideo(AddedSize As Integer)
    
    'Resizes the video to your choosing
    tmpWidth = ActualWidth * 15
        
    tmpHeight = ActualHeight * 15
    
    frmMovieView.Width = tmpWidth + AddedSize
            
    frmMovieView.Height = tmpHeight + AddedSize
        
    Result = PutMultimedia(frmMovieView.hwnd, AliasName, Val(0), Val(0), Val(0), Val(0))
                
    frmMovieView.Height = frmMovieView.Height + 350
    
End Sub

Public Sub CloseMedia()
    
    'Closes the media
    Result = CloseMultimedia(AliasName)
    
    If Result = "Success" Then

        tmrMisc.Enabled = False
        sldMedia.Value = 0
        tmrEndOfFile.Enabled = False
        
        If TimeMode = "Current" Then
            lblTime.Caption = "C00:00:00.00"
        Else
            lblTime.Caption = "R00:00:00.00"
        End If
        
        frmMain.Caption = "DMP2 BETA 1b"
        
    End If

End Sub

Public Sub PlayMedia()

    'Plays the media
    Result = PlayMultimedia(AliasName, "", "")

    If Result = "Success" Then
    
        tmrEndOfFile.Enabled = True
    
    End If
    
End Sub

Public Sub PauseMedia()

    'Pauses the media
    Result = PauseMultimedia(AliasName)

    If Result = "Success" Then

    End If

End Sub

Public Sub StopMedia()
    
    'Stops the media
    Result = StopMultimedia(AliasName)

    If Result = "Success" Then
    
    End If

End Sub

Public Sub AcquireNextFile()

    'Mentioned before, it selects the next file
    If PlaylistIndex = lstPlaylist.ListItems.Count Then
        StopMedia
        CloseMedia
        PlaylistIndex = 1
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    Else
        StopMedia
        CloseMedia
        PlaylistIndex = PlaylistIndex + 1
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    End If
    
End Sub

Public Sub AcquirePreviousFile()

    'Mentioned before, it selects the previous file
    If PlaylistIndex = 1 Then
        StopMedia
        CloseMedia
        PlaylistIndex = lstPlaylist.ListItems.Count
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    Else
        StopMedia
        CloseMedia
        PlaylistIndex = PlaylistIndex - 1
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    End If
    
End Sub

Private Sub tmrMisc_Timer()

    'Updates all of the other info...
    'Like the percent and the slider.
    Percent = GetPercent(AliasName)
    If Not Percent = -1 Then sldMedia.Value = Percent * sldMedia.Max \ 100
    CurrentPosition = GetCurrentMultimediaPos(AliasName)
    CurrentTime = Format(Val(CurrentPosition) / Val(FramesPerSecond), "00.000")
    If TimeMode = "Current" Then
        lblTime.Caption = "C" & TimeToString(Format(Val(CurrentPosition) / Val(FramesPerSecond), "00.000"))
      Else
        lblTime.Caption = "R" & TimeToString(Val(Format(Val(TotalTime) - Val(CurrentPosition) / Val(FramesPerSecond), "00.000")))
    End If

End Sub

'This will convert a given ammount of seconds into...
'Hours : Minutes : Seconds : and Milliseconds
Private Function TimeToString(CurrTime As Single) As String

  Dim sMinutes As String
  Dim sSeconds As String
  Dim sMilliseconds As String
  Dim sHours As String
  Dim iMinutes As Integer
  Dim iSeconds As Integer
  Dim iHours As Integer

    iHours = Int(CurrTime / 3600)
    iMinutes = Int((CurrTime - iHours * 3600) / 60)
    iSeconds = Int(CurrTime - iHours * 3600 - iMinutes * 60)
    sHours = Format$(Str(iHours), "00")
    TimeToString = sHours & ":"
    sMinutes = Format$(Str(Int(iMinutes)), "00")
    sSeconds = Format$(Str(Int(iSeconds)), "00")
    sMilliseconds = Format$(Str(Int((CurrTime - iHours * 3600 - iMinutes * 60 - iSeconds) * 100)), "00")
    TimeToString = TimeToString & sMinutes & ":" & sSeconds & "." & sMilliseconds

End Function

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    
    'Makes the form always on top of other programs and forms
    Dim lFlag As String
    
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.left / Screen.TwipsPerPixelX, _
    myfrm.top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
End Sub

Public Function GetFileTitle(ByVal sFilename As String) As String

'This will get the file title, without the filename...
'E.G. "Test.mp3" will be just "Test"
Dim lPos As Long

    lPos = InStrRev(sFilename, "\")
    
    If lPos > 0 Then

        If lPos < Len(sFilename) Then
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            GetFileTitle = ""
        End If
      Else
        GetFileTitle = sFilename
    End If
    
End Function


Public Sub ResumeMedia()
    
    'This will resume a paused media
    Result = ResumeMultimedia(AliasName)

    If Result = "Success" Then

    End If
    
End Sub

'All of the picscroll subs are for my graphical replacement...
'For the listview control's Column Headers

'DO NOT CHANGE ANYTHING BELOW THIS LINE!
'OR THE COLUMN HEADERS WILL NOT WORK CORRECTLY!

Private Sub picScroll1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim OldX
    OldX = X

End Sub

Private Sub picScroll1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Very Complicated
    Dim Drag, OldX, min, NewPos
    
    On Error Resume Next
    
        min = 200
        NewPos = picScroll1.left + X - OldX
        
        If Button = 1 Then
            If NewPos <= min Then
                picScroll1.left = min
                picNumbers.Width = min
                picScroll2.left = picArtist.Width + picNumbers.Width + 25
                picTitle.left = picArtist.Width + 240
                picArtist.left = min + 10
                lstPlaylist.ColumnHeaders(1).Width = min - 10
            Else
                Drag = (X - OldX)
                picScroll1.left = picScroll1.left + (Drag)
                picScroll2.left = picArtist.Width + picNumbers.Width + 25
                picNumbers.Width = picNumbers.Width + (Drag)
                picArtist.left = picArtist.left + (Drag)
                picTitle.left = picTitle.left + (Drag)
                lstPlaylist.ColumnHeaders(1).Width = lstPlaylist.ColumnHeaders(1).Width + (Drag)
                lstPlaylist.ColumnHeaders(2).left = lstPlaylist.ColumnHeaders(2).left + (Drag)
            End If
        End If
        OldX = X
    On Error GoTo 0 'We dont want errors
    
End Sub

Private Sub picScroll2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim OldX
    OldX = X
    
End Sub

'Wow, almost done!
Private Sub picScroll2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Drag, OldX, min, NewPos
    
    On Error Resume Next
        
        DoEvents
        
        min = picNumbers.Width + 530
        NewPos = picScroll2.left + X - OldX
        
        If Button = 1 Then
            If NewPos <= min Then
                picScroll2.left = min - 10
                picArtist.Width = 500
                picTitle.left = min + 10
                lstPlaylist.ColumnHeaders(2).Width = 520
                lstPlaylist.ColumnHeaders(2).left = min
            Else
                Drag = (X - OldX)
                picScroll2.left = picScroll2.left + (Drag)
                picArtist.Width = picArtist.Width + (Drag)
                picTitle.left = picTitle.left + (Drag)
                lstPlaylist.ColumnHeaders(2).Width = lstPlaylist.ColumnHeaders(2).Width + (Drag)
            End If
        End If
        OldX = X
    On Error GoTo 0
    
End Sub

Private Sub picScroll3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim OldX
    OldX = X
    
End Sub

Private Sub picScroll3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Drag, OldX, min, NewPos
    
    On Error Resume Next
        
        DoEvents
        
        min = Label4.Width + 150
        NewPos = picScroll3.left + X - OldX
        
        If Button = 1 Then
            If NewPos <= min Then
                picTitle.Width = min
                picScroll3.left = min - 80
                lstPlaylist.ColumnHeaders(3).Width = min
            Else
                Drag = (X - OldX)
                picScroll3.left = picScroll3.left + (Drag)
                picTitle.Width = picTitle.Width + (Drag)
                lstPlaylist.ColumnHeaders(3).Width = lstPlaylist.ColumnHeaders(3).Width + (Drag)
            End If
        End If
        OldX = X
    On Error GoTo 0

End Sub

'There you have it, probably the best
'Multimedia program on PSC with the
'Help of Abdullah Al-Ahdal. Please vote,
'Because I really like to hear what people
'Think about DMP2. Thanks for trying it
'Out, I hope you liked it.  Remember,
'This program will not be open source
'For very much longer, so if you want
'The DMP2 code you need to dl it now.
'I am going to set up a website for DMP2
'And will be distributing the program
'With an installer.  And after that I will
'Remove the files from PSC, sorry.

'Sincerely,
'Gabriel Loewen (AKA Paranoid Android)
'Dynamic Media Inc.
