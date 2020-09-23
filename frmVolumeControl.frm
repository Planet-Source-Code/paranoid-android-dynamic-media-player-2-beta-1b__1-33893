VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVolumeControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Volume Control"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.OptionButton OptnChannelLeft 
      Caption         =   "Left"
      Height          =   225
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   585
   End
   Begin VB.OptionButton OptnChannelRight 
      Caption         =   "Right"
      Height          =   225
      Left            =   2280
      TabIndex        =   14
      Top             =   360
      Width           =   700
   End
   Begin VB.OptionButton OptnChannelAllOn 
      Caption         =   "All On"
      Height          =   225
      Left            =   2280
      TabIndex        =   13
      Top             =   600
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton OptnChannelAllOff 
      Caption         =   "Mute"
      Height          =   225
      Left            =   2280
      TabIndex        =   12
      Top             =   840
      Width           =   765
   End
   Begin VB.Frame FrameLeftVol 
      Caption         =   "Left"
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   735
      Begin MSComctlLib.Slider sldLeftVol 
         Height          =   1650
         Left            =   45
         TabIndex        =   17
         Top             =   240
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   2910
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   49
         Left            =   330
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   50
         Left            =   390
         TabIndex        =   10
         Top             =   990
         Width           =   300
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   51
         Left            =   420
         TabIndex        =   9
         Top             =   1650
         Width           =   225
      End
   End
   Begin VB.Frame FrameRightVol 
      Caption         =   "Right"
      Height          =   1935
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   735
      Begin MSComctlLib.Slider sldRightVol 
         Height          =   1650
         Left            =   50
         TabIndex        =   18
         Top             =   240
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   2910
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   52
         Left            =   420
         TabIndex        =   7
         Top             =   1650
         Width           =   225
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   53
         Left            =   390
         TabIndex        =   6
         Top             =   990
         Width           =   300
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   54
         Left            =   330
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Both"
      Height          =   1935
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   735
      Begin MSComctlLib.Slider sldBothVol 
         Height          =   1650
         Left            =   30
         TabIndex        =   19
         Top             =   240
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   2910
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   990
         Width           =   300
      End
      Begin VB.Label Lbcaption 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   420
         TabIndex        =   1
         Top             =   1650
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmVolumeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const AliasName As String = "DMPMedia"
Dim Result As String
Dim vol As Long

Private Sub Command1_Click()

    Me.Hide

End Sub

Private Sub OptnChannelAllOff_Click()
    
    Result = ChannelsControl(AliasName, "all", "off")
    
End Sub

Private Sub OptnChannelAllOn_Click()
    
    Result = ChannelsControl(AliasName, "all", "on") 'turn on the BOTH channel(left & right) for this Alias Multimedia
        
    sldBothVol.Value = 0
    sldLeftVol.Value = 0
    sldRightVol.Value = 0
        
    If Result = "Success" Then

        SetVolume AliasName, "all", 100

    End If
    
End Sub

Private Sub OptnChannelLeft_Click()
    
    Result = ChannelsControl(AliasName, "left", "on")
    Result = ChannelsControl(AliasName, "right", "off")

    sldLeftVol.Value = 0
    sldRightVol.Value = 100
    sldBothVol.Value = 50
    
End Sub

Private Sub OptnChannelRight_Click()
    
    Result = ChannelsControl(AliasName, "right", "on")
    Result = ChannelsControl(AliasName, "left", "off")

    sldLeftVol.Value = 100
    sldRightVol.Value = 0
    sldBothVol.Value = 50

End Sub

Private Sub sldBothVol_Click()
    
    vol = (sldBothVol.Value - 100) * -1

    Result = SetVolume(AliasName, "both", vol)

    If Result = "Success" Then

    End If

End Sub

Private Sub sldLeftVol_Scroll()

    vol = (sldLeftVol.Value - 100) * -1

    Result = SetVolume(AliasName, "left", vol)

    If Result = "Success" Then

    End If
    
End Sub

Private Sub sldRightVol_Click()

    vol = (sldRightVol.Value - 100) * -1

    Result = SetVolume(AliasName, "right", vol)

    If Result = "Success" Then

    End If
    
End Sub
