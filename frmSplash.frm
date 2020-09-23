VERSION 5.00
Begin VB.Form frmSplashScreen 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   1860
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading DMP2 BETA 1..."
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
      Left            =   120
      TabIndex        =   0
      Top             =   1650
      Width           =   4170
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

    Load frmMain
    frmMain.Show
    Unload Me
    
End Sub

Private Sub tmrLoad_Timer()

    Load frmMain
    frmMain.Show
    Unload Me
    
End Sub
