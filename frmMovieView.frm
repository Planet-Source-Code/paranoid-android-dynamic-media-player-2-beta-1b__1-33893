VERSION 5.00
Begin VB.Form frmMovieView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMP2 BETA 1b - Video"
   ClientHeight    =   2685
   ClientLeft      =   5145
   ClientTop       =   4305
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmMovieView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrKey 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMovieView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Const AliasName As String = "DMPMedia"
Private Result As String

Private Sub tmrKey_Timer()

    If frmMain.cmdFullscreen.Value = True Then
        If GetAsyncKeyState(vbKeyEscape) Then
            frmMain.cmdFullscreen.Value = False
            frmMovieView.Caption = "DMP2 BETA 1b - Video"
            frmMain.AlwaysOnTop frmMovieView, False
            frmMovieView.Width = (frmMain.ActualWidth * 15) + frmMain.AddedSize
            frmMovieView.Height = (frmMain.ActualHeight * 15) + frmMain.AddedSize
            Result = PutMultimedia(frmMovieView.hwnd, AliasName, Val(0), Val(0), Val(0), Val(0))
            frmMovieView.Height = frmMovieView.Height + 350
            frmMovieView.left = frmMain.left
            frmMovieView.top = frmMain.top
            ShowCursor (True)
        End If
    End If
        
End Sub
