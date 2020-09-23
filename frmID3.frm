VERSION 5.00
Begin VB.Form frmID3 
   BackColor       =   &H009F5000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Editor"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmID3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGenreCode 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   3735
   End
   Begin VB.ComboBox GenreCombo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmID3.frx":000C
      Left            =   960
      List            =   "frmID3.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtFilename 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox txtAlbum 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtArtist 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
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
      Left            =   3000
      TabIndex        =   18
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File size:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Album:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artist:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ID3 As New MP3Util

Private Sub cmdClose_Click()

    frmMain.Enabled = True
    Unload Me

End Sub

Private Sub cmdSave_Click()

    On Error Resume Next

    ID3.title = txtTitle.Text
    ID3.artist = txtArtist.Text
    ID3.year = txtYear.Text
    ID3.album = txtAlbum.Text
    ID3.comment = txtComments.Text
    ID3.Genre = GenreCombo.ListIndex
    ID3.filename = frmMain.ID3FileName
    ID3.writeTag
    
    frmMain.lstPlaylist.ListItems(frmMain.lstPlaylist.SelectedItem.Index).ListSubItems.Item(1).Text = ID3.artist
    frmMain.lstPlaylist.ListItems(frmMain.lstPlaylist.SelectedItem.Index).ListSubItems.Item(2).Text = ID3.title
    cmdClose_Click
    
    On Error GoTo 0

End Sub

Private Sub Form_Load()

On Error Resume Next
Dim X As Integer
    ID3.filename = frmMain.ID3FileName
    ID3.readTag
    
    For X = 0 To 80
        GenreCombo.AddItem ID3.genreDescription(X)
    Next X
    
    txtTitle.Text = ID3.title
    txtArtist.Text = ID3.artist
    txtYear.Text = ID3.year
    txtAlbum.Text = ID3.album
    txtComments.Text = ID3.comment
    GenreCombo.ListIndex = ID3.Genre
    txtGenreCode.Text = ID3.Genre
    txtFilename = ID3.filename
    txtSize.Text = Format(FileLen(frmMain.ID3FileName) / 1048576, "0.00") & " MB"
On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.Enabled = True
    
End Sub

Private Sub GenreCombo_Click()

    txtGenreCode.Text = ID3.genreCode(GenreCombo.List(GenreCombo.ListIndex))
 
End Sub
