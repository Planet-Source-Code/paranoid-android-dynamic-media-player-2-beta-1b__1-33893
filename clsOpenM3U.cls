VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpenM3U"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FileDataType
    Title       As String
    FullName    As String
    Length      As Long
End Type

Private mvarFilename As String
Private mvarFilePath As String
Private mvarCount As Long

Private Const ArraySize = 250
Private m_File() As FileDataType

Private Sub Class_Initialize()

    mvarCount = 0

End Sub

Public Property Get Count() As Long

    Count = mvarCount

End Property

Public Property Get File(Index As Long) As String

    If Index >= 0 And Index < mvarCount Then
        File = m_File(Index).FullName
    End If

End Property

Public Property Get Filename() As String

    Filename = mvarFilePath & "\" & mvarFilename

End Property

Public Property Let Filename(ByVal vData As String)

  Dim Path$, Filename$, Pos&

    Pos& = InStrRev(vData, "\")
    If Pos <> 0 Then
        Path$ = Left(vData, Pos - 1)
        Filename$ = Mid$(vData, Pos + 1)
      Else
        Path$ = ""
        Filename$ = vData
    End If

    mvarFilename = Filename$
    mvarFilePath = Path$

End Property

Public Property Get FileSeconds(Index As Long) As Long

    If Index >= 0 And Index < mvarCount Then
        FileSeconds = m_File(Index).Length
    End If

End Property

Public Property Get FileTitle(Index As Long) As String

    If Index >= 0 And Index < mvarCount Then
        FileTitle = m_File(Index).Title
    End If

End Property

Public Function Refresh(Optional ByVal Filename As String) As Long

    ScanM3U

    Refresh = mvarCount

End Function

Private Sub ScanM3U()

  Dim FilePath$, tmpString$, i%, FindComma%, FindColon%
  Dim MyText$(), InputCount&, FileIn%

    mvarCount = 0
    FileIn% = FreeFile
    FilePath$ = mvarFilePath

    On Error Resume Next
      Open mvarFilePath & "\" & mvarFilename For Input As #FileIn%
      InputCount = -1
      Do While Not (EOF(FileIn%))
          InputCount = InputCount + 1
          ReDim Preserve MyText(InputCount)
          Line Input #FileIn%, MyText(InputCount)
          If UCase$(Left(MyText(InputCount), 7)) = "#EXTM3U" Then
              InputCount = InputCount - 1
          End If
      Loop

      Close FileIn%

      ReDim m_File(0 To ArraySize - 1) As FileDataType

      For i% = 0 To InputCount
          If UCase$(Left(MyText(i%), 7)) = "#EXTINF" Then
              FindComma% = InStr(1, MyText(i%), ",")
              FindColon% = InStr(1, MyText(i%), ":")
              m_File(mvarCount).Length = Val(Mid$(MyText(i%), FindColon + 1, (FindComma - FindColon) - 1))
              m_File(mvarCount).Title = Mid$(MyText(i%), FindComma% + 1)
            
              If Mid$(MyText(i% + 1), 2, 1) <> ":" Then
                  m_File(mvarCount).FullName = FilePath$ & "\" & MyText(i% + 1)
                Else
                  m_File(mvarCount).FullName = MyText(i% + 1)
              End If
              i% = i% + 1
            Else
              If Mid$(MyText(i%), 2, 1) <> ":" Then
                  m_File(mvarCount).FullName = FilePath$ & "\" & MyText(i%)
                Else
                  m_File(mvarCount).FullName = MyText(i%)
              End If
              m_File(mvarCount).Title = MyText(i%)
          End If
          mvarCount = mvarCount + 1
          If mvarCount Mod ArraySize = 0 Then
              ReDim Preserve m_File(0 To mvarCount + ArraySize - 1) As FileDataType
          End If
      Next i
    
End Sub


