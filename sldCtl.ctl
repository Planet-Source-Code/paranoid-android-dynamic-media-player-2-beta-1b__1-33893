VERSION 5.00
Begin VB.UserControl SliderControl 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E37200&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   ScaleHeight     =   10
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   69
   Begin VB.Shape myShape 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00843800&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "SliderControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************
'       SimpleColorSlider with Float Values        *
'          Written by Andrew Stopakevich           *
'            Modified by Kenneth Hedman            *
'***************************************************

'Default Property Values:
Option Explicit
Const m_def_Min = 0 ':( As 16-bit Integer ?':( Missing Scope
Const m_def_Max = 100 ':( As 16-bit Integer ?':( Missing Scope
Const m_def_Value = 0 ':( As 16-bit Integer ?':( Missing Scope
'Property Variables:
Dim m_Min As Double ':( Missing Scope
Dim m_Max As Double ':( Missing Scope
Dim m_Enabled As Boolean ':( Missing Scope
Dim m_Value As Double ':( Missing Scope
Dim m_roundto As Double ':( Missing Scope
'Event Declarations:
Event MouseMove() ':( Missing Scope
Event MouseDown() ':( Missing Scope
Event MouseUp() ':( Missing Scope
Event Click() ':( Missing Scope

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled

End Property

Public Property Let Enabled(New_Value As Boolean)

    m_Enabled = New_Value
    RefreshMe
    PropertyChanged "Enabled"

End Property

Public Property Get Max() As Double

    Max = m_Max

End Property

Public Property Let Max(ByVal New_Max As Double)

    m_Max = New_Max
    PropertyChanged "Max"
    RefreshMe

End Property

'Min, Max
Public Property Get Min() As Double

    Min = m_Min

End Property

Public Property Let Min(ByVal New_Min As Double)

    m_Min = New_Min
    PropertyChanged "Min"
    RefreshMe

End Property

'Positions
Public Sub RefreshMe()

    If m_Max = m_Min Then m_Max = m_Max + 1 ':( Expand Structure
    If m_Value < m_Min Then m_Value = m_Min ':( Expand Structure
    If m_Value > m_Max Then m_Value = m_Max ':( Expand Structure
    If m_Enabled = True Then ':( Remove Pleonasm
        UserControl.Enabled = True
      Else 'NOT M_ENABLED...
        UserControl.Enabled = False
    End If
    myShape.Width = UserControl.ScaleWidth * (m_Value - m_Min) / (m_Max - m_Min) + 12
    myShape.Height = UserControl.Height

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click

End Sub

'Initalizing the control
Private Sub UserControl_Initialize()

    Call UserControl_Resize

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Value = m_def_Value
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Enabled = True
    RefreshMe

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X = 0 Then Exit Sub ':( Expand Structure or consider reversing Condition
    m_Value = (X * (m_Max - m_Min) + UserControl.ScaleWidth * m_Min) / UserControl.ScaleWidth
    RefreshMe
    myShape.BorderStyle = 1
    myShape.FillColor = &H80FF&
    RaiseEvent MouseDown
    
End Sub

'Moves the slider
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And X > 0 Then
        m_Value = (X * (m_Max - m_Min) + UserControl.ScaleWidth * m_Min) / UserControl.ScaleWidth
        RefreshMe
        RaiseEvent MouseMove
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X = 0 Then Exit Sub ':( Expand Structure or consider reversing Condition
    m_Value = (X * (m_Max - m_Min) + UserControl.ScaleWidth * m_Min) / UserControl.ScaleWidth
    RefreshMe
    myShape.BorderStyle = 0
    myShape.FillColor = &H843800
    RaiseEvent MouseUp

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Enabled = PropBag.ReadProperty("Enabled", True)

    RefreshMe

End Sub

'Resize sub
Private Sub UserControl_Resize()

    myShape.Height = UserControl.Height + 10
    myShape.left = -10
    myShape.top = -10
    RefreshMe

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)

End Sub

'Values
Public Property Get Value() As Double

    Value = Round(m_Value, m_roundto)

End Property

Public Property Let Value(ByVal New_Value As Double)

    m_Value = New_Value
    RefreshMe
    PropertyChanged "Value"

End Property

':) Ulli's VB Code Formatter V2.8.9 (3/23/02 11:37:57 AM) 22 + 165 = 187 Lines
