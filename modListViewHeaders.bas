Attribute VB_Name = "modListViewHeaders"
Public Const HDS_BUTTONS As Long = &H2
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Public Const GWL_STYLE As Long = (-16)
Public Const SWP_DRAWFRAME As Long = &H20
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_FLAGS As Long = SWP_NOZORDER Or _
                                 SWP_NOSIZE Or _
                                 SWP_NOMOVE Or _
                                 SWP_DRAWFRAME
  
Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long


