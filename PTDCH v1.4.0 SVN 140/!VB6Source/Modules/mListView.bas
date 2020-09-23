Attribute VB_Name = "mListView"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Const HDS_BUTTONS = &H2
Private Const GWL_STYLE = (-16)
Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETHEADER = (LVM_FIRST + 31)
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Private Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Sub LVFullRow(lngHwnd As Long)
1:    Dim rStyle As Long
2:    Dim R As Long
3:    On Error GoTo Err

      'get the current ListView style
6:    rStyle = SendMessageLong(lngHwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
      'set the extended bit
9:    rStyle = rStyle Or LVS_EX_FULLROWSELECT
    
      'set the new ListView style
12:   R = SendMessageLong(lngHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)

14:   Exit Sub
15:
Err:
16:   HandleError Err.Number, Err.Description, Erl & "|" & "mListView.LVFullRow()"
End Sub

Public Sub LVFlatHeaders(lTreeViewhWnd As Long)
1:    Dim lS      As Long
2:    Dim lHwnd   As Long
3:    On Error GoTo Err

      'Set the Buttons mode of the ListView's header control:
6:    lHwnd = SendMessageLong(lTreeViewhWnd, LVM_GETHEADER, 0, 0)
    
8:    If (lHwnd <> 0) Then
9:        lS = GetWindowLong(lHwnd, GWL_STYLE)
           '// change this to lS = lS Or HDS_BUTTONS to make them normal again
11:       lS = lS And Not HDS_BUTTONS
12:       SetWindowLong lHwnd, GWL_STYLE, lS
13:   End If
    
15:   Exit Sub
16:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "mListView.LVFlatHeaders()"
End Sub
