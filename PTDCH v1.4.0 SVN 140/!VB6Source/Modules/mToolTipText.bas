Attribute VB_Name = "mToolTipText"
'Thanks to Jim Jose
Option Explicit

'[APIs]
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

'[Types]
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   bottom   As Long
End Type

Private Type TOOLINFO
    lSize   As Long
    lFlags  As Long
    lHwnd   As Long
    lId     As Long
    lpRect  As RECT
    hInst   As Long
    lpStr   As Long
    lParam  As Long
End Type

'[Enums]
Public Enum ToolTipStyleEnum
    [Tip_Normal] = 0
    [Tip_Balloon] = 1
End Enum

Public Enum ToolTipTypeEnum
    [Tip_None] = 0
    [Tip_Info] = 1
    [Tip_Warning] = 2
    [Tip_Error] = 3
End Enum

'[Local variables]
Private m_MousePos    As POINTAPI
Private m_ToolTipHwnd As Long
Private m_ToolTipInfo As TOOLINFO

'[Required constants]
Private Const WM_USER               As Long = &H400
Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOSIZE            As Long = &H1
Private Const TTS_BALLOON           As Long = &H40
Private Const HWND_TOPMOST          As Long = -&H1
Private Const TTF_SUBCLASS          As Long = &H10
Private Const TTS_NOPREFIX          As Long = &H2
Private Const TTM_DELTOOLW          As Long = (WM_USER + 51)
Private Const TTM_ADDTOOLW          As Long = (WM_USER + 50)
Private Const TTM_SETTITLEW         As Long = (WM_USER + 33)
Private Const TTS_ALWAYSTIP         As Long = &H1
Private Const CW_USEDEFAULT         As Long = &H80000000
Private Const SWP_NOACTIVATE        As Long = &H10
Private Const TOOLTIPS_CLASSA       As String = "tooltips_class32"
Private Const TTM_SETTIPBKCOLOR     As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR   As Long = (WM_USER + 20)

Public Sub ShowToolTip(ByVal hwnd As Long, ByVal mToolTipText As String, ByVal mToolTipHead As String, Optional ByVal mToolTipStyle As ToolTipStyleEnum = Tip_Balloon, Optional ByVal mToolTipType As ToolTipTypeEnum = Tip_None, Optional ByVal mBackColor As Long = -1, Optional ByVal mTextColor As Long = -1)
1:    Dim lpRect As RECT
2:    Dim lWinStyle As Long
3:    Dim MousePos As POINTAPI
4:    On Error GoTo Err
    
       'Get the cursor Position
7:     GetCursorPos MousePos
8:     If m_MousePos.X = MousePos.X And m_MousePos.Y = MousePos.Y Then Exit Sub

       'Remove previous ToolTip
11:    RemoveToolTip
12:    If mToolTipText = vbNullString Then Exit Sub
    
       'Create baloon style if desired
15:    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
16:    If mToolTipStyle = Tip_Balloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    
       'Create the tooltip window
19:    m_ToolTipHwnd = CreateWindowEx(0&, _
                                TOOLTIPS_CLASSA, _
                                vbNullString, _
                                lWinStyle, _
                                CW_USEDEFAULT, _
                                CW_USEDEFAULT, _
                                CW_USEDEFAULT, _
                                CW_USEDEFAULT, _
                                hwnd, 0&, _
                                App.hInstance, 0&)
                
       'Make our tooltip window a topmost window
31:    SetWindowPos m_ToolTipHwnd, HWND_TOPMOST, _
                                0&, 0&, _
                                0&, 0&, _
                                SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    
       'Get the rect of the parent control
37:    GetClientRect hwnd, lpRect
    
       'Now set our tooltip info structure
40:    With m_ToolTipInfo
41:        .lSize = Len(m_ToolTipInfo)
42:        .lFlags = TTF_SUBCLASS
43:        .lHwnd = hwnd
44:        .lId = 0
45:        .hInst = App.hInstance
46:        .lpStr = StrPtr(mToolTipText)
47:        .lpRect = lpRect
48:    End With
    
       'Add the tooltip structure
51:    SendMessage m_ToolTipHwnd, TTM_ADDTOOLW, 0&, m_ToolTipInfo

       'Add TextColor + backColor + Icon
54:    If Not mTextColor = -1 Then SendMessage m_ToolTipHwnd, TTM_SETTIPTEXTCOLOR, mTextColor, 0&
55:    If Not mBackColor = -1 Then SendMessage m_ToolTipHwnd, TTM_SETTIPBKCOLOR, mBackColor, 0&
56:    If Not mToolTipHead = vbNullString Then SendMessage m_ToolTipHwnd, TTM_SETTITLEW, mToolTipType, ByVal StrPtr(mToolTipHead)
    
       'Loop to track Mousemove
59:    Do
60:        m_MousePos.X = MousePos.X: m_MousePos.Y = MousePos.Y
61:        GetCursorPos MousePos
62:        If Not m_MousePos.X = MousePos.X Or Not m_MousePos.Y = MousePos.Y Then
63:            RemoveToolTip
64:            Exit Do
65:        End If
66:        DoEvents
67:    Loop
    
69:    Exit Sub
70:
Err:
71:    HandleError Err.Number, Err.Description, Erl & "|" & "mToolTipText.ShowTips()"
End Sub

'[Important. If not included, tooltips don't change when you try to set the toltip text]
Private Sub RemoveToolTip()
1:    On Error GoTo Err
2:    If m_ToolTipHwnd <> 0 Then
3:        Call SendMessage(m_ToolTipInfo.lHwnd, TTM_DELTOOLW, 0, m_ToolTipInfo)
4:        DestroyWindow m_ToolTipHwnd
5:        m_ToolTipHwnd = 0
6:    End If
7:    Exit Sub
8:
Err:
9:    HandleError Err.Number, Err.Description, Erl & "|" & "mToolTipText.RemoveToolTip()"
End Sub

'[OleColor code to Long color conversion]
Public Function TranslateColor(ByVal lcolor As Long) As Long
1:    On Error GoTo Err
2:    If OleTranslateColor(lcolor, 0, TranslateColor) Then
3:        TranslateColor = -1
4:    End If
5:    Exit Function
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "mToolTipText.ShowTips()"
End Function

