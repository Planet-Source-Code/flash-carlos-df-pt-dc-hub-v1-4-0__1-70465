Attribute VB_Name = "mTbsPicStyle"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private Const cFlatB = True

Private Sub FlatBorder(ByVal hwnd As Long)

2:    If cFlatB Then
3:       On Error GoTo Err
4:       Dim TFlat As Long

6:       TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
7:       TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
8:       SetWindowLong hwnd, GWL_EXSTYLE, TFlat
9:       SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
10:      End If
11:   Exit Sub
12:
Err:
13:   HandleError Err.Number, Err.Description, Erl & "|" & "mTbsPicStyle.FlatBorder()"
End Sub

Public Sub SetFlatBorder()

2:    On Error GoTo Err
3:    Dim i As Integer
   
5:    With frmHub
6:      For i = 0 To .picTab.count - 1
7:       FlatBorder .picTab(i).hwnd
8:         PicBkgToTabStrip .picTab(i), .tbsMenu
9:      Next i
        
11:     For i = 0 To .picSTab.count - 1
12:       FlatBorder .picSTab(i).hwnd
13:          PicBkgToTabStrip .picSTab(i), .tbsSecurity
14:     Next i
     
16:     For i = 0 To .picITab.count - 1
17:       FlatBorder .picITab(i).hwnd
18:           PicBkgToTabStrip .picITab(i), .tbsInteractions
19:     Next i
     
21:     For i = 0 To .picTabAdv.count - 1
22:       FlatBorder .picTabAdv(i).hwnd
23:          PicBkgToTabStrip .picTabAdv(i), .tabAdv
24:     Next i
     
26:     For i = 0 To .picHelp.count - 1
27:       FlatBorder .picHelp(i).hwnd
28:          PicBkgToTabStrip .picHelp(i), .tbsHelp
29:     Next i
     
31:     For i = 0 To .picInfo.count - 1
32:       FlatBorder .picInfo(i).hwnd
33:       PicBkgToTabStrip .picInfo(i), .tbsInfo
34:     Next i

36:     For i = 0 To .picStatus.count - 1
37:       FlatBorder .picStatus(i).hwnd
38:       PicBkgToTabStrip .picStatus(i), .tbsStatus
39:     Next i

41:     FlatBorder .sldPriority.hwnd
42:     FlatBorder .picLog(0).hwnd
43:     FlatBorder .picLog(1).hwnd
44:     FlatBorder .picStInfo.hwnd
45:     FlatBorder .sldStatus.hwnd
46:     FlatBorder .tlbScript.hwnd
47:     FlatBorder .lvwScripts.hwnd
48:     FlatBorder .lvwPlugins.hwnd

50:     LVFullRow .lvwCommands.hwnd
51:     LVFullRow .lvwPermIPBan.hwnd
52:     LVFullRow .lvwTempIPBan.hwnd
53:     LVFullRow .lvwUsers.hwnd
54:     LVFullRow .lvwRegistered.hwnd
55:     LVFullRow .lvwBans.hwnd
56:     LVFullRow .lvwChatRom.hwnd
57:     LVFullRow .lvwPlan.hwnd
58:     LVFullRow .lvwLanguages.hwnd
59:     LVFullRow .lvwSqlExplorer.hwnd

60:     LVFlatHeaders .lvwScripts.hwnd
61:     LVFlatHeaders .lvwPlugins.hwnd
62:   End With

64:  Exit Sub
65:
Err:
66:  HandleError Err.Number, Err.Description, Erl & "|" & "mTbsPicStyle.SetFlatBorder()"
End Sub
' Shape a picturebox Background to a 5.0 Tabstrip. This is
' useful when you are placing a picturebox control container
' on a Tabstrip, and want to be sure that the picturebox will
' fill the tabstrip body.
Private Sub PicBkgToTabStrip(pBackground As PictureBox, TbStrip As Object)
1:  pBackground.Left = TbStrip.Left + 80 '15         'right of left border
2:  pBackground.Width = TbStrip.Width - 170 '60      'keep inside right border
3:  pBackground.Top = TbStrip.Top + 360  '330        'below top border
4:  pBackground.Height = TbStrip.Height - 455 '375   'above bottom border
End Sub

'Repaint tiling
Public Sub PaintTileFormBackground(MyForm As Form, MyPicture As IPictureDisp)
1:  Dim i As Long, j As Long
2:    For i = 0 To MyForm.ScaleWidth Step 1770      'Used original image size .. draw across top
3:      For j = 0 To MyForm.ScaleHeight Step 2070   'Used original image size .. draw across height
4:        MyForm.PaintPicture MyPicture, i, j       'draw a frame
5:      Next j
6:    Next i
End Sub

'Repaint tiling
Public Sub PaintTilePicBackground(Mypic As PictureBox, MyPicture As IPictureDisp)
1:  Dim i As Long, j As Long
2:    For i = 0 To Mypic.ScaleWidth Step 1770     'Used original image size .. draw across top
3:      For j = 0 To Mypic.ScaleHeight Step 2070  'Used original image size .. draw across height
4:        Mypic.PaintPicture MyPicture, i, j       'draw a frame
5:      Next j
6:    Next i
End Sub
