VERSION 5.00
Begin VB.Form frmChatRoom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat Room"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   22
      Text            =   "15"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CheckBox chkBox 
      Caption         =   "IsOperator"
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   17
      Top             =   2640
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.ComboBox cmbShare 
      Height          =   315
      ItemData        =   "frmChatRom.frx":0000
      Left            =   1680
      List            =   "frmChatRom.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Close"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "frmChatRom.frx":0004
      Left            =   2520
      List            =   "frmChatRom.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   8
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Text            =   "Chatroom"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox chkBox 
      Caption         =   "IsOperator"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Icon"
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   21
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "TAG"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblHolder 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled"
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   18
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Class"
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   13
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Share"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton_Click(Index As Integer)
1:  On Error GoTo Err

  Select Case Index
        Case 0
3:            Dim intClass As Integer
4:            Dim dblShare As Double
            
            Select Case CStr(cmbClass.Text)
                    Case "2 = Mentored": intClass = 2
                    Case "3 = Registered": intClass = 3
                    Case "4 = Invisible": intClass = 4
                    Case "5 = VIP": intClass = 5
                    Case "6 = Operator": intClass = 6
                    Case "7 = Invisible Operator": intClass = 7
                    Case "8 = Super Operator": intClass = 8
                    Case "9 = Invisible Super Operator": intClass = 9
                    Case "10 = Admin": intClass = 10
                    Case "11 = Invisible Admin": intClass = 11
6:              End Select
            
8:           dblShare = Val(txtBox(1).Text)

              Select Case CStr(cmbShare.Text)
                    Case "Bytes": dblShare = dblShare
                    Case "KiB": dblShare = dblShare * 1024
                    Case "MiB": dblShare = dblShare * 1024 * 1024
                    Case "GiB": dblShare = dblShare * 1024 * 1024 * 1024
10:           End Select

12:           g_objChatRoom.AddChat txtBox(0).Text, _
                                    chkBox(1).Value, _
                                    intClass, chkBox(0).Value, _
                                    dblShare, txtBox(2).Text, _
                                    txtBox(3).Text, _
                                    txtBox(4).Text, _
                                    txtBox(5).Text, _
                                    txtBox(6).Text, _
                                    True
21:           Unload Me
        Case 1: Unload Me
22:       End Select

24:    Exit Sub
25:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|" & "frmChatRoom.cmdButton_Click()"
End Sub

Private Sub cmdHelp_Click()
1: MsgBoxCenter Me, "Valid icon type are from caracter 1 to 15:" & vbNewLine & _
            "   " & "1, 4, 5, 8 and 9: AFK = False" & vbNewLine & _
            "   " & "2, 3, 6, 7, 10 and 11: AFK = True" & vbNewLine & _
            "   " & "12, 13, 14 and 15: chat only client", _
            vbOKOnly Or vbInformation
End Sub

Private Sub Form_Load()
1:  On Error GoTo Err

3:    With cmbShare
4:        .AddItem "Bytes"
5:        .AddItem "KiB"
6:        .AddItem "MiB"
7:        .AddItem "GiB"
8:        .Text = "GiB"
9:   End With

13:  cmbClass.Text = "6 = Operator"

15:    With frmHub.lvwChatRom
16:        lblHolder(0).Caption = .ColumnHeaders(1).Text
17:        lblHolder(7).Caption = .ColumnHeaders(2).Text
18:        lblHolder(6).Caption = .ColumnHeaders(3).Text
19:        lblHolder(1).Caption = .ColumnHeaders(4).Text
20:        lblHolder(2).Caption = .ColumnHeaders(5).Text
21:        lblHolder(3).Caption = .ColumnHeaders(6).Text
22:        lblHolder(4).Caption = .ColumnHeaders(7).Text
23:        lblHolder(5).Caption = .ColumnHeaders(8).Text
24:        lblHolder(8).Caption = .ColumnHeaders(9).Text
25:        lblHolder(9).Caption = .ColumnHeaders(10).Text
26:    End With
    
28:    cmdButton(0).Caption = g_colMessages.Item("msgOK")
29:   cmdButton(1).Caption = g_colMessages.Item("msgCancel")

31:    Exit Sub
32:
Err:
33:   HandleError Err.Number, Err.Description, Erl & "|" & "frmChatRoom.Form_Load()"
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
         PaintTileFormBackground Me, LoadImage(g_objSettings.lngSkin)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub

Private Sub lblHolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub

Private Sub txtBox_Change(Index As Integer)
1:  If Index = 1 Or Index = 6 Then _
         If Not IsNumeric(txtBox(Index).Text) Then _
             txtBox(Index).Text = 0
4:  If Index = 0 Then _
         If Not LenB(txtBox(0).Text) Then _
             cmdButton(0).Enabled = True _
        Else cmdButton(0).Enabled = False
8:  If Index = 6 Then _
         If CInt(txtBox(6).Text) > 15 Then _
            txtBox(Index).Text = 0
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
1:  If Index = 1 Or Index = 6 Then _
         If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then _
            KeyAscii = 0
End Sub
