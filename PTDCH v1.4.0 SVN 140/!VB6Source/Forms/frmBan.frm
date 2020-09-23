VERSION 5.00
Begin VB.Form frmBanTemp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ban Temporary IP"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtReason 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      ToolTipText     =   "Enter the reason why you're banning the name (optional)"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      Index           =   2
      ItemData        =   "frmBan.frx":0000
      Left            =   480
      List            =   "frmBan.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      Index           =   1
      ItemData        =   "frmBan.frx":0004
      Left            =   2400
      List            =   "frmBan.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      Index           =   0
      ItemData        =   "frmBan.frx":0008
      Left            =   1440
      List            =   "frmBan.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the reason why you're banning the name (optional)"
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Day(s):"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Minute(s):"
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Hour(s):"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the time to ban the IP."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the time to ban the IP."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmBanTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
1:  Unload Me
End Sub

Private Sub cmdOK_Click()
1:     Dim ConvTime As Long
   
3:   On Error GoTo Err
   
5:   txtIP.Text = Replace(txtIP.Text, " ", "")
   
7:   If ValidIP(txtIP) Then

        'convert time to long minutes
10:     ConvTime = (Val(cmbTime(2).Text) * 24 * 60) + _
                   (Val(cmbTime(0).Text) * 60) + _
                    Val(cmbTime(1).Text)
  
14:     g_colIPBans.Add txtIP.Text, Val(ConvTime), "", "Admin / GUI", txtReason.Text
15:     Unload Me
16:   Else
17:     MsgBoxCenter Me, """" & txtIP.Text & """" & g_colMessages.Item("msgIPNotValide"), vbInformation
18:   End If
   
20:  Exit Sub

22:
Err:
23:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanTemp.cmdOK_Click()"
End Sub

Private Sub Form_Load()
1:    Dim i As Integer
    
3:    On Error GoTo Err
       
5:    For i = 0 To 23
6:       cmbTime(0).AddItem StrZero(i, 2)
7:    Next i
8:    For i = 1 To 59
9:       cmbTime(1).AddItem StrZero(i, 2)
10:    Next i
11:    For i = 0 To 30
12:       cmbTime(2).AddItem StrZero(i, 2)
13:    Next i
    
15:    cmbTime(0).Text = "00"
16:    cmbTime(1).Text = "01"
17:    cmbTime(2).Text = "00"
    
19:    Me.Caption = g_colMessages.Item("msgBanTempIP")
20:    Labels(0).Caption = g_colMessages.Item("msgEnterBanLength")
21:    Labels(1).Caption = g_colMessages.Item("msgEnterBanLength")
22:    Labels(2).Caption = g_colMessages.Item("msgDays")
23:    Labels(3).Caption = g_colMessages.Item("msgHours")
24:    Labels(4).Caption = g_colMessages.Item("msgMinutes")
25:    Labels(5).Caption = g_colMessages.Item("msgEnterBanReason")
26:    cmdClose.Caption = g_colMessages.Item("msgClose")
27:    cmdOK.Caption = g_colMessages.Item("msgOK")
    
29:  Exit Sub

31:
Err:
32:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanTemp.Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
        PaintTileFormBackground Me, LoadImage(g_objSettings.lngSkin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:   Set frmBanTemp = Nothing
End Sub

Private Sub Labels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub

Private Sub txtIP_Change()
1:   If txtIP.Text = "" Then _
          cmdOK.Enabled = False _
     Else cmdOK.Enabled = True
End Sub

Private Sub txtIP_GotFocus()
1:   txtIP.SelStart = 0
2:   txtIP.SelLength = Len(txtIP)
End Sub
