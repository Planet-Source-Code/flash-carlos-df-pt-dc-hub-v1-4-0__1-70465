VERSION 5.00
Begin VB.Form frmBanPerm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ban permanent IP"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtReason 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      ToolTipText     =   "Enter the reason why you're banning the name (optional)"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the reason why you're banning the name (optional)"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the IP to permanet ban."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmBanPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
1:   Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
1:   txtIP.SelStart = 0
2:   txtIP.SelLength = Len(txtIP)
End Sub

Private Sub cmdOK_Click()
1:   On Error GoTo Err

3:   txtIP.Text = Replace(txtIP.Text, " ", "")

5:   If ValidIP(txtIP) Then
6:      g_colIPBans.Add txtIP.Text, -1, "", "Admin / GUI", txtReason.Text
7:      Unload Me
8:   Else
9:      MsgBoxCenter Me, """" & txtIP.Text & """" & g_colMessages.Item("msgIPNotValide"), vbInformation
10:  End If
   
12:  Exit Sub

14:
Err:
15:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanPerm.cmdOK_Click()"
End Sub

Private Sub Form_Load()
1:    On Error GoTo Err
2:    Me.Caption = g_colMessages.Item("msgBanPermIP")
3:    Labels(0).Caption = g_colMessages.Item("msgEnterPermIP")
4:    Labels(1).Caption = g_colMessages.Item("msgEnterBanReason")
5:    cmdCancel.Caption = g_colMessages.Item("msgCancel")
6:    cmdOK.Caption = g_colMessages.Item("msgOK")
7:    Exit Sub
8:
Err:
9:   HandleError Err.Number, Err.Description, Erl & "|" & "frmBanPerm.Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1:  If g_objSettings.blSkin Then _
         PaintTileFormBackground Me, LoadImage(g_objSettings.lngSkin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:   Set frmBanPerm = Nothing
End Sub

Private Sub Label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub

Private Sub txtIP_Change()
1:   If txtIP.Text = "" Then _
          cmdOK.Enabled = False _
     Else cmdOK.Enabled = True
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
1:   If KeyAscii = 13 Then _
         Call cmdOK_Click
End Sub
