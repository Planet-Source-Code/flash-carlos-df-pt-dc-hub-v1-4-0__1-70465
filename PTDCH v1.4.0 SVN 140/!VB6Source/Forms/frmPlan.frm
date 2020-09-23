VERSION 5.00
Begin VB.Form frmPlan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schelduler"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Close"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   4695
   End
   Begin VB.ComboBox cmbBox 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPlan.frx":0000
      Left            =   2520
      List            =   "frmPlan.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Text            =   "0"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox chkBox 
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   195
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CheckBox chkBox 
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Seperate user with a semicolon (;) (First user in the list is the one used for user object purposes)"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   17
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Increase Type"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Increase Value"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Increase"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Parameter"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Command"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Time"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "User(s) Form"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBox_Click(Index As Integer)
1:    If Index = 1 Then
2:        If chkBox(1).Value = 0 Then
3:            txtBox(4).Enabled = False
4:            cmbBox.Enabled = False
5:            lblHolder(5).Enabled = False
6:            lblHolder(6).Enabled = False
7:        Else
8:            txtBox(4).Enabled = True
9:            cmbBox.Enabled = True
10:            lblHolder(5).Enabled = True
11:            lblHolder(6).Enabled = True
12:        End If
13:    End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
1:    Dim varTemp As Variant
2:    Dim intTemp As Integer
3:    Dim dtTemp As Date
    
    Select Case Index
        Case 0
5:            If Not IsDate(txtBox(1).Text) Then
6:               MsgBoxCenter Me, "Formate date invalid." & vbNewLine & _
                                "Valid formate ex: " & Now, _
                                 vbOKOnly Or vbInformation
9:               Exit Sub
10:            Else
11:               dtTemp = CDate(txtBox(1).Text)
12:            End If
        
14:            If CBool(chkBox(1).Value) Then
                Select Case cmbBox.Text
                    Case "None": varTemp = 0
                    Case "Minute(s)": varTemp = 1
                    Case "Hour(s)": varTemp = 2
                    Case "Day(s)": varTemp = 3
                    Case "Month(s)": varTemp = 4
15:                End Select
16:                intTemp = CInt(txtBox(4).Text)
17:            Else
18:                varTemp = 0
19:                intTemp = 0
20:            End If
            
22:            g_objScheduler.AddPlan txtBox(0).Text, CDate(txtBox(1).Text), CBool(chkBox(0).Value), txtBox(2).Text, txtBox(3).Text, intTemp, CVar(varTemp), txtBox(5).Text
        
        Case 1
        '
25:    End Select
    
27:    Unload Me
    
End Sub

Private Sub Form_Load()
1:  cmbBox.Text = "None"
2:  txtBox(1).Text = Now

4:   With frmHub.lvwPlan
5:        lblHolder(0).Caption = .ColumnHeaders(1).Text
6:        lblHolder(2).Caption = .ColumnHeaders(2).Text
7:        lblHolder(8).Caption = .ColumnHeaders(3).Text
8:        lblHolder(1).Caption = .ColumnHeaders(4).Text
9:        lblHolder(3).Caption = .ColumnHeaders(5).Text
10:       lblHolder(4).Caption = .ColumnHeaders(6).Text
11:       lblHolder(5).Caption = .ColumnHeaders(6).Text
12:       lblHolder(6).Caption = .ColumnHeaders(7).Text
13:       lblHolder(7).Caption = .ColumnHeaders(8).Text
14:  End With

16:   cmdButton(0).Caption = g_colMessages.Item("msgOK")
17:   cmdButton(1).Caption = g_colMessages.Item("msgCancel")
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
1:  If Index = 4 Then _
         If Not IsNumeric(txtBox(4).Text) Then _
             txtBox(4).Text = 0

5:  If txtBox(0).Text <> "" And _
       txtBox(1).Text <> "" And _
       txtBox(2).Text <> "" Then _
         cmdButton(0).Enabled = True _
    Else cmdButton(0).Enabled = False
    
11: txtBox(0).Text = Replace(txtBox(0).Text, " ", "")

13: If Index = 1 Then _
         If txtBox(1).Text = "" Then _
             txtBox(1).Text = Now
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
1:  If Index = 4 Then _
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then _
            KeyAscii = 0
End Sub
