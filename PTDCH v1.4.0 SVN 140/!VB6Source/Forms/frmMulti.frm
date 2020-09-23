VERSION 5.00
Begin VB.Form frmMulti 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Multi Use"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStr 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtStrMultiLine 
      Height          =   1005
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
1:   txtStrMultiLine.Text = ""
2:   txtStr.Text = ""
3:   Unload Me
End Sub

Private Sub cmdOK_Click()
1:   Me.Hide
End Sub

Private Sub Form_Load()
1:   DoEvents
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
         PaintTileFormBackground Me, LoadImage(g_objSettings.lngSkin)
End Sub

Private Sub Form_Resize()
1:    If txtStrMultiLine.Visible And Me.Height > 2000 And Me.Width > 3000 Then
2:        On Error Resume Next
3:        With txtStrMultiLine
4:            .Left = 120
5:            .Top = 360
6:            .Height = Me.Height - 1300
7:            .Width = Me.Width - 360
8:        End With
9:        With cmdOK
10:            .Left = 120
11:            .Top = txtStrMultiLine.Height + 440
12:        End With
13:        With cmdCancel
14:            .Left = 120
15:            .Top = txtStrMultiLine.Height + 440
16:            .Left = txtStrMultiLine.Width - 1100
17:        End With
18:    End If
End Sub

