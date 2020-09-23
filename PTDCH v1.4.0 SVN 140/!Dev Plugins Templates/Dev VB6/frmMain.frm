VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DevTest"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkEnabledPlg 
      Caption         =   "Check1"
      Height          =   195
      Left            =   5400
      TabIndex        =   1
      Top             =   2640
      Width           =   195
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblHolder 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled Plugin"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEnabledPlg_Click()
  If chkEnabledPlg.Value = vbChecked Then _
         g_Main.RunEvent "Switch", True _
    Else g_Main.RunEvent "Switch", False
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    If g_Enabled Then _
         chkEnabledPlg.Value = vbChecked _
    Else chkEnabledPlg.Value = vbUnchecked
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If g_objSettings.MoveForm Then _
        frmHub.RunFunction "frmMove", Me
End Sub

Private Sub Form_Paint()
    If g_objSettings.blSkin Then _
        frmHub.RunFunction "PaintTileFormBackground", Me
End Sub

Private Sub lblHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmHub.RunFunction "ShowToolTip", Me.hWnd, "Enabled or Desabled this plugin.", g_Main.Name, 1, 1
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmHub.RunFunction "ShowToolTip", cmdOK.hWnd, "Click to hide the plugin Form", g_Main.Name, 0, 1
End Sub
