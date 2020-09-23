VERSION 5.00
Begin VB.Form frmDebugSC 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Scripts - Debug Windows"
   ClientHeight    =   1920
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDebug 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu mnuClipboard 
      Caption         =   "Clipboard"
   End
End
Attribute VB_Name = "frmDebugSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sLastLine As String
Private m_bDateTime As Boolean

Public Sub Send(ByVal sText As String)
1:    On Error GoTo Err
2:    If m_bDateTime Then
3:        sText = "[" & Now & "]" & sText
4:    End If
5:    txtDebug.Text = txtDebug.Text & sText & vbNewLine
6:    txtDebug.SelStart = Len(txtDebug.Text)
7:    m_sLastLine = sText
8:    Exit Sub
9:
Err:
11:   HandleError Err.Number, Err.Description, Erl & "|" & "frmDebugSC.Send(" & sText & ")"
End Sub

Public Sub Clear()
1:    txtDebug.Text = ""
End Sub

Public Property Get GetLastLine() As String
1:    On Error Resume Next
2:    GetLastLine = m_sLastLine
End Property

Public Property Let DateTime(ByVal bDateTime As Boolean)
1:    On Error Resume Next
2:    DateTime = m_bDateTime
End Property

Public Sub SetPropertys(Optional ByVal iBackColor As Integer = 1, _
                        Optional ByVal iForeColor As Integer = 0, _
                        Optional ByVal sFontName As String = "Courier New", _
                        Optional ByVal iFontSize As Integer = 8, _
                        Optional ByVal bDateTime As Boolean = False)
5:    On Error GoTo Err
6:    With txtDebug
7:        .BackColor = iBackColor
8:        .ForeColor = iForeColor
9:        .FontName = sFontName
10:       .FontSize = iFontSize
11:   End With
12:   m_bDateTime = bDateTime
13:   Me.Refresh
14:   Exit Sub
15:
Err:
16:   HandleError Err.Number, Err.Description, Erl & "|" & "frmDebugSC.SetPropertys("
End Sub

Private Sub Form_Load()
1:    txtDebug.Left = 0
2:    txtDebug.Top = 0
End Sub

Private Sub Form_Resize()
1:    On Error Resume Next
2:    txtDebug.Width = Me.ScaleWidth
3:    txtDebug.Height = Me.ScaleHeight
End Sub

Private Sub mnuClear_Click()
1:    Call Clear
End Sub

Private Sub mnuClipboard_Click()
1:    If txtDebug.Text = "" Then
2:        Clipboard.Clear
3:        Clipboard.SetText txtDebug.Text
4:    End If
End Sub
