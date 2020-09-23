VERSION 5.00
Begin VB.Form frmLoading 
   BorderStyle     =   0  'None
   Caption         =   "PTDCH - Loading.."
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
1:   On Error GoTo Err
2:   Me.Picture = LoadImage(101)
3:   Show
4:   Pause 500
5:  Exit Sub
6:
Err:
7:  MsgBox Err.Description, vbCritical, "PTDCH - frmLoading.Form_Load()"
End Sub
