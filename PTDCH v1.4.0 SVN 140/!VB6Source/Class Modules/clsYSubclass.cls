VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Sub AttachMessage(iwp As clsYISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
    On Error GoTo Err
    mYSubclass.AttachMessage iwp, hWnd, iMsg
    Exit Sub

Err:
    HandleError Err.Number, Err.Description, Erl & "|clsYSubclass.AttachMessage()"
End Sub

Sub DetachMessage(iwp As clsYISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
    On Error GoTo Err
    mYSubclass.DetachMessage iwp, hWnd, iMsg
    Exit Sub

Err:
    HandleError Err.Number, Err.Description, Erl & "|clsYSubclass.DetachMessage()"
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
1:    On Error GoTo Err
2:   CurrentMessage = mYSubclass.CurrentMessage
3:    Exit Property

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|clsYSubclass.CurrentMessage()"
End Property

Public Function CallOldWindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
6:    On Error GoTo Err
7:   CallOldWindowProc = mYSubclass.CallOldWindowProc(hWnd, iMsg, wParam, lParam)
8:    Exit Function

10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|clsYSubclass.CallOldWindowProc()"
End Function

