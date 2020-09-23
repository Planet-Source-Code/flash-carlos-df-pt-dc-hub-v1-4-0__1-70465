VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimersCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Public Event Timer(ByVal Index As Integer)

Private m_intCount      As Integer
Private m_colTimers     As Collection 'object ref to child clsTimer classes created by this class
Private m_intClassKey   As Integer

Public Property Get Count() As Integer
1:    Count = m_intCount
End Property

Public Function Add() As Integer
1:    Dim objTimer As New clsTimer
2:    On Error GoTo Err
     
4:    m_intCount = m_intCount + 1 '1 based indexes (same as collection object)
    
      'these two friend properties were added to base class (clsTimer) to support this framework
7:    objTimer.Index = m_intCount
8:    objTimer.ParentsColKey = m_intClassKey
    
10:   m_colTimers.Add objTimer, "key:" & CStr(m_intCount)
    
12:   Add = m_intCount
13:   Exit Function
14:
Err:
15:   HandleError Err.Number, Err.Description, Erl & "|" & "clsTimersCol.Add()"
End Function

Public Function Item(ByVal Index As Integer) As clsTimer
1:    On Error GoTo Err
      'Default function of class (so we can refer to it as clsTimer.Item(key:index)
3:    Set Item = m_colTimers("key:" & CStr(Index))
5:    Exit Function
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimersCol.Item(" & Index & ")"
End Function

Public Sub Remove(ByVal iIndex As Integer)
1:    Dim objTimer As clsTimer
2:    On Error Resume Next
3:    Set objTimer = Item(iIndex)
4:    If Err.Number = 0 Then
5:         m_colTimers.Remove "key:" & CStr(objTimer.Index)
6:         If Err.Number = 0 Then
7:             objTimer.Disponse
8:             m_intCount = (m_intCount - 1)
9:         End If
10:   End If
End Sub

Friend Function RaiseTimer_Event(ByVal Index As Integer)
1:    On Error GoTo Err
      'Called from mTimer directly
3:    RaiseEvent Timer(Index)
4:    Exit Function
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimersCol.RaiseTimer_Event(" & Index & ")"
End Function

Private Sub Class_Initialize()
1:    On Error GoTo Err
2:    Set m_colTimers = New Collection
3:    m_intClassKey = RegisterTimerCollection(Me)
4:    Exit Sub
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimersCol.Class_Initialize()"
End Sub

Private Sub Class_Terminate()
1:    Dim objTimer As clsTimer
2:    On Error GoTo Err
    
4:    For Each objTimer In m_colTimers
5:        Set objTimer = Nothing
6:    Next
8:    Set m_colTimers = Nothing
    
10:   g_colTimerS.Remove "key:" & m_intClassKey
11:   Exit Sub
Err:
14:   HandleError Err.Number, Err.Description, Erl & "|" & "clsTimersCol.Class_Terminate()"
End Sub
