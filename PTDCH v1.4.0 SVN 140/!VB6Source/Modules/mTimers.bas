Attribute VB_Name = "mTimers"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'each timer class registers an obj instance in this collection
'key= "id:" & timerID , item = reference to live class object
Public g_colTimer   As New Collection
'each clsTimersCol class registers itself by its class key here
'key= "key:" & intID , item = reference to live class object
Public g_colTimerS  As New Collection

Private m_TimersCount As Integer

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
1:    Dim objTimer        As clsTimer
2:    Dim objTimerCol     As clsTimersCol
    
4:    On Error Resume Next
5:    Set objTimer = g_colTimer("id:" & idEvent)
6:    On Error GoTo Err
    
8:     If objTimer Is Nothing Then
9:         KillTimer 0&, idEvent
10:    Else
11:       If objTimer.ParentsColKey > 0 Then  'this timer is an index in clsTimersCol
12:            Set objTimerCol = g_colTimerS("key:" & objTimer.ParentsColKey)
13:            If objTimerCol Is Nothing Then
14:                KillTimer 0&, idEvent
                   'THIS SHOULDNT HAPPEN: parent collection died?
16:            Else
                   'Raise the event in the parent collection class instead of timer class
18:                objTimerCol.RaiseTimer_Event objTimer.Index
19:            End If
20:        Else
21:            objTimer.RaiseTimer_Event
22:        End If
23:    End If
    
25:    Set objTimer = Nothing
    
27:    Exit Sub
Err:
29:    HandleError Err.Number, Err.Description, Erl & "|" & "mTimers.TimerProc()"
End Sub

'Returns key to this class in collection
Public Function RegisterTimerCollection(ByVal colTimer As clsTimersCol) As Integer
1:    Dim sKey As String
2:    On Error GoTo Err
    
4:    m_TimersCount = m_TimersCount + 1
5:    sKey = "key:" & m_TimersCount 'will always be unique because counting
6:    g_colTimerS.Add colTimer, sKey
7:    RegisterTimerCollection = m_TimersCount
    
9:    Exit Function
Err:
11:   HandleError Err.Number, Err.Description, Erl & "|" & "mTimers.RegisterTimerCollection()"
End Function
