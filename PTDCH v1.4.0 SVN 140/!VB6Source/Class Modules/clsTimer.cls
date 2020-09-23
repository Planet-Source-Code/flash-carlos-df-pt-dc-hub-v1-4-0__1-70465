VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'This class can either be used independantly as single timer object, or it can be
'used through clsTimersCol which will allow for use as an index in an array of g_colTimer
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private mTimerID    As Long
Private mIsActive   As Boolean

Public Interval     As Long
Public Tag          As Variant

Public Event Timer()

'This block was added to base timer class to add support for using as
'part of a parent collection with events

Private mIndex As Long  ' these two added to support using as part
Private mParentsColKey  ' of collection with events

Private mIsNothing As Boolean

Friend Property Let Index(ByVal lIndex As Long)
1:    mIndex = lIndex
End Property

Friend Property Get Index() As Long
1:    Index = mIndex
End Property

Friend Property Let ParentsColKey(ByVal Key As Long)
1:    mParentsColKey = Key
End Property

Friend Property Get ParentsColKey() As Long
1:    ParentsColKey = mParentsColKey
End Property

Public Property Get Enabled() As Boolean
1:    Enabled = mIsActive
End Property

Public Property Let Enabled(ByVal bActivateIt As Boolean)
1:    On Error GoTo Err
    
3:    If bActivateIt And mIsActive Then
4:        Call Reset
5:    ElseIf bActivateIt And Not mIsActive Then
6:        Call Enable
7:    ElseIf mIsActive And Not bActivateIt Then
8:        Call Disable
9:    ElseIf Not mIsActive And Not bActivateIt Then
          'dumb ass
11:   End If
    
13:   Exit Property
14:
Err:
15:   HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.Enable(" & bActivateIt & ")"
End Property

Public Sub EnableTimer(ByVal lInterval As Long)
1:    On Error GoTo Err
    
3:    If lInterval = Interval And mIsActive Then
4:        TimerDestroy
5:    End If
6:    Interval = Interval
7:    Call Enable
8:    Exit Sub
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.EnableTimer(" & lInterval & ")"
End Sub

Public Sub Enable()
1:    On Error GoTo Err
2:    If mIsActive Then
3:        TimerDestroy
4:    End If
5:    Call TimerCreate
6:    Exit Sub
Err:
8:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.Enable()"
End Sub

Public Sub Disable()
1:    On Error GoTo Err
2:    Call TimerDestroy
3:    Exit Sub
Err:
5:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.Disable()"
End Sub

Public Sub Reset()
1:    On Error GoTo Err
2:    Call TimerDestroy
3:    Call TimerCreate
4:    Exit Sub
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.Reset()"
End Sub

Public Sub Disponse()
1:    Call TimerDestroy
2:    mIsNothing = True
End Sub

Public Sub RaiseTimer_Event()
1:    On Error GoTo Err
      'Must be public so that Timer object can't terminate while client's _Timer
      'event is being processed--Friend wouldn't prevent this disaster
4:    RaiseEvent Timer
5:    Exit Sub
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.RaiseTimer_Event()"
End Sub

Private Sub Class_Initialize()
1:    On Error GoTo Err
2:    Interval = 1000
3:    Exit Sub
Err:
5:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.Class_Initialize()"
End Sub

Private Sub Class_Terminate()
1:    On Error GoTo Err
2:    If Not mIsNothing Then
3:       Call TimerDestroy
4:    End If
5:    Exit Sub
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.Class_Initialize()"
End Sub

Private Function TimerCreate() As Boolean
1:     On Error GoTo Err
    
3:     If Interval <= 0 Then Exit Function
    
5:     mTimerID = SetTimer(0&, 0&, Interval, AddressOf TimerProc)
    
7:     If mTimerID Then
8:         TimerCreate = True
9:         g_colTimer.Add Me, "id:" & mTimerID
10:        mIsActive = True
11:    Else
12:        TimerCreate = False
13:        mTimerID = 0
14:        mIsActive = False
15:    End If
    
17:    Exit Function
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.TimerProc()"
End Function

Private Function TimerDestroy() As Long
1:    Dim i As Integer, f As Boolean
2:    On Error GoTo Err

4:    If TimerExists() Then
5:        f = KillTimer(0, mTimerID)
6:        g_colTimer.Remove "id:" & mTimerID
7:        TimerDestroy = True
8:        mIsActive = False
9:    End If

11:   Exit Function
Err:
13:   HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.TimerProc()"
End Function

Private Function TimerExists() As Boolean
1:    Dim objTimer As clsTimer
    
3:    On Error Resume Next
4:    Set objTimer = g_colTimer("id:" & mTimerID)

7:    If Err.Number = 0 Then
8:         TimerExists = True
9:    Else
10:        TimerExists = False
11:   End If
    
13:   Set objTimer = Nothing
    
15:   Exit Function
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "clsTimer.TimerProc()"
End Function
