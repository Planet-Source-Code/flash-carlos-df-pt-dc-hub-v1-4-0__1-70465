Attribute VB_Name = "mYSubclass"
Option Explicit

'************************************************************************
' SSubTmr object
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' Mod by fLaSh for PT DC Hub
'************************************************************************

' The implementation of the Subclassing part of the SSubTmr object.
' Use this module + clsYISubclass.Cls to replace dependency on the DLL.

' declares:
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_DESTROY = &H2

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private m_iCurrentMessage As Long
Private m_iProcOld As Long
Private m_f As Long

Public Property Get CurrentMessage() As Long
1:  CurrentMessage = m_iCurrentMessage
End Property

Private Sub ErrRaise(e As Long)
1:   On Error GoTo Err
   
3:   Dim sText As String, sSource As String
   
5:   If e > 1000 Then
6:        sSource = App.EXEName & ".WindowProc"
          Select Case e
            Case eeCantSubclass
7:               sText = "Can't subclass window"
            Case eeAlreadyAttached
8:               sText = "Message already handled by another class"
            Case eeInvalidWindow
9:               sText = "Invalid window"
            Case eeNoExternalWindow
10:               sText = "Can't modify external window"
11:          End Select
12:       Err.Raise e Or vbObjectError, sSource, sText
13:   Else
      ' Raise standard Visual Basic error
15:   Err.Raise e, sSource
16:   End If
   
18:   Exit Sub
19:
Err:
20:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.ErrRaise()"
End Sub

Private Property Get MessageCount(ByVal hWnd As Long) As Long
1:    On Error GoTo Err
    
3:    Dim sName As String
4:    sName = "C" & hWnd
5:    MessageCount = GetProp(hWnd, sName)
    
7:  Exit Property

9:
Err:
10: HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageCount()"
End Property

Private Property Let MessageCount(ByVal hWnd As Long, ByVal count As Long)
1:    On Error GoTo Err
    
3:    Dim sName As String
4:    m_f = 1
5:    sName = "C" & hWnd
6:    m_f = SetProp(hWnd, sName, count)
7:    If (count = 0) Then
8:      RemoveProp hWnd, sName
9:    End If
'   logMessage "Changed message count for " & Hex(hwnd) & " to " & count

12:  Exit Property

14:
Err:
15:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageCount()"
End Property

Private Property Get OldWindowProc(ByVal hWnd As Long) As Long
1:    On Error GoTo Err
    
3:    Dim sName As String
4:    sName = hWnd
5:    OldWindowProc = GetProp(hWnd, sName)

7:  Exit Property

9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.OldWindowProc()"
End Property
    
Private Property Let OldWindowProc(ByVal hWnd As Long, ByVal lPtr As Long)
1:   On Error GoTo Err
   
3:   Dim sName As String
4:   m_f = 1
5:   sName = hWnd
6:   m_f = SetProp(hWnd, sName, lPtr)
7:   If (lPtr = 0) Then
8:      RemoveProp hWnd, sName
9:   End If
'   logMessage "Changed Window Proc for " & Hex(hwnd) & " to " & Hex(lPtr)

12: Exit Property

14:
Err:
15:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.OldWindowProc()"
End Property

Private Property Get MessageClassCount(ByVal hWnd As Long, ByVal iMsg As Long) As Long
1:    On Error GoTo Err
    
3:    Dim sName As String
4:    sName = hWnd & "#" & iMsg & "C"
5:    MessageClassCount = GetProp(hWnd, sName)
    
7:    Exit Property

9:
Err:
10:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClassCount()"
End Property

Private Property Let MessageClassCount(ByVal hWnd As Long, ByVal iMsg As Long, ByVal count As Long)
1:    On Error GoTo Err
    
3:    Dim sName As String
4:    sName = hWnd & "#" & iMsg & "C"
5:    m_f = SetProp(hWnd, sName, count)
6:    If (count = 0) Then
7:       RemoveProp hWnd, sName
8:    End If
'   logMessage "Changed message count for " & Hex(hwnd) & " Message " & iMsg & " to " & count
    
11:    Exit Property

13:
Err:
14:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClassCount()"
End Property

Private Property Get MessageClass(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Index As Long) As Long
1:   On Error GoTo Err
    
3:   Dim sName As String
4:   sName = hWnd & "#" & iMsg & "#" & Index
5:   MessageClass = GetProp(hWnd, sName)
   
7:   Exit Property

9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClass()"
End Property
    
Private Property Let MessageClass(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Index As Long, ByVal classPtr As Long)
1:   On Error GoTo Err
   
3:   Dim sName As String
4:   sName = hWnd & "#" & iMsg & "#" & Index
5:   m_f = SetProp(hWnd, sName, classPtr)
6:   If (classPtr = 0) Then
7:      RemoveProp hWnd, sName
8:   End If
'   logMessage "Changed message class for " & Hex(hwnd) & " Message " & iMsg & " Index " & index & " to " & Hex(classPtr)
   
11:  Exit Property

13:
Err:
14:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClass()"
End Property

Sub AttachMessage( _
      iwp As clsYISubclass, _
      ByVal hWnd As Long, _
      ByVal iMsg As Long _
   )

    Dim procOld As Long
    Dim msgCount As Long
    Dim msgClassCount As Long
    Dim msgClass As Long
    
    On Error GoTo Err

   ' --------------------------------------------------------------------
   ' 1) Validate window
   ' --------------------------------------------------------------------
   If IsWindow(hWnd) = False Then
      ErrRaise eeInvalidWindow
      Exit Sub
   End If
   
   If IsWindowLocal(hWnd) = False Then
      ErrRaise eeNoExternalWindow
      Exit Sub
   End If

   ' --------------------------------------------------------------------
   ' 2) Check if this class is already attached for this message:
   ' --------------------------------------------------------------------
   msgClassCount = MessageClassCount(hWnd, iMsg)
   If (msgClassCount > 0) Then
      For msgClass = 1 To msgClassCount
         If (MessageClass(hWnd, iMsg, msgClass) = ObjPtr(iwp)) Then
            ErrRaise eeAlreadyAttached
            Exit Sub
         End If
      Next msgClass
   End If

   ' --------------------------------------------------------------------
   ' 3) Associate this class with this message for this window:
   ' --------------------------------------------------------------------
   MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) + 1
   If (m_f = 0) Then
      ' Failed, out of memory:
      ErrRaise 5
      Exit Sub
   End If
   
   ' --------------------------------------------------------------------
   ' 4) Associate the class pointer:
   ' --------------------------------------------------------------------
   MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = ObjPtr(iwp)
   If (m_f = 0) Then
      ' Failed, out of memory:
      MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
      ErrRaise 5
      Exit Sub
   End If

   ' --------------------------------------------------------------------
   ' 5) Get the message count
   ' --------------------------------------------------------------------
   msgCount = MessageCount(hWnd)
   If msgCount = 0 Then
      
      ' Subclass window by installing window procedure
      procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
      If procOld = 0 Then
         ' remove class:
         MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
         ' remove class count:
         MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
         
         ErrRaise eeCantSubclass
         Exit Sub
      End If
      
      ' Associate old procedure with handle
      OldWindowProc(hWnd) = procOld
      If m_f = 0 Then
         ' SPM: Failed to VBSetProp, windows properties database problem.
         ' Has to be out of memory.
         
         ' Put the old window proc back again:
         SetWindowLong hWnd, GWL_WNDPROC, procOld
         ' remove class:
         MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
         ' remove class count:
         MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
         
         ' Raise an error:
         ErrRaise 5
         Exit Sub
      End If
   End If
   
      
   ' Count this message
   MessageCount(hWnd) = MessageCount(hWnd) + 1
   If m_f = 0 Then
      ' SPM: Failed to set prop, windows properties database problem.
      ' Has to be out of memory
      
      ' remove class:
      MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
      ' remove class count contribution:
      MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
      
      ' If we haven't any messages on this window then remove the subclass:
      If (MessageCount(hWnd) = 0) Then
         ' put old window proc back again:
         procOld = OldWindowProc(hWnd)
         If Not (procOld = 0) Then
            SetWindowLong hWnd, GWL_WNDPROC, procOld
            OldWindowProc(hWnd) = 0
         End If
      End If
      
      ' Raise the error:
      ErrRaise 5
      Exit Sub
   End If
       
   Exit Sub
Err:
   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.AttachMessage()"
End Sub

Sub DetachMessage( _
      iwp As clsYISubclass, _
      ByVal hWnd As Long, _
      ByVal iMsg As Long _
   )
    Dim msgClassCount As Long
    Dim msgClass As Long
    Dim msgClassIndex As Long
    Dim msgCount As Long
    Dim procOld As Long
    
    On Error GoTo Err
    
   ' --------------------------------------------------------------------
   ' 1) Validate window
   ' --------------------------------------------------------------------
   If IsWindow(hWnd) = False Then
      ' for compatibility with the old version, we don't
      ' raise a message:
      ' ErrRaise eeInvalidWindow
      Exit Sub
   End If
   
   If IsWindowLocal(hWnd) = False Then
      ' for compatibility with the old version, we don't
      ' raise a message:
      ' ErrRaise eeNoExternalWindow
      Exit Sub
   End If
    
   ' --------------------------------------------------------------------
   ' 2) Check if this message is attached for this class:
   ' --------------------------------------------------------------------
   msgClassCount = MessageClassCount(hWnd, iMsg)
   If (msgClassCount > 0) Then
      msgClassIndex = 0
      For msgClass = 1 To msgClassCount
         If (MessageClass(hWnd, iMsg, msgClass) = ObjPtr(iwp)) Then
            msgClassIndex = msgClass
            Exit For
         End If
      Next msgClass
      
      If (msgClassIndex = 0) Then
         ' fail silently
         Exit Sub
      Else
         ' remove this message class:
         
         ' a) Anything above this index has to be shifted up:
         For msgClass = msgClassIndex To msgClassCount - 1
            MessageClass(hWnd, iMsg, msgClass) = MessageClass(hWnd, iMsg, msgClass + 1)
         Next msgClass
         
         ' b) The message class at the end can be removed:
         MessageClass(hWnd, iMsg, msgClassCount) = 0
         
         ' c) Reduce the message class count:
         MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
         
      End If
      
   Else
       ' fail silently
      Exit Sub
   End If
   
   ' ---------------------------------------------------------------------
   ' 3) Reduce the message count:
   ' ---------------------------------------------------------------------
   msgCount = MessageCount(hWnd)
   If (msgCount = 1) Then
      ' remove the subclass:
      procOld = OldWindowProc(hWnd)
      If Not (procOld = 0) Then
         ' Unsubclass by reassigning old window procedure
         Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
      End If
      ' remove the old window proc:
      OldWindowProc(hWnd) = 0
   End If
   MessageCount(hWnd) = MessageCount(hWnd) - 1
   
   Exit Sub
Err:
   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.DetachMessage()"
End Sub

Private Function WindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
    
7:    Dim procOld As Long
8:    Dim msgClassCount As Long
9:    Dim bCalled As Boolean
10:    Dim pSubClass As Long
11:    Dim iwp As clsYISubclass
12:    Dim iwpT As clsYISubclass
13:    Dim iIndex As Long
14:    Dim bDestroy As Boolean
    
16:    On Error GoTo Err
    
   ' Get the old procedure from the window
19:  procOld = OldWindowProc(hWnd)
20:   Debug.Assert procOld <> 0
    
22:   If (procOld = 0) Then
      ' we can't work, we're not subclassed properly.
24:      Exit Function
25:   End If
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
     
    ' Get the number of instances for this msg/hwnd:
35:   bCalled = False
   
37:   If (MessageClassCount(hWnd, iMsg) > 0) Then
38:      iIndex = MessageClassCount(hWnd, iMsg)
      
40:      Do While (iIndex >= 1)
41:         pSubClass = MessageClass(hWnd, iMsg, iIndex)
         
43:         If (pSubClass = 0) Then
               ' Not handled by this instance
45:         Else
               ' Turn pointer into a reference:
47:            CopyMemory iwpT, pSubClass, 4
48:            Set iwp = iwpT
49:            CopyMemory iwpT, 0&, 4
            
               ' Store the current message, so the client can check it:
52:            m_iCurrentMessage = iMsg
            
54:            With iwp
                  ' Preprocess (only checked first time around):
56:               If (iIndex = 1) Then
57:                  If (.MsgResponse = emrPreprocess) Then
58:                     If Not (bCalled) Then
59:                        WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                                  wParam, ByVal lParam)
61:                        bCalled = True
62:                     End If
63:                  End If
64:               End If
                 ' Consume (this message is always passed to all control
                 ' instances regardless of whether any single one of them
                 ' requests to consume it):
68:               WindowProc = .WindowProc(hWnd, iMsg, wParam, ByVal lParam)
69:            End With
70:         End If
         
72:         iIndex = iIndex - 1
73:      Loop
      
         ' PostProcess (only check this the last time around):
76:      If Not (iwp Is Nothing) And Not (procOld = 0) Then
77:          If iwp.MsgResponse = emrPostProcess Then
78:             If Not (bCalled) Then
79:                WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                          wParam, ByVal lParam)
81:                bCalled = True
82:             End If
83:          End If
84:      End If
            
86:   Else
         ' Not handled:
88:      If (iMsg = WM_DESTROY) Then
           ' If WM_DESTROY isn't handled already, we should
           ' clear up any subclass
91:         pClearUp hWnd
92:         WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
         
95:      Else
96:         WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
98:      End If
99:   End If
    
101:   Exit Function
102:
Err:
103:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.WindowProc()"
End Function
 
Public Function CallOldWindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
   
7:    On Error GoTo Err
    
9:    Dim iProcOld As Long
10:    iProcOld = OldWindowProc(hWnd)
    
12:    If Not (iProcOld = 0) Then
13:      CallOldWindowProc = CallWindowProc(iProcOld, hWnd, iMsg, wParam, lParam)
14:    End If
    
16:    Exit Function

18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.CallOldWindowProc()"
End Function

Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    On Error GoTo Err
    
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
    
    Exit Function
Err:
    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.IsWindowLocal()"
End Function

Private Sub logMessage(ByVal sMsg As String)
   'Debug.Print sMsg
End Sub

Private Sub pClearUp(ByVal hWnd As Long)
1:    Dim msgCount As Long
2:    Dim procOld As Long
    
4:    On Error GoTo Err
     
     ' this is only called if you haven't explicitly cleared up
     ' your subclass from the caller.  You will get a minor
     ' resource leak as it does not clear up any message
     ' specific properties.
10:   msgCount = MessageCount(hWnd)
11:   If (msgCount > 0) Then
         ' remove the subclass:
13:      procOld = OldWindowProc(hWnd)
14:      If Not (procOld = 0) Then
            ' Unsubclass by reassigning old window procedure
16:         Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
17:      End If
         ' remove the old window proc:
19:      OldWindowProc(hWnd) = 0
20:      MessageCount(hWnd) = 0
21:   End If
   
23:   Exit Sub

25:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.pClearUp()"
End Sub
